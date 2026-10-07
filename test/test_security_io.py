"""入力・公開Web取得・CLI秘密値の安全境界を固定する。"""

import contextlib
import io
from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest
from unittest.mock import patch
import zipfile

from ai_assess_runtime.cli import CliDefaults, build_argument_parser
from ai_assess_runtime.http_safety import (
    HTML_MIME_TYPES,
    SafeFetchError,
    public_https_url,
    safe_fetch_public_https,
)
from ai_assess_runtime import source_input


def _resolver(addresses: list[str]):
    return lambda *args, **kwargs: [
        (2, 1, 6, "", (address, 443)) for address in addresses
    ]


class _FakeResponse:
    def __init__(self, *, body: bytes, content_type: str = "text/html",
                 content_length: str | None = None) -> None:
        self.url = "https://example.com/report"
        self.status_code = 200
        self.headers = {"content-type": content_type}
        if content_length is not None:
            self.headers["content-length"] = content_length
        self._body = body

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False

    def raise_for_status(self) -> None:
        return None

    def iter_bytes(self):
        yield from (self._body[:3], self._body[3:])


class _FakeClient:
    def __init__(self, response: _FakeResponse) -> None:
        self.response = response

    def stream(self, method: str, url: str, *, params=None):
        del method, url, params
        return self.response


class SecurityIoTests(unittest.TestCase):
    def test_public_https_dns_must_resolve_only_to_global_addresses(self) -> None:
        self.assertTrue(public_https_url(
            "https://example.com/report", resolve_dns=True,
            resolver=_resolver(["93.184.216.34"]),
        ))
        self.assertFalse(public_https_url(
            "https://example.com/report", resolve_dns=True,
            resolver=_resolver(["93.184.216.34", "10.0.0.8"]),
        ))
        self.assertFalse(public_https_url(
            "https://metadata.local/report", resolve_dns=True,
            resolver=_resolver(["169.254.169.254"]),
        ))

    def test_safe_fetch_streams_with_mime_and_byte_limit(self) -> None:
        response = _FakeResponse(body=b"<html>ok</html>")
        fetched = safe_fetch_public_https(
            _FakeClient(response), "https://example.com/report",
            max_bytes=32, allowed_mime_types=HTML_MIME_TYPES,
            resolver=_resolver(["93.184.216.34"]),
        )
        self.assertEqual(b"<html>ok</html>", fetched.content)
        self.assertIn("html", fetched.text)

        with self.assertRaisesRegex(SafeFetchError, "許容サイズ"):
            safe_fetch_public_https(
                _FakeClient(_FakeResponse(body=b"0123456789")),
                "https://example.com/report", max_bytes=5,
                allowed_mime_types=HTML_MIME_TYPES,
                resolver=_resolver(["93.184.216.34"]),
            )

    def test_safe_fetch_rejects_declared_oversize_and_unexpected_mime(self) -> None:
        with self.assertRaisesRegex(SafeFetchError, "許容サイズ"):
            safe_fetch_public_https(
                _FakeClient(_FakeResponse(body=b"x", content_length="100")),
                "https://example.com/report", max_bytes=10,
                allowed_mime_types=HTML_MIME_TYPES,
                resolver=_resolver(["93.184.216.34"]),
            )
        with self.assertRaisesRegex(SafeFetchError, "Content-Type"):
            safe_fetch_public_https(
                _FakeClient(_FakeResponse(body=b"GIF89a", content_type="image/gif")),
                "https://example.com/report", max_bytes=20,
                allowed_mime_types=HTML_MIME_TYPES,
                resolver=_resolver(["93.184.216.34"]),
            )

    def test_text_and_xlsx_expansion_limits_fail_before_parsing(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            directory = Path(temporary_directory)
            text_path = directory / "large.txt"
            text_path.write_bytes(b"12345")
            with patch.object(source_input, "MAX_TEXT_INPUT_BYTES", 4):
                with self.assertRaisesRegex(ValueError, "許容サイズ"):
                    source_input.load_source_text(text_path)

            xlsx_path = directory / "bomb.xlsx"
            with zipfile.ZipFile(xlsx_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                archive.writestr("xl/workbook.xml", b"123456")
            with zipfile.ZipFile(xlsx_path) as archive:
                with self.assertRaisesRegex(ValueError, "展開後サイズ"):
                    source_input._validate_xlsx_archive(
                        archive, max_members=10, max_member_bytes=10,
                        max_uncompressed_bytes=5, max_compression_ratio=1000,
                    )

    def test_cli_bounds_research_rounds_and_hides_deprecated_api_key(self) -> None:
        parser = build_argument_parser(CliDefaults(
            input_file=Path("input.xlsx"), provider="oci_responses",
            settings=SimpleNamespace(
                COMPARTMENT_ID="", GENAI_ENDPOINT="", OCI_PROFILE="DEFAULT",
                OCI_CONFIG_FILE="config", OCI_GENAI_PROJECT_OCID="", OCI_REGION="",
            ),
        ))
        self.assertNotIn("--openai-api-key", parser.format_help())
        for value in ("0", "9", "invalid"):
            with self.subTest(value=value), contextlib.redirect_stderr(io.StringIO()):
                with self.assertRaises(SystemExit):
                    parser.parse_args(["--research-max-rounds", value])
        warning = io.StringIO()
        with contextlib.redirect_stderr(warning):
            args = parser.parse_args(["--openai-api-key", "secret"])
        self.assertEqual("secret", args.openai_api_key)
        self.assertIn("OPENAI_API_KEY", warning.getvalue())

    def test_container_copies_runtime_and_requires_mounted_input(self) -> None:
        project = Path(__file__).resolve().parents[1]
        containerfile = (project / "Containerfile").read_text(encoding="utf-8")
        self.assertIn("COPY ai_assess_runtime ./ai_assess_runtime", containerfile)
        self.assertIn('VOLUME ["/work"]', containerfile)
        self.assertIn('CMD ["--help"]', containerfile)
        self.assertNotIn('CMD ["service_input.txt"]', containerfile)
        dockerignore = (project / ".dockerignore").read_text(encoding="utf-8")
        self.assertTrue(dockerignore.startswith("#"))
        self.assertIn("assessment_inputs/*", dockerignore)
        self.assertNotIn("assessment_config.py", [line for line in dockerignore.splitlines() if line.startswith("!")])


if __name__ == "__main__":
    unittest.main()
