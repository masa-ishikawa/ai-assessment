"""公開Web調査で使用する、サイズ制限付きHTTPS取得プリミティブ。"""

from __future__ import annotations

from dataclasses import dataclass
import ipaddress
import socket
from typing import Callable, Iterable, Mapping
from urllib.parse import urlparse


DEFAULT_MAX_DOCUMENT_BYTES = 25 * 1024 * 1024
DEFAULT_MAX_SEARCH_BYTES = 2 * 1024 * 1024
DOCUMENT_MIME_TYPES = frozenset({
    "application/pdf",
    "application/octet-stream",
    "application/xhtml+xml",
    "text/html",
    "text/plain",
})
HTML_MIME_TYPES = frozenset({"application/xhtml+xml", "text/html", "text/plain"})


class SafeFetchError(RuntimeError):
    """安全条件を満たさないURLまたはレスポンスを表す。"""


@dataclass(frozen=True, slots=True)
class FetchedResponse:
    """本文を上限付きで読み込んだ、既存抽出処理と互換なレスポンス。"""

    url: str
    content: bytes
    content_type: str
    status_code: int
    headers: Mapping[str, str]

    @property
    def text(self) -> str:
        content_type = self.content_type.casefold()
        charset = "utf-8"
        if "charset=" in content_type:
            charset = content_type.split("charset=", 1)[1].split(";", 1)[0].strip() or charset
        try:
            return self.content.decode(charset, errors="replace")
        except LookupError:
            return self.content.decode("utf-8", errors="replace")


def _is_global_address(value: str) -> bool:
    try:
        return ipaddress.ip_address(value).is_global
    except ValueError:
        return False


def public_https_url(
    value: object,
    *,
    resolve_dns: bool = False,
    resolver: Callable[..., Iterable[tuple]] = socket.getaddrinfo,
) -> bool:
    """HTTPS URLを構文検証し、取得直前にはDNS解決先も公開IPに限定する。"""
    try:
        parsed = urlparse(str(value or ""))
        port = parsed.port
    except ValueError:
        return False
    hostname = (parsed.hostname or "").rstrip(".").casefold()
    if (
        parsed.scheme != "https"
        or not hostname
        or parsed.username
        or parsed.password
        or (port is not None and port != 443)
        or hostname == "localhost"
        or hostname.endswith(".localhost")
        or hostname.endswith(".local")
    ):
        return False
    try:
        address = ipaddress.ip_address(hostname)
    except ValueError:
        if "." not in hostname:
            return False
        if not resolve_dns:
            return True
        try:
            answers = list(resolver(hostname, 443, type=socket.SOCK_STREAM))
        except (OSError, socket.gaierror):
            return False
        addresses = {
            str(answer[4][0]).split("%", 1)[0]
            for answer in answers
            if len(answer) >= 5 and answer[4]
        }
        # 1件でも内部・予約アドレスを含むホストは拒否する。DNS rebindingや
        # split-horizon DNSで内部サービスへ到達する経路を残さない。
        return bool(addresses) and all(_is_global_address(item) for item in addresses)
    return address.is_global


def _normalized_mime(headers: Mapping[str, str]) -> str:
    return str(headers.get("content-type", "")).split(";", 1)[0].strip().casefold()


def safe_fetch_public_https(
    client: object,
    url: str,
    *,
    max_bytes: int = DEFAULT_MAX_DOCUMENT_BYTES,
    allowed_mime_types: frozenset[str] = DOCUMENT_MIME_TYPES,
    params: Mapping[str, str] | None = None,
    resolver: Callable[..., Iterable[tuple]] = socket.getaddrinfo,
) -> FetchedResponse:
    """公開HTTPSだけを取得し、MIME・Content-Length・実読込量を制限する。"""
    if max_bytes <= 0:
        raise ValueError("max_bytesは正数で指定してください。")
    if not public_https_url(url, resolve_dns=True, resolver=resolver):
        raise SafeFetchError("公開HTTPS URLとして検証できませんでした。")

    # httpx.Client.stream互換の最小契約だけを要求し、呼び出し側でClientを共有できるようにする。
    with client.stream("GET", url, params=params) as response:  # type: ignore[attr-defined]
        response.raise_for_status()
        final_url = str(response.url)
        if not public_https_url(final_url, resolve_dns=True, resolver=resolver):
            raise SafeFetchError("取得先URLのDNS解決結果が公開IPではありません。")
        headers = {str(key).casefold(): str(value) for key, value in response.headers.items()}
        content_length = headers.get("content-length", "").strip()
        if content_length:
            try:
                if int(content_length) > max_bytes:
                    raise SafeFetchError("Web資料が許容サイズを超えています。")
            except ValueError:
                raise SafeFetchError("不正なContent-Lengthを受信しました。") from None
        mime = _normalized_mime(headers)
        if mime and mime not in allowed_mime_types:
            raise SafeFetchError(f"許可されていないContent-Typeです: {mime}")

        body = bytearray()
        for chunk in response.iter_bytes():
            body.extend(chunk)
            if len(body) > max_bytes:
                raise SafeFetchError("Web資料が許容サイズを超えています。")

    content = bytes(body)
    sniffed_mime = mime
    if not sniffed_mime:
        if content.startswith(b"%PDF"):
            sniffed_mime = "application/pdf"
        elif content.lstrip().startswith((b"<", b"<!")):
            sniffed_mime = "text/html"
        else:
            raise SafeFetchError("Content-Typeを確認できないWeb資料です。")
    if sniffed_mime == "application/octet-stream" and not content.startswith(b"%PDF"):
        raise SafeFetchError("application/octet-streamはPDF資料に限り許可されます。")
    return FetchedResponse(
        url=final_url,
        content=content,
        content_type=str(headers.get("content-type", sniffed_mime)),
        status_code=int(response.status_code),
        headers=headers,
    )
