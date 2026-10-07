import copy
import unittest
from unittest.mock import patch

from ai_assess_runtime import quantitative_contract as contract


class QuantitativeDisplayContractTests(unittest.TestCase):
    @staticmethod
    def _research() -> dict:
        return {
            "industry_sources": [
                {
                    "id": f"R{index}",
                    "title": f"公開事例{index}",
                    "url": f"https://example.com/case{index}",
                    "excerpt": f"処理時間を{index * 10}%削減",
                }
                for index in range(1, 4)
            ],
            "quantitative_evidence": {
                "status": "sufficient",
                "approved_source_ids": ["R1", "R2", "R3"],
            },
        }

    @staticmethod
    def _evidence() -> dict:
        return {
            "evidence_mode": "external_verified",
            "benchmarks": [
                {
                    "claim_id": f"Q0{index}",
                    "claim_status": "verified_external",
                    "source_id": f"R{index}",
                    "source_url": f"https://example.com/case{index}",
                    "source_title": f"公開事例{index}",
                    "source_locator": "導入効果",
                    "source_metric": f"処理時間を{index * 10}%削減",
                    "evidence_excerpt": f"処理時間を{index * 10}%削減",
                }
                for index in range(1, 4)
            ],
        }

    def test_missing_evidence_uses_measurement_design_without_numeric_claims(self) -> None:
        assessment = {"company_name": "Example株式会社", "service_name": "Example Service"}
        research = {"industry_sources": []}
        assessment_before = copy.deepcopy(assessment)
        research_before = copy.deepcopy(research)

        result = contract.resolve_quantitative_display_contract(
            assessment, research, {}, missing_evidence=["同業務の公開効果が不足"],
        )

        self.assertEqual(contract.QuantitativeEvidenceStatus.INSUFFICIENT, result.evidence_status)
        self.assertEqual(
            contract.QuantitativeDisplayMode.MEASUREMENT_DESIGN_ONLY,
            result.display_mode,
        )
        self.assertFalse(result.numeric_claims_allowed)
        self.assertEqual("pre_poc", result.measurement_design["status"])
        self.assertEqual(4, len(result.measurement_design["items"]))
        self.assertEqual(("同業務の公開効果が不足",), result.missing_evidence)
        self.assertEqual(assessment_before, assessment)
        self.assertEqual(research_before, research)

    def test_contract_returns_defensive_measurement_design_copies(self) -> None:
        result = contract.resolve_quantitative_display_contract({}, {}, {})
        first = result.measurement_design
        first["items"][0]["kpi"] = "改ざん"

        self.assertNotEqual("改ざん", result.measurement_design["items"][0]["kpi"])
        self.assertFalse(result.as_dict()["numeric_claims_allowed"])

    def test_three_independently_verified_claims_allow_numeric_display(self) -> None:
        expected_design = {
            "schema_version": contract.POC_MEASUREMENT_DESIGN_SCHEMA_VERSION,
            "status": "external_verified",
            "items": [{"claim_id": f"Q0{index}"} for index in range(1, 4)],
        }
        evidence = self._evidence()
        research = self._research()
        with patch.object(
            contract, "_quantitative_claim_is_verifiable", return_value=True,
        ), patch.object(
            contract, "poc_measurement_design_from_verified_evidence",
            return_value=expected_design,
        ):
            result = contract.resolve_quantitative_display_contract(
                {"company_name": "Example株式会社"}, research, evidence,
            )

        self.assertEqual(contract.QuantitativeEvidenceStatus.VERIFIED, result.evidence_status)
        self.assertEqual(contract.QuantitativeDisplayMode.EXTERNAL_VERIFIED, result.display_mode)
        self.assertTrue(result.numeric_claims_allowed)
        self.assertEqual(("Q01", "Q02", "Q03"), result.approved_claim_ids)
        self.assertEqual(("R1", "R2", "R3"), result.approved_source_ids)
        self.assertEqual(expected_design, result.measurement_design)

    def test_three_claims_with_failed_binding_are_not_displayable(self) -> None:
        evidence = self._evidence()
        with patch.object(
            contract, "_quantitative_claim_is_verifiable", return_value=True,
        ), patch.object(
            contract, "poc_measurement_design_from_verified_evidence", return_value={},
        ):
            result = contract.resolve_quantitative_display_contract(
                {}, self._research(), evidence,
            )

        self.assertEqual(
            contract.QuantitativeEvidenceStatus.CONTRACT_INCOMPLETE,
            result.evidence_status,
        )
        self.assertFalse(result.numeric_claims_allowed)
        self.assertEqual("pre_poc", result.measurement_design["status"])

    def test_materialization_writes_display_and_audit_from_same_decision(self) -> None:
        assessment: dict = {}
        research: dict = {}
        decision = contract.resolve_quantitative_display_contract(
            assessment, research, {}, missing_evidence=["原典不足"],
        )

        outcome = contract.materialize_quantitative_display_contract(
            assessment, research, decision,
        )

        self.assertEqual("measurement_design_only", outcome["status"])
        self.assertEqual(assessment["poc_measurement_design"], research["poc_measurement_design"])
        self.assertEqual(
            "measurement_design_only",
            research["quantitative_display_contract"]["display_mode"],
        )
        self.assertFalse(research["quantitative_display_contract"]["numeric_claims_allowed"])
        self.assertEqual(
            "unsupported_numeric_claims_omitted",
            research["poc_measurement_design_audit"]["generation_source"],
        )
        self.assertTrue(research["quantitative_analysis"]["unsupported_claims_omitted"])

        payload = {"assessment": assessment, "research": research}
        self.assertEqual([], contract.validate_quantitative_evidence(payload))

    def test_validation_detects_display_contract_tampering(self) -> None:
        assessment: dict = {}
        research: dict = {}
        decision = contract.resolve_quantitative_display_contract(assessment, research, {})
        contract.materialize_quantitative_display_contract(assessment, research, decision)
        research["quantitative_display_contract"]["numeric_claims_allowed"] = True

        issues = contract.validate_quantitative_evidence({
            "assessment": assessment,
            "research": research,
        })

        self.assertTrue(any("external_verified" in issue for issue in issues))
        self.assertTrue(any("verified" in issue for issue in issues))


if __name__ == "__main__":
    unittest.main()
