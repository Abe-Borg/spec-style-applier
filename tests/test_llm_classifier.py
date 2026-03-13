"""Tests for core.llm_classifier — request construction, response parsing, and validation."""

import json
import pytest
from core.llm_classifier import (
    _build_api_kwargs,
    _build_user_message,
    _parse_classification_response,
    _validate_classifications,
    _split_bundle_into_chunks,
    _merge_chunk_results,
    validate_and_repair_classifications,
    merge_classifications,
)


# ---------------------------------------------------------------------------
# P2-001: Request construction
# ---------------------------------------------------------------------------

class TestBuildApiKwargs:
    """_build_api_kwargs produces a valid kwargs dict for messages.stream()."""

    def test_required_keys_present(self):
        kwargs = _build_api_kwargs(
            model="claude-sonnet-4-20250514",
            system_prompt="You are a classifier.",
            user_message="Classify this.",
        )
        assert "model" in kwargs
        assert "max_tokens" in kwargs
        assert "temperature" in kwargs
        assert "system" in kwargs
        assert "messages" in kwargs

    def test_model_passed_through(self):
        kwargs = _build_api_kwargs(
            model="claude-opus-4-6",
            system_prompt="sys",
            user_message="usr",
        )
        assert kwargs["model"] == "claude-opus-4-6"

    def test_thinking_is_dict(self):
        kwargs = _build_api_kwargs(
            model="m", system_prompt="s", user_message="u",
        )
        assert isinstance(kwargs["thinking"], dict)
        assert kwargs["thinking"]["type"] == "adaptive"

    def test_output_config_is_dict(self):
        kwargs = _build_api_kwargs(
            model="m", system_prompt="s", user_message="u",
        )
        assert isinstance(kwargs["output_config"], dict)
        assert kwargs["output_config"]["effort"] == "high"

    def test_no_unexpected_kwargs(self):
        """Only well-known Anthropic SDK params should be present."""
        kwargs = _build_api_kwargs(
            model="m", system_prompt="s", user_message="u",
        )
        known = {
            "model", "max_tokens", "temperature", "thinking",
            "output_config", "system", "messages",
        }
        assert set(kwargs.keys()) <= known

    def test_messages_shape(self):
        kwargs = _build_api_kwargs(
            model="m", system_prompt="s", user_message="hello",
        )
        msgs = kwargs["messages"]
        assert isinstance(msgs, list)
        assert len(msgs) == 1
        assert msgs[0]["role"] == "user"
        assert msgs[0]["content"] == "hello"


class TestBuildUserMessage:
    def test_includes_roles(self):
        msg = _build_user_message({"paragraphs": []}, ["PART", "ARTICLE"])
        assert '"PART"' in msg
        assert '"ARTICLE"' in msg

    def test_includes_bundle(self):
        bundle = {"paragraphs": [{"paragraph_index": 0, "text": "Hello"}]}
        msg = _build_user_message(bundle, ["PART"])
        assert '"Hello"' in msg


# ---------------------------------------------------------------------------
# P2-001: Response parsing
# ---------------------------------------------------------------------------

class TestParseClassificationResponse:
    def test_plain_json(self):
        text = '{"classifications": [], "notes": []}'
        result = _parse_classification_response(text)
        assert result["classifications"] == []

    def test_json_with_markdown_code_block(self):
        text = '```json\n{"classifications": [{"paragraph_index": 0, "csi_role": "PART"}]}\n```'
        result = _parse_classification_response(text)
        assert len(result["classifications"]) == 1

    def test_json_with_generic_code_block(self):
        text = '```\n{"classifications": []}\n```'
        result = _parse_classification_response(text)
        assert result["classifications"] == []

    def test_invalid_json_raises(self):
        with pytest.raises(json.JSONDecodeError):
            _parse_classification_response("not json at all")

    def test_whitespace_handled(self):
        text = '  \n {"classifications": []} \n '
        result = _parse_classification_response(text)
        assert result["classifications"] == []


# ---------------------------------------------------------------------------
# P2-001: Validation
# ---------------------------------------------------------------------------

class TestValidateClassifications:
    ROLES = ["PART", "ARTICLE", "PARAGRAPH", "SUBPARAGRAPH"]

    def test_valid_input_passes(self):
        data = {
            "classifications": [
                {"paragraph_index": 0, "csi_role": "PART"},
                {"paragraph_index": 5, "csi_role": "ARTICLE"},
            ]
        }
        result = _validate_classifications(data, self.ROLES)
        assert len(result["classifications"]) == 2

    def test_unknown_role_filtered(self):
        data = {
            "classifications": [
                {"paragraph_index": 0, "csi_role": "INVENTED_ROLE"},
            ]
        }
        result = _validate_classifications(data, self.ROLES)
        assert len(result["classifications"]) == 0

    def test_non_dict_items_filtered(self):
        data = {
            "classifications": [
                "not a dict",
                {"paragraph_index": 0, "csi_role": "PART"},
            ]
        }
        result = _validate_classifications(data, self.ROLES)
        assert len(result["classifications"]) == 1

    def test_missing_fields_filtered(self):
        data = {
            "classifications": [
                {"paragraph_index": 0},  # missing csi_role
                {"csi_role": "PART"},  # missing paragraph_index
            ]
        }
        result = _validate_classifications(data, self.ROLES)
        assert len(result["classifications"]) == 0

    def test_non_dict_top_level_raises(self):
        with pytest.raises(ValueError):
            _validate_classifications("not a dict", self.ROLES)

    def test_missing_classifications_key_returns_empty(self):
        # Missing key defaults to empty list (current behavior)
        result = _validate_classifications({"items": []}, self.ROLES)
        assert len(result["classifications"]) == 0

    def test_notes_preserved(self):
        data = {
            "classifications": [],
            "notes": ["some note"],
        }
        result = _validate_classifications(data, self.ROLES)
        assert result["notes"] == ["some note"]

    def test_negative_index_filtered_with_total(self):
        data = {
            "classifications": [
                {"paragraph_index": -1, "csi_role": "PART"},
            ]
        }
        result = _validate_classifications(data, self.ROLES, total_paragraphs=10)
        assert len(result["classifications"]) == 0

    def test_out_of_range_index_filtered(self):
        data = {
            "classifications": [
                {"paragraph_index": 99, "csi_role": "PART"},
            ]
        }
        result = _validate_classifications(data, self.ROLES, total_paragraphs=10)
        assert len(result["classifications"]) == 0

    def test_in_range_index_passes(self):
        data = {
            "classifications": [
                {"paragraph_index": 9, "csi_role": "PART"},
            ]
        }
        result = _validate_classifications(data, self.ROLES, total_paragraphs=10)
        assert len(result["classifications"]) == 1

    def test_duplicate_index_keeps_last(self):
        data = {
            "classifications": [
                {"paragraph_index": 0, "csi_role": "PART"},
                {"paragraph_index": 0, "csi_role": "ARTICLE"},
            ]
        }
        result = _validate_classifications(data, self.ROLES)
        assert len(result["classifications"]) == 1
        assert result["classifications"][0]["csi_role"] == "ARTICLE"

    def test_results_sorted_by_index(self):
        data = {
            "classifications": [
                {"paragraph_index": 5, "csi_role": "ARTICLE"},
                {"paragraph_index": 1, "csi_role": "PART"},
                {"paragraph_index": 3, "csi_role": "PARAGRAPH"},
            ]
        }
        result = _validate_classifications(data, self.ROLES)
        indices = [c["paragraph_index"] for c in result["classifications"]]
        assert indices == [1, 3, 5]

    def test_no_total_paragraphs_allows_any_nonnegative(self):
        """Without total_paragraphs, any non-negative int index is accepted."""
        data = {
            "classifications": [
                {"paragraph_index": 999999, "csi_role": "PART"},
            ]
        }
        result = _validate_classifications(data, self.ROLES)
        assert len(result["classifications"]) == 1


# ---------------------------------------------------------------------------
# Chunking and merging
# ---------------------------------------------------------------------------

class TestSplitBundleIntoChunks:
    def test_small_bundle_not_split(self):
        bundle = {
            "paragraphs": [{"paragraph_index": i, "text": f"p{i}"} for i in range(10)],
            "document_meta": {},
            "available_roles": ["PART"],
            "filter_report": {},
        }
        chunks = _split_bundle_into_chunks(bundle)
        assert len(chunks) == 1

    def test_large_bundle_split(self):
        # Create a bundle large enough to split (>300 paragraphs AND large JSON)
        paragraphs = [
            {"paragraph_index": i, "text": "x" * 1000} for i in range(350)
        ]
        bundle = {
            "paragraphs": paragraphs,
            "document_meta": {},
            "available_roles": ["PART"],
            "filter_report": {},
        }
        chunks = _split_bundle_into_chunks(bundle)
        assert len(chunks) > 1


class TestMergeChunkResults:
    def test_single_chunk(self):
        result = _merge_chunk_results([
            {"classifications": [{"paragraph_index": 0, "csi_role": "PART"}], "notes": []}
        ])
        assert len(result["classifications"]) == 1

    def test_dedup_by_paragraph_index(self):
        result = _merge_chunk_results([
            {"classifications": [{"paragraph_index": 0, "csi_role": "PART"}], "notes": []},
            {"classifications": [{"paragraph_index": 0, "csi_role": "ARTICLE"}], "notes": []},
        ])
        # Later chunk overrides
        assert len(result["classifications"]) == 1
        assert result["classifications"][0]["csi_role"] == "ARTICLE"

    def test_notes_merged(self):
        result = _merge_chunk_results([
            {"classifications": [], "notes": ["note1"]},
            {"classifications": [], "notes": ["note2"]},
        ])
        assert result["notes"] == ["note1", "note2"]


# ---------------------------------------------------------------------------
# P2-006: Validate and repair
# ---------------------------------------------------------------------------

class TestValidateAndRepair:
    ROLES = ["PART", "ARTICLE", "PARAGRAPH", "SUBPARAGRAPH"]

    def _bundle(self, texts):
        return {
            "paragraphs": [{"paragraph_index": i, "text": t} for i, t in enumerate(texts)]
        }

    def test_full_coverage_passes(self):
        llm = {"classifications": [
            {"paragraph_index": 0, "csi_role": "PART"},
            {"paragraph_index": 1, "csi_role": "ARTICLE"},
        ]}
        result = validate_and_repair_classifications(
            llm, self.ROLES, {0, 1}, self._bundle(["PART 1", "1.01 X"])
        )
        assert len(result["classifications"]) == 2

    def test_repair_via_marker(self):
        """Missing index with an obvious marker gets repaired."""
        llm = {"classifications": [
            {"paragraph_index": 0, "csi_role": "PART"},
        ]}
        result = validate_and_repair_classifications(
            llm, self.ROLES, {0, 1}, self._bundle(["PART 1", "1.01 SUMMARY"])
        )
        assert len(result["classifications"]) == 2
        roles = {c["paragraph_index"]: c["csi_role"] for c in result["classifications"]}
        assert roles[1] == "ARTICLE"

    def test_low_coverage_raises(self):
        """Coverage below 95% should raise ValueError."""
        # 20 ambiguous paragraphs, LLM only classifies 1
        texts = [f"ambiguous text {i}" for i in range(20)]
        llm = {"classifications": [
            {"paragraph_index": 0, "csi_role": "PARAGRAPH"},
        ]}
        with pytest.raises(ValueError, match="coverage too low"):
            validate_and_repair_classifications(
                llm, self.ROLES, set(range(20)), {"paragraphs": [
                    {"paragraph_index": i, "text": t} for i, t in enumerate(texts)
                ]}
            )


class TestMergeClassifications:
    def test_basic_merge(self):
        preclassified = {0: "PART", 2: "ARTICLE"}
        llm = {"classifications": [
            {"paragraph_index": 1, "csi_role": "PARAGRAPH"},
        ], "notes": []}
        result = merge_classifications(preclassified, llm, total_paragraphs=10)
        assert len(result["classifications"]) == 3
        roles = {c["paragraph_index"]: c["csi_role"] for c in result["classifications"]}
        assert roles == {0: "PART", 1: "PARAGRAPH", 2: "ARTICLE"}

    def test_llm_overrides_preclassified(self):
        preclassified = {0: "PART"}
        llm = {"classifications": [
            {"paragraph_index": 0, "csi_role": "ARTICLE"},
        ], "notes": []}
        result = merge_classifications(preclassified, llm, total_paragraphs=10)
        roles = {c["paragraph_index"]: c["csi_role"] for c in result["classifications"]}
        assert roles[0] == "ARTICLE"

    def test_out_of_range_excluded(self):
        preclassified = {99: "PART"}
        llm = {"classifications": [], "notes": []}
        result = merge_classifications(preclassified, llm, total_paragraphs=10)
        assert len(result["classifications"]) == 0
