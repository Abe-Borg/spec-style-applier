import types

import pytest

from core.llm_classifier import classify_target_document, _merge_chunk_results


class _FakeStream:
    def __init__(self, payload):
        self.payload = payload

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def get_final_text(self):
        return self.payload


class _FakeMessages:
    def __init__(self, payload='{"classifications": []}'):
        self.last_kwargs = None
        self.payload = payload

    def stream(self, **kwargs):
        self.last_kwargs = kwargs
        return _FakeStream(self.payload)


class _FakeClient:
    def __init__(self):
        self.messages = _FakeMessages()


def test_output_config_is_dict(monkeypatch):
    fake = _FakeClient()

    fake_anthropic = types.SimpleNamespace(Anthropic=lambda api_key: fake)
    monkeypatch.setitem(__import__("sys").modules, "anthropic", fake_anthropic)

    bundle = {
        "paragraphs": [],
        "available_roles": ["PART"],
        "deterministic_classifications": [],
    }
    result = classify_target_document(bundle, ["PART"], api_key="x", model="m")
    assert fake.messages.last_kwargs is None
    assert result["notes"] == ["LLM skipped: all paragraphs classified deterministically."]


def test_classify_calls_llm_for_unresolved(monkeypatch):
    fake = _FakeClient()
    fake.messages.payload = '{"classifications": [{"paragraph_index": 3, "csi_role": "PART"}]}'
    fake_anthropic = types.SimpleNamespace(Anthropic=lambda api_key: fake)
    monkeypatch.setitem(__import__("sys").modules, "anthropic", fake_anthropic)

    bundle = {
        "paragraphs": [{"paragraph_index": 3, "text": "A"}],
        "available_roles": ["PART"],
        "deterministic_classifications": [],
    }
    classify_target_document(bundle, ["PART"], api_key="x", model="m")
    assert fake.messages.last_kwargs["output_config"] == {"effort": "high"}


def test_merge_chunk_results_conflict_raises():
    with pytest.raises(ValueError, match="conflicts"):
        _merge_chunk_results([
            {"classifications": [{"paragraph_index": 4, "csi_role": "PART"}], "notes": []},
            {"classifications": [{"paragraph_index": 4, "csi_role": "ARTICLE"}], "notes": []},
        ])
