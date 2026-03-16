from core.batch_classifier import build_batch_requests, reassemble_file_classifications


def test_build_batch_requests_creates_chunked_custom_ids(monkeypatch):
    from core import batch_classifier as bc

    monkeypatch.setattr(
        bc,
        "_split_bundle_into_chunks",
        lambda bundle: [{"paragraphs": [{"paragraph_index": 1}]}, {"paragraphs": [{"paragraph_index": 2}]}],
    )

    reqs = build_batch_requests(
        file_bundles={"my-file.docx": {"paragraphs": []}},
        available_roles=["PART"],
        model="m",
    )

    assert [r["custom_id"] for r in reqs] == ["my-file__chunk0", "my-file__chunk1"]
    assert reqs[0]["params"]["model"] == "m"
    assert reqs[0]["params"]["output_config"] == {"effort": "high"}


def test_reassemble_file_classifications_merges_chunks(monkeypatch):
    from core import batch_classifier as bc

    split_chunks = [
        {"paragraphs": [{"paragraph_index": 1}]},
        {"paragraphs": [{"paragraph_index": 2}]},
    ]
    monkeypatch.setattr(bc, "_split_bundle_into_chunks", lambda bundle: split_chunks)

    out = reassemble_file_classifications(
        results={
            "a__chunk0": {"classifications": [{"paragraph_index": 1, "csi_role": "PART"}]},
            "a__chunk1": {"classifications": [{"paragraph_index": 2, "csi_role": "PART"}]},
        },
        file_bundles={
            "a.docx": {
                "paragraphs": [{"paragraph_index": 1}, {"paragraph_index": 2}],
                "deterministic_classifications": [],
                "available_roles": ["PART"],
            }
        },
        available_roles=["PART"],
    )

    assert "a.docx" in out
    assert len(out["a.docx"]["classifications"]) == 2
