from pathlib import Path

from batch_runner import BatchResult, run_batch_concurrent


def test_run_batch_concurrent_sorts_results_and_calls_callback(monkeypatch):
    completed = []

    def fake_process_single_file(
        docx_path,
        arch_registry,
        env_registry,
        arch_styles_xml,
        available_roles,
        api_key,
        output_dir,
        extract_base_dir=Path("output"),
        model="claude-opus-4-6",
    ):
        return BatchResult(
            filename=docx_path.name,
            success=True,
            output_path=output_dir / f"{docx_path.stem}_PHASE2_FORMATTED.docx",
            log=[f"Processed {docx_path.name}"],
            error=None,
            duration_seconds=0.1,
        )

    monkeypatch.setattr("batch_runner.process_single_file", fake_process_single_file)

    docx_paths = [Path("b.docx"), Path("a.docx"), Path("c.docx")]

    results = run_batch_concurrent(
        docx_paths=docx_paths,
        arch_registry={},
        env_registry={},
        arch_styles_xml="",
        available_roles=[],
        api_key="k",
        output_dir=Path("out"),
        max_workers=2,
        on_file_complete=lambda result: completed.append(result.filename),
    )

    assert sorted(completed) == ["a.docx", "b.docx", "c.docx"]
    assert [item.filename for item in results] == ["a.docx", "b.docx", "c.docx"]
