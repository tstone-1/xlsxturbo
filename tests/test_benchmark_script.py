"""The benchmark script's data generator and output check.

The published tables on ``docs/performance.md`` come from ``benchmarks/benchmark.py``.
Two of its parts decide whether those numbers mean anything: the generator must
produce the column mix it claims, and the output check must refuse a workbook that
does not hold the whole frame, or a library that silently wrote less would look fast.
"""

from __future__ import annotations

import importlib.util
import json
import sys
from collections.abc import Callable
from pathlib import Path
from types import ModuleType
from typing import Any

import pandas as pd
import pytest
import xlsxturbo

from tests.helpers import REPO_ROOT, repo_checkout_available

pytestmark = pytest.mark.skipif(not repo_checkout_available(), reason="imports benchmarks/ from a source checkout")


@pytest.fixture(scope="module")
def bench() -> ModuleType:
    """Import ``benchmarks/benchmark.py``, which is a script, not a package.

    Returns:
        The module.
    """
    name = "xlsxturbo_benchmark_script"
    spec = importlib.util.spec_from_file_location(name, REPO_ROOT / "benchmarks" / "benchmark.py")
    assert spec is not None
    assert spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    # Registered first: @dataclass resolves the module through sys.modules.
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


@pytest.mark.parametrize(
    ("shape", "prefixes"),
    [("mixed", {"int", "float", "str", "date", "bool"}), ("numeric", {"int", "float"}), ("strings", {"str"})],
)
def test_each_shape_generates_its_column_types(bench: ModuleType, shape: str, prefixes: set[str]) -> None:
    """Column names carry the type, and each shape uses exactly its own types."""
    df = bench.generate_test_data(50, 16, shape=shape)
    assert df.shape == (50, 16)
    assert {name.split("_")[0] for name in df.columns} == prefixes


@pytest.mark.skipif(sys.platform == "win32", reason="--memory is unavailable on Windows")
def test_memory_failure_is_reported(
    bench: ModuleType, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str],
) -> None:
    """A failed optional memory run remains visible in JSON and fails the command."""
    monkeypatch.setattr(bench, "BENCHMARK_FUNCS", [("xlsxturbo", scripted_writer(bench, "xlsxturbo", [True] * 2))])
    monkeypatch.setattr(bench, "measure_peak_memory", lambda *_args: None)
    monkeypatch.setattr(
        bench.sys, "argv", ["benchmark", "--rows", "2", "--cols", "2", "--runs", "1", "--json", "--memory"],
    )
    assert bench.main() == 1
    result = json.loads(capsys.readouterr().out)["results"][0]
    assert result["successful_runs"] == 1
    assert result["peak_memory_mb"] is None
    assert result["errors"] == ["memory run failed; see stderr for details"]


def test_generation_is_seeded(bench: ModuleType) -> None:
    """Every library and every run writes the same data."""
    pd.testing.assert_frame_equal(bench.generate_test_data(30, 8), bench.generate_test_data(30, 8))


@pytest.mark.parametrize(("index", "letters"), [(1, "A"), (26, "Z"), (27, "AA"), (50, "AX"), (703, "AAA")])
def test_column_letters(bench: ModuleType, index: int, letters: str) -> None:
    """1-based column numbers map to Excel's letters."""
    assert bench.column_letters(index) == letters


def test_check_output_accepts_a_complete_workbook(bench: ModuleType, tmp_path: Path) -> None:
    """A workbook holding the frame and its header passes."""
    df = bench.generate_test_data(20, 30)
    target = tmp_path / "full.xlsx"
    xlsxturbo.df_to_xlsx(df, target)
    bench.check_output(str(target), df)


def test_check_output_refuses_missing_rows(bench: ModuleType, tmp_path: Path) -> None:
    """A workbook with fewer rows than the frame is refused, not timed."""
    df = bench.generate_test_data(20, 3)
    target = tmp_path / "short.xlsx"
    xlsxturbo.df_to_xlsx(df.head(19), target)
    with pytest.raises(RuntimeError, match="expected 'A1:C21'"):
        bench.check_output(str(target), df)


def test_check_output_refuses_missing_interior_cells(bench: ModuleType, tmp_path: Path) -> None:
    """A full bounding rectangle cannot conceal the cells pandas' streaming export loses."""
    frame = pd.DataFrame({"a": [1, 2, 3], "b": [4, 5, 6], "c": [7, 8, 9]})
    complete = tmp_path / "complete.xlsx"
    incomplete = tmp_path / "incomplete.xlsx"
    frame.to_excel(complete, index=False, engine="xlsxwriter")
    frame.to_excel(incomplete, index=False, engine="xlsxwriter", engine_kwargs={"options": {"constant_memory": True}})
    assert bench.written_dimension(str(complete)) == bench.written_dimension(str(incomplete))
    bench.check_output(str(complete), frame)
    with pytest.raises(RuntimeError, match="differs from the input at B2"):
        bench.check_output(str(incomplete), frame)


def test_check_output_refuses_wrong_interior_values(bench: ModuleType, tmp_path: Path) -> None:
    """A populated cell with the wrong value is no more complete than a missing one."""
    frame = pd.DataFrame({"a": [1, 2, 3], "b": [4, 5, 6]})
    changed = frame.copy()
    changed.loc[1, "b"] = 99
    target = tmp_path / "changed.xlsx"
    changed.to_excel(target, index=False)
    with pytest.raises(RuntimeError, match="differs from the input at B3"):
        bench.check_output(str(target), frame)


@pytest.mark.parametrize("library", ["xlsxturbo", "pandas + openpyxl", "pandas + xlsxwriter", "polars"])
@pytest.mark.parametrize("shape", ["mixed", "numeric", "strings"])
def test_all_benchmark_writers_verify_generated_data(
    bench: ModuleType, tmp_path: Path, library: str, shape: str,
) -> None:
    """Every compared writer passes real readback for each supported workload."""
    frame = bench.generate_test_data(12, 8, shape=shape)
    result = dict(bench.BENCHMARK_FUNCS)[library](frame, str(tmp_path / "output.xlsx"), len(frame))
    assert result.success, result.error


def reject_output(_filepath: str, _frame: pd.DataFrame) -> None:
    """Stand in for readback detecting a defective workbook."""
    raise RuntimeError("synthetic readback rejection")


@pytest.mark.parametrize("library", ["xlsxturbo", "pandas + openpyxl", "pandas + xlsxwriter", "polars"])
@pytest.mark.parametrize("verify_output", [True, False])
def test_benchmark_writers_obey_readback_result(
    bench: ModuleType, tmp_path: Path, monkeypatch: pytest.MonkeyPatch, library: str, verify_output: bool,
) -> None:
    """Timing wrappers must fail rejected output, except during the separate memory run."""
    monkeypatch.setattr(bench, "check_output", reject_output)
    frame = bench.generate_test_data(3, 5)
    result = dict(bench.BENCHMARK_FUNCS)[library](
        frame, str(tmp_path / "output.xlsx"), len(frame), verify_output=verify_output,
    )
    assert result.success is not verify_output
    assert result.error == ("synthetic readback rejection" if verify_output else None)


@pytest.mark.skipif(sys.platform == "win32", reason="memory measurement needs resource")
@pytest.mark.parametrize("library", ["xlsxturbo", "pandas + openpyxl", "pandas + xlsxwriter", "polars"])
def test_memory_child_excludes_readback(
    bench: ModuleType, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str], library: str,
) -> None:
    """The memory caller disables readback, so its allocations cannot enter the measurement."""
    monkeypatch.setattr(bench, "check_output", reject_output)
    bench.run_memory_child(library, 3, 5, "mixed")
    report = json.loads(capsys.readouterr().out)
    assert report["ok"] is True
    assert report["error"] is None


def scripted_writer(bench: ModuleType, name: str, outcomes: list[bool]) -> Callable[..., Any]:
    """Make a deterministic writer for success/failure reporting tests.

    Args:
        bench: Imported benchmark module.
        name: Library name.
        outcomes: Success flags in invocation order, including warmup if used.

    Returns:
        A writer returning the next outcome on every call.
    """
    remaining = iter(outcomes)

    def write(_frame: pd.DataFrame, _path: str, _rows: int) -> Any:
        """Return the next scripted result."""
        succeeded = next(remaining)
        return bench.BenchmarkResult(name, 2.0, 1.5, 0.1, succeeded, None if succeeded else "synthetic failure")

    return write


def test_failed_baseline_does_not_report_a_speedup(bench: ModuleType, monkeypatch: pytest.MonkeyPatch) -> None:
    """Missing baseline measurements stay null, and failed libraries stay visible."""
    monkeypatch.setattr(bench, "BENCHMARK_FUNCS", [
        ("xlsxturbo", scripted_writer(bench, "xlsxturbo", [False] * 3)),
        ("competitor", scripted_writer(bench, "competitor", [True] * 3)),
    ])
    results = bench.run_benchmarks(pd.DataFrame({"a": [1]}), 1, 1, runs=3, warmup=False, verbose=False)
    assert results["xlsxturbo"].median_time is None
    assert results["competitor"].speedup_vs_xlsxturbo is None
    output = json.loads(bench.format_json_output(results, 1, 1, 3, {}))
    failed = next(result for result in output["results"] if result["library"] == "xlsxturbo")
    assert failed["successful_runs"] == 0
    assert failed["attempted_runs"] == 3
    assert len(failed["errors"]) == 3
    assert failed["median_time_seconds"] is None
    assert all(result["speedup_vs_xlsxturbo"] is None for result in output["results"])


def test_partial_failure_reports_actual_sample_count(bench: ModuleType, monkeypatch: pytest.MonkeyPatch) -> None:
    """Two successful samples cannot masquerade as all three requested runs."""
    monkeypatch.setattr(bench, "BENCHMARK_FUNCS", [
        ("xlsxturbo", scripted_writer(bench, "xlsxturbo", [True, False, True])),
    ])
    results = bench.run_benchmarks(pd.DataFrame({"a": [1]}), 1, 1, runs=3, warmup=False, verbose=False)
    output = json.loads(bench.format_json_output(results, 1, 1, 3, {}))["results"][0]
    assert output["attempted_runs"] == 3
    assert output["successful_runs"] == 2
    assert output["all_times"] == [2.0, 2.0]
    assert output["errors"] == ["run 2: synthetic failure"]
    assert output["speedup_vs_xlsxturbo"] is None


@pytest.mark.parametrize("output_format", ["--json", "--markdown", "--quiet"])
@pytest.mark.parametrize("outcomes", [[False] * 4, [True, True, False, True], [True] * 4, [False, True, True, True]])
def test_benchmark_exit_status_and_reports_include_failures(
    bench: ModuleType, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str],
    output_format: str, outcomes: list[bool],
) -> None:
    """Failures, including warmup, survive every output format and make the command fail."""
    monkeypatch.setattr(bench, "BENCHMARK_FUNCS", [("xlsxturbo", scripted_writer(bench, "xlsxturbo", outcomes))])
    monkeypatch.setattr(sys, "argv", ["benchmark.py", "--rows", "2", "--cols", "2", "--runs", "3", output_format])
    assert bench.main() == (0 if all(outcomes) else 1)
    output = capsys.readouterr().out
    assert ("synthetic failure" in output) is (not all(outcomes))
    if output_format == "--json":
        result = json.loads(output)["results"][0]
        assert result["successful_runs"] == sum(outcomes[1:])
        assert result["attempted_runs"] == 3
        assert len(result["errors"]) == outcomes.count(False)
    else:
        # Console output pads this fraction; markdown does not.
        assert f"{sum(outcomes[1:])}/3" in output.replace(" ", "")


@pytest.mark.skipif(sys.platform == "win32", reason="peak memory uses the resource module")
def test_memory_child_reports_a_number(bench: ModuleType) -> None:
    """The fresh-process memory measurement runs end to end."""
    increase = bench.measure_peak_memory("xlsxturbo", 200, 5, "mixed")
    assert increase is not None
    assert increase >= 0
