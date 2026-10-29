"""The benchmark scripts' run order and output checks.

``benchmark_compression.py``, ``benchmark_threads.py`` and ``benchmark_parallel.py``
publish tables on ``docs/performance.md``. Each refuses an output that is not the
reference workbook, so a mode that writes less or different data cannot look fast;
these tests show each check can fail. ``run_order`` decides that no writer, level or
thread count always runs in the same slot.
"""

from __future__ import annotations

import importlib.util
import io
import sys
from pathlib import Path
from types import ModuleType
from typing import Any

import pandas as pd
import pytest
import xlsxturbo

from tests.helpers import REPO_ROOT, repo_checkout_available

pytestmark = pytest.mark.skipif(not repo_checkout_available(), reason="imports benchmarks/ from a source checkout")


def load_script(name: str, monkeypatch: pytest.MonkeyPatch) -> ModuleType:
    """Import one of the scripts in ``benchmarks/``, which import ``benchmark`` as a sibling.

    Args:
        name: The script's file name without ``.py``.
        monkeypatch: Puts ``benchmarks/`` on ``sys.path`` for the import.

    Returns:
        The module.
    """
    monkeypatch.syspath_prepend(str(REPO_ROOT / "benchmarks"))
    module_name = f"xlsxturbo_{name}_script"
    spec = importlib.util.spec_from_file_location(module_name, REPO_ROOT / "benchmarks" / f"{name}.py")
    assert spec is not None
    assert spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    # Registered first: @dataclass resolves the module through sys.modules.
    monkeypatch.setitem(sys.modules, module_name, module)
    spec.loader.exec_module(module)
    return module


@pytest.fixture
def bench(monkeypatch: pytest.MonkeyPatch) -> ModuleType:
    """``benchmarks/benchmark.py``."""
    return load_script("benchmark", monkeypatch)


@pytest.fixture
def compression(monkeypatch: pytest.MonkeyPatch) -> ModuleType:
    """``benchmarks/benchmark_compression.py``."""
    return load_script("benchmark_compression", monkeypatch)


@pytest.fixture
def threads(monkeypatch: pytest.MonkeyPatch) -> ModuleType:
    """``benchmarks/benchmark_threads.py``."""
    return load_script("benchmark_threads", monkeypatch)


@pytest.fixture
def parallel(monkeypatch: pytest.MonkeyPatch) -> ModuleType:
    """``benchmarks/benchmark_parallel.py``."""
    return load_script("benchmark_parallel", monkeypatch)


def export(frame: pd.DataFrame, level: int = 6) -> io.BytesIO:
    """Write a frame to a new buffer.

    Args:
        frame: The frame.
        level: The compression level.

    Returns:
        The buffer.
    """
    buffer = io.BytesIO()
    xlsxturbo.df_to_xlsx(frame, buffer, compression_level=level)
    return buffer


@pytest.mark.parametrize("count", [1, 2, 4, 9])
def test_run_order_puts_each_item_in_each_slot_once(bench: ModuleType, count: int) -> None:
    """Over as many runs as items, every item takes every position exactly once."""
    items = list(range(count))
    orders = [bench.run_order(items, run_num) for run_num in range(1, count + 1)]
    assert all(sorted(order) == items for order in orders)
    for position in range(count):
        assert sorted(order[position] for order in orders) == items


def test_run_benchmarks_rotates_the_writers(bench: ModuleType, monkeypatch: pytest.MonkeyPatch) -> None:
    """The timed runs call the writers in rotated order, so no library always goes first."""
    calls: list[str] = []

    def writer(name: str) -> Any:
        def write(_frame: pd.DataFrame, _path: str, _rows: int, **_kwargs: object) -> Any:
            calls.append(name)
            return bench.BenchmarkResult(name, 1.0, 1.0, 0.1, True)

        return write

    names = ["a", "b", "c"]
    monkeypatch.setattr(bench, "BENCHMARK_FUNCS", [(name, writer(name)) for name in names])
    bench.run_benchmarks(pd.DataFrame({"x": [1]}), 1, 1, runs=3, warmup=False, verbose=False)
    assert calls == ["a", "b", "c", "b", "c", "a", "c", "a", "b"]


def test_archive_contents_ignores_packing_and_sees_data(bench: ModuleType) -> None:
    """Two compression levels unpack to the same workbook; a changed cell does not."""
    frame = pd.DataFrame({"a": range(200), "b": ["x"] * 200})
    fast, small = export(frame, 1), export(frame, 9)
    assert fast.getvalue() != small.getvalue()
    assert bench.archive_contents(fast) == bench.archive_contents(small)
    changed = frame.copy()
    changed.loc[100, "a"] = -1
    assert bench.archive_contents(export(changed)) != bench.archive_contents(fast)


def test_compression_check_accepts_every_level(compression: ModuleType) -> None:
    """A real run's warmup outputs pass, and every level is timed for both functions."""
    frame = compression.generate_frame(300)
    results = compression.measure(frame, [1, 6, 9], runs=2, verbose=False)
    assert sorted(results) == [1, 6, 9]
    for result in results.values():
        assert all(len(result.times[function]) == 2 for function in compression.FUNCTIONS)
    assert results[1].size_bytes > results[9].size_bytes
    assert "| 6 (default) |" in compression.format_markdown(results, 300, 2)


@pytest.mark.parametrize("function", ["df_to_xlsx", "csv_to_xlsx"])
def test_compression_check_refuses_a_different_workbook(compression: ModuleType, function: str) -> None:
    """A level whose output holds other data is refused, for either function."""
    frame = compression.generate_frame(50)
    outputs = {(name, level): export(frame, level) for name in compression.FUNCTIONS for level in (1, 6)}
    compression.check_levels(outputs, frame)
    outputs[function, 1] = export(frame.head(49), 1)
    with pytest.raises(RuntimeError, match=f"{function} at level 1"):
        compression.check_levels(outputs, frame)


def test_compression_check_refuses_a_wrong_default_output(compression: ModuleType) -> None:
    """Level 6 is compared with the frame itself, not only with the other levels."""
    frame = compression.generate_frame(50)
    other = frame.copy()
    other.loc[10, "quantity"] = -1
    outputs = {(name, level): export(other, level) for name in compression.FUNCTIONS for level in (1, 6)}
    with pytest.raises(RuntimeError, match="differs from the input"):
        compression.check_levels(outputs, frame)


def test_threads_measure_times_every_count(threads: ModuleType, tmp_path: Path) -> None:
    """A real run's threaded exports pass the check, and every function and pool size is timed."""
    frame = pd.DataFrame({"a": range(100), "b": ["text"] * 100})
    csv_path = tmp_path / "frame.csv"
    frame.to_csv(csv_path, index=False)
    times = threads.measure(frame, str(csv_path), [1, 2], jobs=4, runs=2, verbose=False)
    assert {
        name: {count: len(samples) for count, samples in by_count.items()} for name, by_count in times.items()
    } == {"df_to_xlsx": {1: 2, 2: 2}, "csv_to_xlsx": {1: 2, 2: 2}}
    assert "| 2 |" in threads.format_markdown(times, 100, 2, 4, 2)


def test_threads_measure_refuses_a_csv_that_is_not_the_frame(threads: ModuleType, tmp_path: Path) -> None:
    """The ``csv_to_xlsx`` leg is checked against the frame, not only against itself."""
    frame = pd.DataFrame({"a": range(10)})
    csv_path = tmp_path / "frame.csv"
    frame.assign(a=frame["a"] + 1).to_csv(csv_path, index=False)
    with pytest.raises(RuntimeError, match="differs from the input"):
        threads.measure(frame, str(csv_path), [1, 2], jobs=2, runs=1, verbose=False)


def test_threads_check_refuses_one_different_export(threads: ModuleType) -> None:
    """One export out of a batch that differs from the serial one is refused."""
    frame = pd.DataFrame({"a": range(100)})
    reference = export(frame)
    batch = [export(frame) for _ in range(3)]
    threads.check_buffers(reference, {2: batch}, frame)
    batch[1] = export(frame.head(99))
    with pytest.raises(RuntimeError, match="export 1 on 2 threads"):
        threads.check_buffers(reference, {2: batch}, frame)


def test_threads_check_refuses_a_wrong_reference(threads: ModuleType) -> None:
    """The serial export is compared with the frame, so a shared defect is still caught."""
    frame = pd.DataFrame({"a": range(10)})
    wrong = export(frame.assign(a=frame["a"] + 1))
    with pytest.raises(RuntimeError, match="differs from the input"):
        threads.check_buffers(wrong, {1: [wrong]}, frame)


def test_parallel_benchmark_checks_and_times_both_modes(parallel: ModuleType, tmp_path: Path) -> None:
    """A real run passes the mode check and times each mode once per pair."""
    csv_path = parallel.generate_test_csv(str(tmp_path / "data.csv"), 300, 12, verbose=False)
    results = parallel.benchmark_conversion(csv_path, runs=2, verbose=False)
    assert set(results) == {False, True}
    assert "`parallel=True`" in parallel.format_markdown(results, 300, 12, 2)


def test_parallel_check_refuses_different_counts_or_content(parallel: ModuleType, tmp_path: Path) -> None:
    """Different reported counts and a different workbook are each refused."""
    frame = pd.DataFrame({"a": range(20)})
    single, other = tmp_path / "single.xlsx", tmp_path / "other.xlsx"
    xlsxturbo.df_to_xlsx(frame, single)
    xlsxturbo.df_to_xlsx(frame.head(19), other)
    parallel.check_modes({False: ((21, 1), single), True: ((21, 1), single)})
    with pytest.raises(RuntimeError, match="parallel reported"):
        parallel.check_modes({False: ((21, 1), single), True: ((20, 1), single)})
    with pytest.raises(RuntimeError, match="different workbook"):
        parallel.check_modes({False: ((21, 1), single), True: ((21, 1), other)})
