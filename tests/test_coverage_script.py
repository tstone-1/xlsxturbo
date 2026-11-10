"""The coverage script puts an ordinary build back when it is done.

``scripts/coverage_report.py`` installs an extension built with
``-Cinstrument-coverage`` into the venv. Left there, every later local test or
benchmark run measures the instrumented build, and nothing else notices: the
stale-build check compares modification times, and the instrumented build is the
newest file. These tests drive ``main()`` with every external command replaced by
a recorder, so they need no LLVM tools and build nothing.
"""

from __future__ import annotations

import importlib.util
import subprocess
import sys
import types
from pathlib import Path
from types import ModuleType
from typing import Any

import pytest

from tests.helpers import REPO_ROOT, repo_checkout_available

pytestmark = pytest.mark.skipif(
    not repo_checkout_available(), reason="imports scripts/ from a source checkout"
)


@pytest.fixture
def script() -> ModuleType:
    """Import ``scripts/coverage_report.py`` afresh, so a monkeypatch cannot leak.

    Returns:
        The module.
    """
    name = "xlsxturbo_coverage_report_script"
    spec = importlib.util.spec_from_file_location(name, REPO_ROOT / "scripts" / "coverage_report.py")
    assert spec is not None
    assert spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class Recorder:
    """Stands in for both ``run`` and ``subprocess.run``, in one ordered log."""

    def __init__(self, profile_dir: Path, fail_on: str | None = None, restore_code: int = 0) -> None:
        """Set up the recorder.

        Args:
            profile_dir: Where the instrumented pytest pass should leave a profile.
            fail_on: An argument that makes the script's ``run`` exit like a failed command.
            restore_code: The exit code the final rebuild reports.
        """
        self.profile_dir = profile_dir
        self.fail_on = fail_on
        self.restore_code = restore_code
        self.calls: list[tuple[str, list[str], dict[str, str]]] = []

    def script_run(
        self, command: list[str], *, env: dict[str, str] | None = None, capture: bool = False
    ) -> subprocess.CompletedProcess[str]:
        """Record a call through the script's own ``run``.

        Returns:
            A successful completed process.
        """
        del capture
        self.calls.append(("run", list(command), dict(env or {})))
        if self.fail_on is not None and self.fail_on in command:
            sys.exit(f"[FAIL] command failed (1): {' '.join(command)}")
        if "pytest" in command and env and "RUSTFLAGS" in env:
            # The instrumented pass is what produces the profiles step 4 merges.
            (self.profile_dir / "xlsxturbo-1-1.profraw").write_bytes(b"")
        return subprocess.CompletedProcess(command, 0, stdout="", stderr="")

    def subprocess_run(self, command: list[str], **kwargs: Any) -> subprocess.CompletedProcess[str]:
        """Record a direct ``subprocess.run``, which only the final rebuild uses.

        Returns:
            A completed process with ``restore_code``.
        """
        self.calls.append(("subprocess", list(command), dict(kwargs.get("env") or {})))
        return subprocess.CompletedProcess(command, self.restore_code)


def install(script: ModuleType, recorder: Recorder, monkeypatch: pytest.MonkeyPatch) -> None:
    """Replace every external effect of the script with the recorder."""
    monkeypatch.setattr(script, "PROFILE_DIR", recorder.profile_dir)
    monkeypatch.setattr(script, "llvm_tool", lambda name: name)
    monkeypatch.setattr(script, "find_extension", lambda: Path("extension"))
    monkeypatch.setattr(script, "build_instrumented_binaries", lambda env: ([Path("rust-tests")], []))
    monkeypatch.setattr(script, "run", recorder.script_run)
    monkeypatch.setattr(script.subprocess, "run", recorder.subprocess_run)
    # Step 5 probes for coverage.py, which the CI test jobs do not install.
    monkeypatch.setitem(sys.modules, "coverage", types.ModuleType("coverage"))
    monkeypatch.setattr(sys, "argv", ["coverage_report.py"])


def is_develop(command: list[str]) -> bool:
    """Whether a command is a ``maturin develop`` build."""
    return any(Path(part).stem == "maturin" for part in command) and "develop" in command


def is_instrumented(env: dict[str, str]) -> bool:
    """Whether an environment carries the instrumentation flag."""
    return any("instrument-coverage" in env.get(name, "") for name in ("RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS"))


def test_the_last_build_is_uninstrumented(
    script: ModuleType, tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    """After a successful run the last command is an ordinary rebuild, after both pytest passes."""
    recorder = Recorder(tmp_path / "coverage")
    install(script, recorder, monkeypatch)

    assert script.main() == 0

    builds = [env for _, command, env in recorder.calls if is_develop(command)]
    # The control: the instrumented build was there to be undone.
    assert [is_instrumented(env) for env in builds] == [True, False]
    _, command, env = recorder.calls[-1]
    assert is_develop(command)
    assert not is_instrumented(env)
    pytest_passes = [i for i, (_, command, _) in enumerate(recorder.calls) if "pytest" in command]
    assert len(pytest_passes) == 2
    assert max(pytest_passes) < len(recorder.calls) - 1
    assert "[OK] the venv holds an uninstrumented build" in capsys.readouterr().out


def test_a_failed_run_still_restores_the_build(
    script: ModuleType, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """A failure after the instrumented build is installed still ends with the rebuild."""
    recorder = Recorder(tmp_path / "coverage", fail_on="pytest")
    install(script, recorder, monkeypatch)

    with pytest.raises(SystemExit) as raised:
        script.main()

    # The original failure survives the cleanup; a rebuild must not replace it.
    assert "command failed" in str(raised.value.code)
    _, command, env = recorder.calls[-1]
    assert is_develop(command)
    assert not is_instrumented(env)


def test_a_failed_rebuild_warns_instead_of_raising(
    script: ModuleType, tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    """The rebuild runs in a ``finally``, so it reports a failure rather than raising one."""
    recorder = Recorder(tmp_path / "coverage", restore_code=1)
    install(script, recorder, monkeypatch)

    assert script.restore_uninstrumented_build() is False
    assert "[WARN] the venv still holds the INSTRUMENTED build" in capsys.readouterr().out


def test_an_inherited_instrument_flag_is_dropped(
    script: ModuleType, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """A caller's own ``RUSTFLAGS`` with the flag would make the rebuild instrumented again."""
    recorder = Recorder(tmp_path / "coverage")
    install(script, recorder, monkeypatch)
    monkeypatch.setenv("RUSTFLAGS", "-Cinstrument-coverage")

    assert script.restore_uninstrumented_build() is True
    _, command, env = recorder.calls[-1]
    assert is_develop(command)
    assert not is_instrumented(env)


def test_the_instrumented_builds_keep_their_profiles_out_of_the_repo_root(
    script: ModuleType, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Build scripts and proc macros are instrumented too, and write a profile on exit.

    Measured: without a profile path for the two builds, a run left seven
    ``default_*.profraw`` files in the repository root, written during steps 1 and 3.
    They must also stay out of the top level of the profile directory, which step 4
    merges.
    """
    profile_dir = tmp_path / "coverage"
    recorder = Recorder(profile_dir)
    install(script, recorder, monkeypatch)
    build_envs: list[dict[str, str]] = []

    def build(env: dict[str, str]) -> tuple[list[Path], list[Path]]:
        build_envs.append(env)
        return [Path("rust-tests")], []

    monkeypatch.setattr(script, "build_instrumented_binaries", build)

    assert script.main() == 0

    develop = [env for _, command, env in recorder.calls if is_develop(command) and is_instrumented(env)]
    assert len(build_envs) == 1
    assert len(develop) == 1
    for env in (*build_envs, *develop):
        target = Path(env["LLVM_PROFILE_FILE"])
        assert target.parent.is_relative_to(profile_dir)
        assert target.parent != profile_dir


def test_the_python_layer_pass_keeps_its_profiles_out_of_the_repo_root(
    script: ModuleType, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Step 5 runs the instrumented extension too, so it needs a profile path of its own.

    It also keeps the coverage marker, which ``tests/test_concurrency.py`` keys its
    timing skip on.
    """
    profile_dir = tmp_path / "coverage"
    recorder = Recorder(profile_dir)
    install(script, recorder, monkeypatch)

    script.report_python_layer()

    passes = [env for _, command, env in recorder.calls if "pytest" in command]
    assert len(passes) == 1
    env = passes[0]
    assert env.get("XLSXTURBO_COVERAGE") == "1"
    assert Path(env["LLVM_PROFILE_FILE"]).parent.is_relative_to(profile_dir)
