"""The ``xlsxturbo`` console script installed by the wheel.

``python/xlsxturbo/cli.py`` wraps ``csv_to_xlsx``. The Rust binary in
``src/main.rs`` does the same for source builds, and the two must accept the same
options and exit with the same codes; ``TestParityWithRustBinary`` compares them.
"""

from __future__ import annotations

import datetime as dt
import re
import subprocess
import sys
from importlib.metadata import entry_points
from pathlib import Path

import pytest
import xlsxturbo
from xlsxturbo import cli

from tests.helpers import HAS_OPENPYXL, REPO_ROOT, active_ws, load_workbook, repo_checkout_available

pytestmark = pytest.mark.skipif(not HAS_OPENPYXL, reason="openpyxl required for content verification")


@pytest.fixture
def csv_file(tmp_path: Path) -> Path:
    """A small CSV whose date is ambiguous between US and European order.

    Args:
        tmp_path: pytest's per-test directory.

    Returns:
        The CSV path.
    """
    path = tmp_path / "in.csv"
    path.write_text("id,day\n1,01-02-2024\n")
    return path


class TestConversion:
    """What the command does when asked correctly."""

    def test_success_prints_ok_and_writes_the_file(
        self, csv_file: Path, tmp_path: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        """Exit 0, ``OK <rows> <cols>`` on stdout, and a readable workbook."""
        target = tmp_path / "out.xlsx"

        assert cli.main([str(csv_file), str(target), "--sheet-name", "Data"]) == 0

        assert capsys.readouterr().out == "OK 2 2\n"
        assert load_workbook(target).sheetnames == ["Data"]

    @pytest.mark.parametrize(("order", "expected"), [("US", dt.datetime(2024, 1, 2)), ("eu", dt.datetime(2024, 2, 1))])
    def test_date_order_reaches_the_conversion(
        self, csv_file: Path, tmp_path: Path, order: str, expected: dt.datetime
    ) -> None:
        """``--date-order`` is honoured, in any case."""
        target = tmp_path / "out.xlsx"

        assert cli.main([str(csv_file), str(target), "-d", order]) == 0

        assert active_ws(load_workbook(target))["B2"].value == expected

    def test_verbose_reports_on_stderr_only(
        self, csv_file: Path, tmp_path: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        """``-v`` never adds to stdout, which scripts parse."""
        assert cli.main([str(csv_file), str(tmp_path / "out.xlsx"), "-v", "-p"]) == 0

        captured = capsys.readouterr()
        assert captured.out == "OK 2 2\n"
        assert "Converted 2 rows x 2 cols" in captured.err
        assert "Parallel: True" in captured.err


class TestExitCodes:
    """1 means the conversion failed; 2 means the command line was wrong."""

    def test_failed_conversion_exits_1(self, tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
        """A missing input is a failed conversion, reported on stderr."""
        assert cli.main([str(tmp_path / "missing.csv"), str(tmp_path / "out.xlsx")]) == cli.EXIT_FAILURE

        captured = capsys.readouterr()
        assert captured.out == ""
        assert captured.err.startswith("Error: Failed to open input file")

    def test_invalid_date_order_exits_2(self, csv_file: Path, tmp_path: Path) -> None:
        """An invalid date order is a usage error, as in the Rust binary."""
        with pytest.raises(SystemExit) as exc:
            cli.main([str(csv_file), str(tmp_path / "out.xlsx"), "--date-order", "ymd"])
        assert exc.value.code == cli.EXIT_USAGE

    def test_missing_argument_exits_2(self) -> None:
        """The usage errors argparse catches itself use the same code."""
        with pytest.raises(SystemExit) as exc:
            cli.main(["only-one.csv"])
        assert exc.value.code == cli.EXIT_USAGE

    def test_every_accepted_date_order_is_accepted_by_csv_to_xlsx(self, csv_file: Path, tmp_path: Path) -> None:
        """The CLI's list comes from ``types.DateOrder``; the extension must agree with it."""
        for order in cli.DATE_ORDERS:
            xlsxturbo.csv_to_xlsx(csv_file, tmp_path / f"{order}.xlsx", date_order=order)


class TestInstallation:
    """The command is reachable the ways the docs say."""

    def test_console_script_is_registered(self) -> None:
        """The installed distribution declares ``xlsxturbo = xlsxturbo.cli:main``."""
        scripts = entry_points(group="console_scripts", name="xlsxturbo")
        assert [ep.value for ep in scripts] == ["xlsxturbo.cli:main"]

    def test_python_dash_m(self) -> None:
        """``python -m xlsxturbo`` runs the same command."""
        result = subprocess.run(
            [sys.executable, "-m", "xlsxturbo", "--version"],
            capture_output=True,
            text=True,
            check=False,
        )
        assert result.returncode == 0
        assert result.stdout.strip() == f"xlsxturbo {xlsxturbo.__version__}"


@pytest.mark.skipif(not repo_checkout_available(), reason="needs src/main.rs from a source checkout")
class TestParityWithRustBinary:
    """The Python command and ``src/main.rs`` offer the same options."""

    @staticmethod
    def rust_options() -> set[str]:
        """The option strings clap derives from ``struct Args``.

        Returns:
            Every ``-x``/``--long-name`` the Rust binary accepts, ``--version`` and
            ``--help`` included.
        """
        source = (REPO_ROOT / "src" / "main.rs").read_text(encoding="utf-8")
        body = re.search(r"struct Args \{(.*?)\n\}", source, re.DOTALL)
        assert body, "struct Args not found in src/main.rs"
        options = {"-h", "--help"}
        if "#[command(version)]" in source:
            options |= {"-V", "--version"}
        for attrs, field in re.findall(r"#\[arg\(([^)]*)\)\]\s*(\w+):", body.group(1)):
            flags = {part.strip().split("=")[0].strip() for part in attrs.split(",")}
            if "long" in flags:
                options.add("--" + field.replace("_", "-"))
            if "short" in flags:
                options.add("-" + field[0])
        assert len(options) > 4, f"parsed too few options from src/main.rs: {options}"
        return options

    def test_same_options(self) -> None:
        """Both CLIs accept exactly the same flags."""
        python_options = {opt for action in cli.build_parser()._actions for opt in action.option_strings}
        assert python_options == self.rust_options()

    def test_same_positionals(self) -> None:
        """Both take an input and an output path, in that order."""
        python_positionals = [a.dest for a in cli.build_parser()._actions if not a.option_strings]
        assert python_positionals == ["input", "output"]
        source = (REPO_ROOT / "src" / "main.rs").read_text(encoding="utf-8")
        assert re.search(r"struct Args \{\s*///[^\n]*\n\s*input: String,\s*///[^\n]*\n\s*output: String,", source)
