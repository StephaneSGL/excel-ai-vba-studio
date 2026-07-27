"""Windows regression checks for conflict-safe workbook path replacement."""

from __future__ import annotations

import hashlib
import importlib.util
import os
import sys
import tempfile
import types
from pathlib import Path


if os.name != "nt":
    print("Atomic write-back regression skipped: Windows semantics required.")
    raise SystemExit(0)


def install_pyopenvba_stubs() -> None:
    package = types.ModuleType("pyopenvba")
    package.__path__ = []
    sys.modules["pyopenvba"] = package

    cfb = types.ModuleType("pyopenvba.cfb")
    cfb.CFB = object
    sys.modules[cfb.__name__] = cfb

    excel = types.ModuleType("pyopenvba.excel")
    excel.ExcelFile = object
    sys.modules[excel.__name__] = excel

    exceptions = types.ModuleType("pyopenvba.exceptions")
    exceptions.CFBError = type("CFBError", (Exception,), {})
    exceptions.PyOpenVBAError = type("PyOpenVBAError", (Exception,), {})
    exceptions.VBAProjectError = type("VBAProjectError", (Exception,), {})
    sys.modules[exceptions.__name__] = exceptions

    vba = types.ModuleType("pyopenvba.vba")
    vba.VBAModuleKind = type(
        "VBAModuleKind",
        (),
        {"standard": "standard", "other": "other"},
    )
    vba.detect_signature = lambda _cfb: None
    vba.parse_project_stream = lambda _data: None
    sys.modules[vba.__name__] = vba


def load_cli():
    install_pyopenvba_stubs()
    cli_path = Path(__file__).resolve().parents[1] / "native" / "vba-writeback" / "cli.py"
    spec = importlib.util.spec_from_file_location("vba_writeback_cli", cli_path)
    if spec is None or spec.loader is None:
        raise AssertionError(f"Could not load {cli_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


cli = load_cli()
with tempfile.TemporaryDirectory(prefix="excel-ai-atomic-writeback-") as directory:
    root = Path(directory)
    workbook = root / "book.xlsm"
    replacement = root / "replacement.xlsm"
    baseline = b"baseline workbook"
    patched = b"patched workbook"
    external = b"newer external workbook"
    workbook.write_bytes(baseline)
    replacement.write_bytes(patched)
    backup = cli.create_backup_path(workbook, hashlib.sha256(baseline).hexdigest())

    cli.replace_file_with_backup(workbook, replacement, backup)
    assert workbook.read_bytes() == patched
    assert backup.read_bytes() == baseline
    assert not replacement.exists()

    preserved = cli.restore_displaced_workbook(
        workbook,
        backup,
        sha256(workbook),
        backup,
    )
    assert preserved == backup
    assert workbook.read_bytes() == baseline
    assert backup.read_bytes() == baseline

    workbook.write_bytes(external)
    preserved = cli.restore_displaced_workbook(
        workbook,
        backup,
        hashlib.sha256(patched).hexdigest(),
        backup,
    )
    assert preserved == backup
    assert workbook.read_bytes() == external

    workbook.unlink()
    preserved = cli.restore_missing_workbook(workbook, backup, backup)
    assert preserved == backup
    assert workbook.read_bytes() == baseline
    assert backup.read_bytes() == baseline

    partial_workbook = root / "partial.xlsm"
    partial_replacement = root / "partial-replacement.xlsm"
    partial_backup = cli.create_backup_path(
        partial_workbook,
        hashlib.sha256(baseline).hexdigest(),
    )
    partial_replacement.write_bytes(patched)
    partial_backup.write_bytes(baseline)
    try:
        cli.handle_failed_atomic_replace(
            partial_workbook,
            partial_replacement,
            partial_backup,
            hashlib.sha256(patched).hexdigest(),
            OSError(1177, "simulated partial ReplaceFileW failure"),
        )
        raise AssertionError("Partial replacement failure was not reported")
    except cli.WritebackError as error:
        assert error.code == "WORKBOOK_REPLACE_FAILED"
    assert partial_workbook.read_bytes() == baseline
    assert partial_backup.read_bytes() == baseline
    assert partial_replacement.read_bytes() == patched

    concurrent_target = root / "concurrent-winner.xlsm"
    original_move = cli.move_file_without_replacement

    def install_concurrent_winner(_source: Path, target: Path) -> None:
        target.write_bytes(external)
        raise OSError(183, "simulated destination already exists")

    cli.move_file_without_replacement = install_concurrent_winner
    try:
        preserved = cli.restore_missing_workbook(
            concurrent_target,
            backup,
            backup,
        )
    finally:
        cli.move_file_without_replacement = original_move
    assert preserved == backup
    assert concurrent_target.read_bytes() == external
    assert backup.read_bytes() == baseline

print(
    "Atomic write-back regression passed: displaced backup, safe restore, "
    "partial failure recovery, concurrent winner preserved."
)
