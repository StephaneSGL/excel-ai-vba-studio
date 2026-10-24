"""Synthetic, cross-platform regressions for native OOXML inflation budgets."""

from __future__ import annotations

import hashlib
import importlib.util
import io
import sys
import tempfile
import types
import zipfile
from pathlib import Path


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
    spec = importlib.util.spec_from_file_location("vba_writeback_cli_budget", cli_path)
    if spec is None or spec.loader is None:
        raise AssertionError(f"Could not load {cli_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def write_zip(path: Path, entries: list[tuple[str, bytes]]) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in entries:
            archive.writestr(name, payload)


cli = load_cli()
with tempfile.TemporaryDirectory(prefix="excel-ai-native-zip-budget-") as directory:
    root = Path(directory)
    normal = root / "normal.xlsm"
    write_zip(
        normal,
        [
            ("[Content_Types].xml", b"<Types/>"),
            ("xl/workbook.xml", b"workbook"),
            ("xl/vbaProject.bin", b"vba payload"),
            ("xl/media/highly-compressible.bin", b"A" * 4096),
            ("xl/empty-dir/", b""),
        ],
    )
    with zipfile.ZipFile(normal, "r") as archive:
        empty_directory = archive.getinfo("xl/empty-dir/")
        assert empty_directory.file_size == 0
        assert empty_directory.compress_size == 2

    original_read = zipfile.ZipFile.read

    def forbidden_eager_read(*_args, **_kwargs):
        raise AssertionError("zip_payload_hashes must stream members instead of using ZipFile.read")

    zipfile.ZipFile.read = forbidden_eager_read
    try:
        hashes = cli.zip_payload_hashes(
            normal,
            maximum_entry_bytes=4096,
            maximum_total_bytes=8192,
        )
    finally:
        zipfile.ZipFile.read = original_read

    assert "xl/vbaProject.bin" not in hashes
    assert hashes["xl/workbook.xml"] == hashlib.sha256(b"workbook").hexdigest()
    assert hashes["xl/media/highly-compressible.bin"] == hashlib.sha256(b"A" * 4096).hexdigest()
    assert hashes["xl/empty-dir/"] == hashlib.sha256(b"").hexdigest()

    exact = root / "exact.xlsm"
    write_zip(exact, [("exact.bin", b"E" * 64)])
    cli._preflight_zip_payloads(
        exact,
        maximum_entry_bytes=64,
        maximum_total_bytes=64,
    )
    try:
        cli._preflight_zip_payloads(
            exact,
            maximum_entry_bytes=63,
            maximum_total_bytes=64,
        )
        raise AssertionError("Per-entry overflow was accepted")
    except cli.WritebackError as error:
        assert error.code == "WORKBOOK_TOO_LARGE"

    aggregate = root / "aggregate.xlsm"
    write_zip(aggregate, [("one.bin", b"1" * 40), ("two.bin", b"2" * 40)])
    try:
        cli._preflight_zip_payloads(
            aggregate,
            maximum_entry_bytes=64,
            maximum_total_bytes=79,
        )
        raise AssertionError("Aggregate overflow was accepted")
    except cli.WritebackError as error:
        assert error.code == "WORKBOOK_TOO_LARGE"

    class UnderstatedInfo:
        filename = "understated.bin"
        file_size = 4
        compress_size = 5

        @staticmethod
        def is_dir() -> bool:
            return False

    class UnderstatedArchive:
        @staticmethod
        def open(_info, _mode):
            return io.BytesIO(b"12345")

    try:
        cli._consume_zip_entry(
            UnderstatedArchive(),
            UnderstatedInfo(),
            maximum_entry_bytes=8,
            remaining_total_bytes=8,
        )
        raise AssertionError("An understated expanded size was accepted")
    except ValueError as error:
        assert "inconsistent" in str(error)

    preflight_error = cli.WritebackError("WORKBOOK_TOO_LARGE", "synthetic preflight refusal")
    original_preflight = cli._preflight_zip_payloads
    original_excel_file = cli.ExcelFile

    class ForbiddenExcelFile:
        def __init__(self, *_args, **_kwargs):
            nonlocal_state["constructed"] = True
            raise AssertionError("ExcelFile was constructed before ZIP preflight")

    nonlocal_state = {"constructed": False}
    preflight_paths: list[Path] = []

    def refuse_preflight(path: Path):
        preflight_paths.append(path)
        raise preflight_error

    cli._preflight_zip_payloads = refuse_preflight
    cli.ExcelFile = ForbiddenExcelFile
    try:
        for request_handler, request in [
            (cli.inspect_request, {"workbookPath": str(normal)}),
            (cli.fingerprint_request, {"workbookPath": str(normal)}),
            (
                cli.apply_request,
                {
                    "workbookPath": str(normal),
                    "expectedWorkbookSha256": hashlib.sha256(
                        normal.read_bytes()
                    ).hexdigest(),
                    "patches": [
                        {
                            "moduleName": "Module1",
                            "componentKind": "module",
                            "source": 'Attribute VB_Name = "Module1"\n',
                        }
                    ],
                },
            ),
        ]:
            try:
                request_handler(request)
                raise AssertionError("Request bypassed ZIP payload preflight")
            except cli.WritebackError as error:
                assert error is preflight_error
        assert not nonlocal_state["constructed"]
        assert len(preflight_paths) == 3
        assert all(path != normal for path in preflight_paths)
        assert all(not path.parent.exists() for path in preflight_paths)
    finally:
        cli._preflight_zip_payloads = original_preflight
        cli.ExcelFile = original_excel_file

print(
    "Native ZIP budget regression passed: declared and actual inflation are bounded, "
    "hashing streams, and every operation preflights before pyOpenVBA."
)
