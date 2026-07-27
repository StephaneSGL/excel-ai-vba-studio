"""Transactional VBA source write-back for Excel AI & VBA Studio.

This executable never starts Office, never changes AccessVBOM, and never runs
macros.  It edits the documented VBA binary container through pyOpenVBA.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import secrets
import shutil
import stat
import sys
import tempfile
import time
import zipfile
from pathlib import Path
from typing import Any, NoReturn

from pyopenvba.cfb import CFB
from pyopenvba.excel import ExcelFile
from pyopenvba.exceptions import CFBError, PyOpenVBAError, VBAProjectError
from pyopenvba.vba import (
    VBAModuleKind,
    detect_signature,
    parse_project_stream,
)


SCHEMA_VERSION = 1
MAX_REQUEST_BYTES = 5 * 1024 * 1024
MAX_WORKBOOK_BYTES = 512 * 1024 * 1024
MAX_PATCHES = 64
MAX_MODULE_SOURCE_CHARACTERS = 2_000_000
MAX_PROJECT_SOURCE_CHARACTERS = 8_000_000
SUPPORTED_EXTENSIONS = {".xlsm", ".xlam"}
SUPPORTED_KINDS = {"module", "class", "document", "userform"}
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
WINDOWS_REPARSE_POINT = 0x400
ATTRIBUTE_NAME_RE = re.compile(
    r'(?im)^Attribute[ \t]+VB_Name[ \t]*=[ \t]*"([^"\r\n]+)"[ \t]*\r?$'
)


class WritebackError(Exception):
    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


def assert_no_reparse_point_chain(candidate: Path) -> None:
    """Refuse symlinks and Windows junctions without resolving through them."""
    full_path = Path(os.path.abspath(candidate))
    current = Path(full_path.anchor)
    for part in full_path.parts[len(current.parts) :]:
        current /= part
        try:
            item_stat = os.lstat(current)
        except FileNotFoundError:
            break
        if stat.S_ISLNK(item_stat.st_mode) or (
            getattr(item_stat, "st_file_attributes", 0) & WINDOWS_REPARSE_POINT
        ):
            raise WritebackError(
                "REPARSE_POINT_REFUSED",
                f"Symbolic links and Windows reparse points are refused: {current}",
            )


def assert_local_windows_path(candidate: Path) -> None:
    raw_path = str(candidate).replace("/", "\\")
    if raw_path.startswith("\\\\"):
        raise WritebackError(
            "NETWORK_PATH_REFUSED",
            "UNC and device network paths are refused. Copy the workbook to a local drive.",
        )
    if os.name != "nt":
        return
    import ctypes

    drive_root = Path(candidate.anchor)
    drive_type = ctypes.windll.kernel32.GetDriveTypeW(str(drive_root))
    if drive_type in {0, 1, 4}:
        raise WritebackError(
            "NETWORK_PATH_REFUSED",
            "The workbook must be stored on a verified local Windows drive.",
        )


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha256_text(value: str) -> str:
    return sha256_bytes(value.encode("utf-8"))


def sha256_file(file_path: Path) -> str:
    digest = hashlib.sha256()
    with file_path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def require_string(
    record: dict[str, Any],
    key: str,
    *,
    allow_empty: bool = False,
) -> str:
    value = record.get(key)
    if not isinstance(value, str) or (not allow_empty and not value):
        raise WritebackError("INVALID_REQUEST", f"{key} must be a string.")
    return value


def validate_module_name(value: str) -> str:
    if (
        len(value) > 31
        or value in {".", ".."}
        or any(ord(char) < 32 for char in value)
        or any(char in value for char in '/\\:"')
    ):
        raise WritebackError(
            "INVALID_MODULE_NAME",
            f"Invalid VBA component name: {value!r}.",
        )
    return value


def normalize_newlines(value: str) -> str:
    return value.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")


def split_userform_source(source: str) -> tuple[str, str]:
    normalized = normalize_newlines(source)
    match = ATTRIBUTE_NAME_RE.search(normalized)
    if match is None:
        raise WritebackError(
            "USERFORM_SOURCE_INVALID",
            "An existing UserForm source must contain Attribute VB_Name.",
        )
    return normalized[: match.start()], normalized[match.start() :]


def assert_attribute_name(source: str, expected_name: str) -> None:
    match = ATTRIBUTE_NAME_RE.search(source)
    if match is not None and match.group(1).casefold() != expected_name.casefold():
        raise WritebackError(
            "MODULE_NAME_MISMATCH",
            "Attribute VB_Name does not match the requested component name.",
        )


def read_request(request_path: Path) -> dict[str, Any]:
    if not request_path.is_absolute():
        raise WritebackError("INVALID_REQUEST_PATH", "Request path must be absolute.")
    size = request_path.stat().st_size
    if size <= 0 or size > MAX_REQUEST_BYTES:
        raise WritebackError(
            "REQUEST_TOO_LARGE",
            f"Request must be between 1 and {MAX_REQUEST_BYTES} bytes.",
        )
    try:
        value = json.loads(request_path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise WritebackError("INVALID_REQUEST", f"Cannot read request: {error}") from error
    if not isinstance(value, dict):
        raise WritebackError("INVALID_REQUEST", "Request root must be an object.")
    if value.get("schemaVersion") != SCHEMA_VERSION:
        raise WritebackError(
            "UNSUPPORTED_SCHEMA",
            f"schemaVersion must be {SCHEMA_VERSION}.",
        )
    return value


def read_project_metadata(workbook: ExcelFile) -> tuple[set[str], set[str], set[str], set[str]]:
    cfb = CFB.from_bytes(workbook.vba_project_bytes())
    try:
        project_stream = parse_project_stream(cfb.get_stream("PROJECT"))
    except (KeyError, VBAProjectError) as error:
        raise WritebackError(
            "PROJECT_STREAM_INVALID",
            f"Cannot classify VBA components: {error}",
        ) from error
    return (
        {name.casefold() for name in project_stream.standard_modules},
        {name.casefold() for name in project_stream.class_modules},
        {name.casefold() for name, _ in project_stream.document_modules},
        {name.casefold() for name in project_stream.base_classes},
    )


def classify_existing_component(
    module_name: str,
    standard_names: set[str],
    class_names: set[str],
    document_names: set[str],
    form_names: set[str],
) -> str:
    key = module_name.casefold()
    if key in form_names:
        return "userform"
    if key in document_names:
        return "document"
    if key in standard_names:
        return "module"
    if key in class_names:
        return "class"
    # Some producers omit PROJECT declarations. The dir record still tells us
    # whether a component is procedural, but it cannot distinguish a class
    # from a document or designer. Fail closed instead of guessing.
    raise WritebackError(
        "COMPONENT_KIND_UNKNOWN",
        f"Cannot safely classify existing component {module_name!r}.",
    )


def zip_payload_hashes(workbook_path: Path) -> dict[str, str]:
    with zipfile.ZipFile(workbook_path, "r") as archive:
        return {
            info.filename: sha256_bytes(archive.read(info.filename))
            for info in archive.infolist()
            if info.filename != "xl/vbaProject.bin"
        }


def designer_stream_hashes(workbook: ExcelFile, form_names: set[str]) -> dict[str, str]:
    cfb = CFB.from_bytes(workbook.vba_project_bytes())
    hashes: dict[str, str] = {}

    def walk_storage(storage_index: int, prefix: str) -> None:
        storage = cfb._directory[storage_index]  # pyOpenVBA 3.1.0 pinned API
        for child_index in sorted(cfb._collect_subtree(storage.child_id)):
            child = cfb._directory[child_index]
            child_path = f"{prefix}/{child.name}"
            if child.obj_type == 1:
                walk_storage(child_index, child_path)
            elif child.obj_type == 2:
                hashes[child_path] = sha256_bytes(cfb._read_stream(child))

    for form_name in sorted(form_names):
        try:
            storage_index = cfb._find_storage_index(form_name)
        except KeyError as error:
            raise WritebackError(
                "USERFORM_DESIGNER_MISSING",
                f"Designer storage is missing for UserForm {form_name!r}.",
            ) from error
        walk_storage(storage_index, form_name)
    return hashes


def project_stream_fingerprint(workbook: ExcelFile) -> dict[str, Any]:
    cfb = CFB.from_bytes(workbook.vba_project_bytes())
    streams: dict[str, str] = {}
    storages: list[str] = []

    def walk_storage(storage_index: int, prefix: str) -> None:
        storage = cfb._directory[storage_index]  # pyOpenVBA 3.1.0 pinned API
        for child_index in sorted(cfb._collect_subtree(storage.child_id)):
            child = cfb._directory[child_index]
            child_path = f"{prefix}/{child.name}" if prefix else child.name
            if child.obj_type == 1:
                storages.append(child_path)
                walk_storage(child_index, child_path)
            elif child.obj_type == 2:
                streams[child_path] = sha256_bytes(cfb._read_stream(child))

    walk_storage(0, "")
    inventory = {
        "streams": dict(sorted(streams.items())),
        "storages": sorted(storages),
    }
    encoded = json.dumps(
        inventory,
        ensure_ascii=True,
        separators=(",", ":"),
        sort_keys=True,
    ).encode("ascii")
    return {
        "sha256": sha256_bytes(encoded),
        "streamCount": len(streams),
        "storageCount": len(storages),
    }


def fingerprint_request(request: dict[str, Any]) -> dict[str, Any]:
    workbook_path = validate_workbook_path(request)
    with ExcelFile(workbook_path) as workbook:
        project = workbook.vba_project()
        validation_errors = workbook.validate()
        if validation_errors:
            raise WritebackError(
                "VBA_PROJECT_INVALID",
                "VBA project validation failed: " + "; ".join(validation_errors),
            )
        signature = detect_signature(workbook._get_cfb())
        fingerprint = project_stream_fingerprint(workbook)
        return {
            "ok": True,
            "operation": "fingerprint",
            "workbookSha256": sha256_file(workbook_path),
            "projectName": project.name,
            "protected": bool(
                project.protection is not None and project.protection.has_password
            ),
            "signed": signature.present,
            "signatureKinds": signature.kinds,
            "projectFingerprintSha256": fingerprint["sha256"],
            "projectStreamCount": fingerprint["streamCount"],
            "projectStorageCount": fingerprint["storageCount"],
        }


def create_backup_path(workbook_path: Path, original_hash: str) -> Path:
    backup_directory = workbook_path.parent / ".excel-ai-vba-backups"
    assert_no_reparse_point_chain(backup_directory)
    backup_directory.mkdir(exist_ok=True)
    assert_no_reparse_point_chain(backup_directory)
    stamp = time.strftime("%Y%m%d-%H%M%S")
    for _ in range(16):
        backup_path = backup_directory / (
            f"{workbook_path.stem}.{stamp}.{original_hash[:12]}."
            f"{secrets.token_hex(4)}{workbook_path.suffix}"
        )
        if not backup_path.exists():
            assert_no_reparse_point_chain(backup_path)
            return backup_path
    raise WritebackError(
        "BACKUP_NAME_UNAVAILABLE",
        "A unique backup path could not be allocated.",
    )


def replace_file_with_backup(
    workbook_path: Path,
    replacement_path: Path,
    displaced_path: Path,
) -> None:
    """Atomically install replacement_path and capture the actual destination."""
    if os.name != "nt":
        raise WritebackError(
            "WINDOWS_REQUIRED",
            "Transactional VBA write-back requires Windows ReplaceFile semantics.",
        )

    import ctypes
    from ctypes import wintypes

    replace_file = ctypes.WinDLL("kernel32", use_last_error=True).ReplaceFileW
    replace_file.argtypes = [
        wintypes.LPCWSTR,
        wintypes.LPCWSTR,
        wintypes.LPCWSTR,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.LPVOID,
    ]
    replace_file.restype = wintypes.BOOL
    if not replace_file(
        str(workbook_path),
        str(replacement_path),
        str(displaced_path),
        0,
        None,
        None,
    ):
        error_code = ctypes.get_last_error()
        raise OSError(
            error_code,
            ctypes.FormatError(error_code),
            str(workbook_path),
        )


def move_file_without_replacement(source_path: Path, target_path: Path) -> None:
    """Atomically move a same-volume recovery file only if target is absent."""
    if os.name != "nt":
        raise WritebackError(
            "WINDOWS_REQUIRED",
            "Transactional VBA write-back requires Windows move semantics.",
        )

    import ctypes
    from ctypes import wintypes

    move_file = ctypes.WinDLL("kernel32", use_last_error=True).MoveFileExW
    move_file.argtypes = [
        wintypes.LPCWSTR,
        wintypes.LPCWSTR,
        wintypes.DWORD,
    ]
    move_file.restype = wintypes.BOOL
    move_file_write_through = 0x00000008
    if not move_file(
        str(source_path),
        str(target_path),
        move_file_write_through,
    ):
        error_code = ctypes.get_last_error()
        raise OSError(
            error_code,
            ctypes.FormatError(error_code),
            str(target_path),
        )


def preserve_displaced_candidate(
    workbook_path: Path,
    candidate_path: Path,
    primary_backup_path: Path,
) -> Path:
    candidate_absolute = Path(os.path.abspath(candidate_path))
    primary_absolute = Path(os.path.abspath(primary_backup_path))
    if candidate_absolute == primary_absolute:
        return primary_backup_path

    candidate_hash = sha256_file(candidate_path)
    preserved_path = create_backup_path(workbook_path, candidate_hash)
    shutil.copy2(candidate_path, preserved_path)
    if sha256_file(preserved_path) != candidate_hash:
        raise WritebackError(
            "BACKUP_VALIDATION_FAILED",
            "A concurrently displaced workbook backup could not be verified.",
        )
    return preserved_path


def restore_missing_workbook(
    workbook_path: Path,
    candidate_path: Path,
    primary_backup_path: Path,
) -> Path:
    """Restore candidate only while the workbook path remains absent."""
    candidate_hash = sha256_file(candidate_path)
    preserved_path = preserve_displaced_candidate(
        workbook_path,
        candidate_path,
        primary_backup_path,
    )
    temporary_paths: list[Path] = []
    try:
        for _ in range(8):
            if workbook_path.exists():
                return preserved_path

            restore_handle, restore_name = tempfile.mkstemp(
                prefix=f".{workbook_path.stem}.excel-ai-missing-restore-",
                suffix=workbook_path.suffix,
                dir=workbook_path.parent,
            )
            os.close(restore_handle)
            restore_path = Path(restore_name)
            temporary_paths.append(restore_path)
            shutil.copy2(candidate_path, restore_path)
            if sha256_file(restore_path) != candidate_hash:
                raise WritebackError(
                    "BACKUP_VALIDATION_FAILED",
                    "A missing workbook recovery copy could not be verified.",
                )
            try:
                move_file_without_replacement(restore_path, workbook_path)
            except OSError:
                if workbook_path.exists():
                    return preserved_path
                continue

            if workbook_path.exists():
                return preserved_path

        raise WritebackError(
            "ROLLBACK_FAILED",
            "The workbook path remained unavailable during recovery. "
            f"The displaced version is preserved at {preserved_path}.",
        )
    finally:
        for temporary_path in temporary_paths:
            try:
                if temporary_path.exists():
                    temporary_path.unlink()
            except OSError:
                pass


def restore_displaced_workbook(
    workbook_path: Path,
    displaced_path: Path,
    installed_hash: str,
    primary_backup_path: Path,
) -> Path:
    """Restore the latest file displaced by this process without blind overwrite."""
    candidate_path = displaced_path
    cleanup_paths: list[Path] = []
    try:
        for _ in range(8):
            preserved_path = preserve_displaced_candidate(
                workbook_path,
                candidate_path,
                primary_backup_path,
            )
            if (
                Path(os.path.abspath(candidate_path))
                != Path(os.path.abspath(primary_backup_path))
                and candidate_path not in cleanup_paths
            ):
                cleanup_paths.append(candidate_path)
            try:
                current_hash = sha256_file(workbook_path)
            except FileNotFoundError:
                return restore_missing_workbook(
                    workbook_path,
                    preserved_path,
                    preserved_path,
                )
            if current_hash != installed_hash:
                return preserved_path

            restore_handle, restore_name = tempfile.mkstemp(
                prefix=f".{workbook_path.stem}.excel-ai-restore-",
                suffix=workbook_path.suffix,
                dir=workbook_path.parent,
            )
            os.close(restore_handle)
            restore_path = Path(restore_name)
            cleanup_paths.append(restore_path)
            shutil.copy2(candidate_path, restore_path)

            capture_handle, capture_name = tempfile.mkstemp(
                prefix=f".{workbook_path.stem}.excel-ai-displaced-",
                suffix=workbook_path.suffix,
                dir=workbook_path.parent,
            )
            os.close(capture_handle)
            capture_path = Path(capture_name)
            capture_path.unlink()

            candidate_hash = sha256_file(restore_path)
            try:
                replace_file_with_backup(
                    workbook_path,
                    restore_path,
                    capture_path,
                )
            except OSError as error:
                recovery_candidate = (
                    capture_path if capture_path.exists() else candidate_path
                )
                if workbook_path.exists():
                    preserved_path = preserve_displaced_candidate(
                        workbook_path,
                        recovery_candidate,
                        primary_backup_path,
                    )
                else:
                    preserved_path = restore_missing_workbook(
                        workbook_path,
                        recovery_candidate,
                        primary_backup_path,
                    )
                if (
                    Path(os.path.abspath(recovery_candidate))
                    != Path(os.path.abspath(primary_backup_path))
                    and recovery_candidate not in cleanup_paths
                ):
                    cleanup_paths.append(recovery_candidate)
                raise WritebackError(
                    "ROLLBACK_FAILED",
                    "A recovery swap failed after preserving or restoring the "
                    f"displaced workbook at {preserved_path}: {error}",
                ) from error

            captured_hash = sha256_file(capture_path)
            if captured_hash == installed_hash:
                cleanup_paths.append(capture_path)
                return preserved_path

            candidate_path = capture_path
            installed_hash = candidate_hash

        preserved_path = preserve_displaced_candidate(
            workbook_path,
            candidate_path,
            primary_backup_path,
        )
        if (
            Path(os.path.abspath(candidate_path))
            != Path(os.path.abspath(primary_backup_path))
            and candidate_path not in cleanup_paths
        ):
            cleanup_paths.append(candidate_path)
        raise WritebackError(
            "ROLLBACK_FAILED",
            "Repeated concurrent writes prevented automatic restoration. "
            f"The last displaced version is preserved at {preserved_path}.",
        )
    finally:
        for temporary_path in cleanup_paths:
            try:
                if temporary_path.exists():
                    temporary_path.unlink()
            except OSError:
                pass


def handle_failed_atomic_replace(
    workbook_path: Path,
    replacement_path: Path,
    backup_path: Path,
    replacement_hash: str,
    error: OSError,
) -> NoReturn:
    """Recover every documented partial ReplaceFileW failure state."""
    if backup_path.exists():
        if workbook_path.exists():
            try:
                current_hash = sha256_file(workbook_path)
            except FileNotFoundError:
                current_hash = ""
            if current_hash == replacement_hash:
                recovery_path = restore_displaced_workbook(
                    workbook_path,
                    backup_path,
                    replacement_hash,
                    backup_path,
                )
            else:
                recovery_path = preserve_displaced_candidate(
                    workbook_path,
                    backup_path,
                    backup_path,
                )
        else:
            recovery_path = restore_missing_workbook(
                workbook_path,
                backup_path,
                backup_path,
            )
        raise WritebackError(
            "WORKBOOK_REPLACE_FAILED",
            "The atomic replacement failed after displacing the workbook. "
            f"The displaced version was restored or preserved at {recovery_path}: {error}",
        ) from error

    if not workbook_path.exists():
        recovery_path = create_backup_path(
            workbook_path,
            replacement_hash,
        )
        shutil.copy2(replacement_path, recovery_path)
        raise WritebackError(
            "ROLLBACK_FAILED",
            "The atomic replacement failed and the workbook path is "
            f"missing. The validated edited copy is preserved at {recovery_path}: {error}",
        ) from error

    raise WritebackError(
        "WORKBOOK_REPLACE_FAILED",
        f"The workbook could not be replaced atomically: {error}",
    ) from error


def validate_workbook_path(request: dict[str, Any]) -> Path:
    workbook_path = Path(require_string(request, "workbookPath"))
    if not workbook_path.is_absolute():
        raise WritebackError("INVALID_WORKBOOK_PATH", "Workbook path must be absolute.")
    assert_local_windows_path(workbook_path)
    assert_no_reparse_point_chain(workbook_path)
    if workbook_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise WritebackError(
            "UNSUPPORTED_WORKBOOK",
            "Automatic VBA write-back supports .xlsm and .xlam only.",
        )
    if not workbook_path.is_file():
        raise WritebackError("WORKBOOK_NOT_FOUND", f"Workbook not found: {workbook_path}")
    workbook_size = workbook_path.stat().st_size
    if workbook_size <= 0 or workbook_size > MAX_WORKBOOK_BYTES:
        raise WritebackError(
            "WORKBOOK_TOO_LARGE",
            f"Workbook exceeds the {MAX_WORKBOOK_BYTES}-byte safety limit.",
        )
    return workbook_path


def inspect_request(request: dict[str, Any]) -> dict[str, Any]:
    workbook_path = validate_workbook_path(request)
    with ExcelFile(workbook_path) as workbook:
        project = workbook.vba_project()
        validation_errors = workbook.validate()
        if validation_errors:
            raise WritebackError(
                "VBA_PROJECT_INVALID",
                "VBA project validation failed: " + "; ".join(validation_errors),
            )
        (
            standard_names,
            class_names,
            document_names,
            form_names,
        ) = read_project_metadata(workbook)
        signature = detect_signature(workbook._get_cfb())
        modules: list[dict[str, Any]] = []
        total_characters = 0
        for module in project.modules:
            total_characters += len(module.source)
            if total_characters > MAX_PROJECT_SOURCE_CHARACTERS:
                raise WritebackError(
                    "PROJECT_TOO_LARGE",
                    f"VBA sources exceed {MAX_PROJECT_SOURCE_CHARACTERS} characters.",
                )
            modules.append(
                {
                    "name": module.name,
                    "componentKind": classify_existing_component(
                        module.name,
                        standard_names,
                        class_names,
                        document_names,
                        form_names,
                    ),
                    "source": module.source,
                    "sourceSha256": sha256_text(module.source),
                }
            )
        return {
            "ok": True,
            "operation": "inspect",
            "workbookSha256": sha256_file(workbook_path),
            "projectName": project.name,
            "protected": bool(
                project.protection is not None and project.protection.has_password
            ),
            "signed": signature.present,
            "signatureKinds": signature.kinds,
            "designerStreamsSha256": designer_stream_hashes(workbook, form_names),
            "modules": modules,
        }


def apply_request(request: dict[str, Any]) -> dict[str, Any]:
    workbook_path = validate_workbook_path(request)

    expected_workbook_hash = require_string(request, "expectedWorkbookSha256").lower()
    if SHA256_RE.fullmatch(expected_workbook_hash) is None:
        raise WritebackError(
            "INVALID_WORKBOOK_HASH",
            "expectedWorkbookSha256 must be a lowercase SHA-256 digest.",
        )

    patches = request.get("patches")
    if not isinstance(patches, list) or not 1 <= len(patches) <= MAX_PATCHES:
        raise WritebackError(
            "INVALID_PATCHES",
            f"patches must contain between 1 and {MAX_PATCHES} entries.",
        )

    parsed_patches: list[dict[str, str]] = []
    total_source_characters = 0
    seen_names: set[str] = set()
    for raw_patch in patches:
        if not isinstance(raw_patch, dict):
            raise WritebackError("INVALID_PATCH", "Each patch must be an object.")
        module_name = validate_module_name(require_string(raw_patch, "moduleName"))
        component_kind = require_string(raw_patch, "componentKind").lower()
        if component_kind not in SUPPORTED_KINDS:
            raise WritebackError(
                "INVALID_COMPONENT_KIND",
                f"Unsupported componentKind: {component_kind!r}.",
            )
        source = require_string(raw_patch, "source", allow_empty=True)
        if len(source) > MAX_MODULE_SOURCE_CHARACTERS:
            raise WritebackError(
                "SOURCE_TOO_LARGE",
                f"{module_name} exceeds {MAX_MODULE_SOURCE_CHARACTERS} characters.",
            )
        total_source_characters += len(source)
        if total_source_characters > MAX_PROJECT_SOURCE_CHARACTERS:
            raise WritebackError(
                "PROJECT_TOO_LARGE",
                f"Patch payload exceeds {MAX_PROJECT_SOURCE_CHARACTERS} characters.",
            )
        key = module_name.casefold()
        if key in seen_names:
            raise WritebackError(
                "DUPLICATE_COMPONENT",
                f"Component {module_name!r} appears more than once.",
            )
        seen_names.add(key)
        parsed_patches.append(
            {
                "moduleName": module_name,
                "componentKind": component_kind,
                "source": source,
                "expectedDesignerSha256": str(
                    raw_patch.get("expectedDesignerSha256", "")
                ).lower(),
            }
        )

    work_handle, work_name = tempfile.mkstemp(
        prefix=f".{workbook_path.stem}.excel-ai-vba-",
        suffix=workbook_path.suffix,
        dir=workbook_path.parent,
    )
    os.close(work_handle)
    work_path = Path(work_name)
    assert_no_reparse_point_chain(work_path)
    backup_path: Path | None = None
    try:
        shutil.copy2(workbook_path, work_path)
        current_workbook_hash = sha256_file(work_path)
        if current_workbook_hash != expected_workbook_hash:
            raise WritebackError(
                "STALE_WORKBOOK",
                "Workbook changed after export. Refresh the VBA project before applying edits.",
            )
        original_zip_hashes = zip_payload_hashes(work_path)
        with ExcelFile(work_path) as workbook:
            project = workbook.vba_project()
            validation_errors = workbook.validate()
            if validation_errors:
                raise WritebackError(
                    "VBA_PROJECT_INVALID",
                    "VBA project validation failed: " + "; ".join(validation_errors),
                )
            if project.protection is not None and project.protection.has_password:
                raise WritebackError(
                    "VBA_PROJECT_PROTECTED",
                    "Password-protected VBA projects are not modified.",
                )
            signature = detect_signature(workbook._get_cfb())
            if signature.present:
                raise WritebackError(
                    "VBA_PROJECT_SIGNED",
                    "Digitally signed VBA projects are not modified because edits invalidate the signature.",
                )

            (
                standard_names,
                class_names,
                document_names,
                form_names,
            ) = read_project_metadata(workbook)
            before_form_hashes = designer_stream_hashes(workbook, form_names)
            before_sources = {
                module.name.casefold(): module.source for module in project.modules
            }
            expected_sources = dict(before_sources)
            modified_names: list[str] = []

            for patch in parsed_patches:
                module_name = patch["moduleName"]
                component_kind = patch["componentKind"]
                source = normalize_newlines(patch["source"])
                key = module_name.casefold()
                existing = next(
                    (module for module in project.modules if module.name.casefold() == key),
                    None,
                )

                if existing is None:
                    if component_kind == "userform":
                        raise WritebackError(
                            "NEW_USERFORM_UNSUPPORTED",
                            "Creating a new UserForm designer is not supported. Use an existing template form.",
                        )
                    if component_kind == "document":
                        raise WritebackError(
                            "NEW_DOCUMENT_MODULE_UNSUPPORTED",
                            "Document modules are created only by the Excel workbook structure.",
                        )
                    assert_attribute_name(source, module_name)
                    project.add_module(
                        module_name,
                        source,
                        kind=(
                            VBAModuleKind.standard
                            if component_kind == "module"
                            else VBAModuleKind.other
                        ),
                    )
                    expected_sources[key] = project.get_module(module_name).source
                    modified_names.append(module_name)
                    continue

                actual_kind = classify_existing_component(
                    module_name,
                    standard_names,
                    class_names,
                    document_names,
                    form_names,
                )
                if actual_kind != component_kind:
                    raise WritebackError(
                        "COMPONENT_KIND_MISMATCH",
                        f"{module_name} is {actual_kind}, not {component_kind}.",
                    )

                if component_kind == "userform":
                    # A VBE-exported .frm concatenates textual designer metadata
                    # with the code module. Inside vbaProject.bin the designer is
                    # stored separately (UserFormName/... streams); only the
                    # Attribute/code suffix belongs in VBA/UserFormName.
                    designer_source, source = split_userform_source(source)
                    expected_designer_hash = patch["expectedDesignerSha256"]
                    if SHA256_RE.fullmatch(expected_designer_hash) is None:
                        raise WritebackError(
                            "USERFORM_DESIGNER_HASH_REQUIRED",
                            "Existing UserForm edits require expectedDesignerSha256.",
                        )
                    if sha256_text(designer_source) != expected_designer_hash:
                        raise WritebackError(
                            "USERFORM_DESIGNER_CHANGED",
                            "UserForm layout changes are not written in v1; only code-behind may change.",
                        )

                assert_attribute_name(source, module_name)
                workbook.set_module(module_name, source)
                updated_source = project.get_module(module_name).source
                expected_sources[key] = updated_source
                if updated_source != before_sources[key]:
                    modified_names.append(module_name)

            if not modified_names:
                if sha256_file(workbook_path) != expected_workbook_hash:
                    raise WritebackError(
                        "STALE_WORKBOOK",
                        "Workbook changed while the no-op write was being checked.",
                    )
                return {
                    "ok": True,
                    "changed": False,
                    "modifiedModules": [],
                    "workbookSha256": current_workbook_hash,
                    "backupPath": None,
                }
            backup_path = create_backup_path(workbook_path, current_workbook_hash)
            workbook.save()

        with ExcelFile(work_path) as validated:
            validation_errors = validated.validate()
            if validation_errors:
                raise WritebackError(
                    "WRITEBACK_VALIDATION_FAILED",
                    "Patched VBA project is inconsistent: " + "; ".join(validation_errors),
                )
            after_project = validated.vba_project()
            after_sources = {
                module.name.casefold(): module.source for module in after_project.modules
            }
            if set(after_sources) != set(expected_sources):
                raise WritebackError(
                    "MODULE_SET_CHANGED",
                    "The VBA module set changed unexpectedly during write-back.",
                )
            for module_name, expected_source in expected_sources.items():
                if after_sources[module_name] != expected_source:
                    raise WritebackError(
                        "SOURCE_ROUNDTRIP_FAILED",
                        f"Source validation failed for {module_name}.",
                    )
            (
                _standard_names,
                _class_names,
                _document_names,
                validated_form_names,
            ) = read_project_metadata(validated)
            after_form_hashes = designer_stream_hashes(
                validated,
                validated_form_names,
            )
            if after_form_hashes != before_form_hashes:
                raise WritebackError(
                    "USERFORM_DESIGNER_CHANGED",
                    "UserForm designer streams changed unexpectedly.",
                )

        if zip_payload_hashes(work_path) != original_zip_hashes:
            raise WritebackError(
                "OOXML_PAYLOAD_CHANGED",
                "A non-VBA OOXML part changed unexpectedly.",
            )

        if sha256_file(workbook_path) != expected_workbook_hash:
            raise WritebackError(
                "STALE_WORKBOOK",
                "Workbook changed during validation. The external version was left untouched.",
            )
        patched_workbook_hash = sha256_file(work_path)
        try:
            replace_file_with_backup(workbook_path, work_path, backup_path)
        except OSError as error:
            handle_failed_atomic_replace(
                workbook_path,
                work_path,
                backup_path,
                patched_workbook_hash,
                error,
            )

        displaced_workbook_hash = sha256_file(backup_path)
        if displaced_workbook_hash != expected_workbook_hash:
            preserved_path = restore_displaced_workbook(
                workbook_path,
                backup_path,
                patched_workbook_hash,
                backup_path,
            )
            raise WritebackError(
                "STALE_WORKBOOK",
                "Workbook changed during the atomic commit. "
                f"The displaced external version was restored or preserved at {preserved_path}.",
            )

        installed_workbook_hash = sha256_file(workbook_path)
        if installed_workbook_hash != patched_workbook_hash:
            raise WritebackError(
                "STALE_WORKBOOK",
                "Workbook changed immediately after the atomic commit. "
                "The newer external version was left untouched.",
            )
        return {
            "ok": True,
            "changed": True,
            "modifiedModules": modified_names,
            "workbookSha256": installed_workbook_hash,
            "backupPath": str(backup_path),
        }
    finally:
        try:
            if work_path.exists():
                work_path.unlink()
        except OSError:
            pass


def emit_result(result: dict[str, Any], exit_code: int) -> int:
    sys.stdout.write(json.dumps(result, ensure_ascii=False, separators=(",", ":")))
    sys.stdout.write("\n")
    return exit_code


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="excel-ai-vba-writeback")
    parser.add_argument("request", type=Path, help="Absolute path to a JSON request.")
    args = parser.parse_args(argv)
    try:
        request = read_request(args.request)
        operation = request.get("operation", "apply")
        if operation == "inspect":
            return emit_result(inspect_request(request), 0)
        if operation == "fingerprint":
            return emit_result(fingerprint_request(request), 0)
        if operation != "apply":
            raise WritebackError(
                "UNSUPPORTED_OPERATION",
                "operation must be apply, inspect, or fingerprint.",
            )
        return emit_result(apply_request(request), 0)
    except WritebackError as error:
        return emit_result(
            {"ok": False, "code": error.code, "message": str(error)},
            2,
        )
    except (CFBError, PyOpenVBAError, OSError, ValueError, KeyError) as error:
        return emit_result(
            {"ok": False, "code": "WRITEBACK_FAILED", "message": str(error)},
            3,
        )
    except Exception as error:  # pragma: no cover - last-resort fail-closed path
        return emit_result(
            {"ok": False, "code": "UNEXPECTED_ERROR", "message": str(error)},
            4,
        )


if __name__ == "__main__":
    raise SystemExit(main())
