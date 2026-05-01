from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from pprint import pformat
from typing import Any

import FreeSimpleGUI as sg

try:
    import win32wnet
except ImportError:  # pragma: no cover - allows non-Windows/sans-pywin32 import
    win32wnet = None

try:
    import tomllib
except ModuleNotFoundError:  # pragma: no cover - Python < 3.11 fallback
    import tomli as tomllib


@dataclass
class ScanResult:
    folder_path: Path
    toml_found: bool
    file_count: int
    last_modified: datetime | None = None
    status: str = "OK"
    metadata: dict[str, Any] = field(default_factory=dict)


class NetworkConnectionError(Exception):
    pass


def parse_toml_file(path: Path) -> dict[str, Any]:
    with path.open("rb") as fp:
        return tomllib.load(fp)


def get_unc_share_root(path: str) -> str:
    if not path.startswith("\\\\"):
        return ""

    parts = [part for part in path.strip("\\").split("\\") if part]
    if len(parts) < 2:
        return path
    return f"\\\\{parts[0]}\\{parts[1]}"


def connect_network_share(remote_path: str, username: str = "", password: str = "") -> bool:
    share_root = get_unc_share_root(remote_path)
    if not share_root:
        return False

    if win32wnet is None:
        raise NetworkConnectionError(
            "pywin32 is required to connect network shares explicitly."
        )

    net_resource = win32wnet.NETRESOURCE()
    net_resource.lpRemoteName = share_root

    try:
        win32wnet.WNetAddConnection2(
            net_resource,
            password or None,
            username or None,
            0,
        )
    except win32wnet.error as exc:
        # Windows error 1219 means another connection to the same server exists
        # with different credentials. The old CLI handled this by disconnecting
        # and retrying, which is useful on shared lab machines.
        error_code = getattr(exc, "winerror", exc.args[0] if exc.args else None)
        if error_code == 1219:
            disconnect_network_share(share_root)
            win32wnet.WNetAddConnection2(
                net_resource,
                password or None,
                username or None,
                0,
            )
        else:
            raise NetworkConnectionError(str(exc)) from exc

    return True


def disconnect_network_share(remote_path: str) -> None:
    share_root = get_unc_share_root(remote_path)
    if win32wnet is None or not share_root:
        return

    try:
        win32wnet.WNetCancelConnection2(share_root, 0, 0)
    except win32wnet.error:
        pass


def scan_folders(root_path: str | Path, toml_filename: str) -> list[ScanResult]:
    root = Path(root_path)
    results: list[ScanResult] = []

    try:
        root_exists = root.exists()
    except OSError as exc:
        return [
            ScanResult(
                folder_path=root,
                toml_found=False,
                file_count=0,
                status=f"Cannot access root path: {exc}",
            )
        ]

    if not root_exists:
        return [
            ScanResult(
                folder_path=root,
                toml_found=False,
                file_count=0,
                status="Root path does not exist",
            )
        ]

    pending = [root]
    while pending:
        folder = pending.pop()
        toml_path = folder / toml_filename
        metadata: dict[str, Any] = {}
        status = "OK"
        last_modified = None
        toml_found = False
        file_count = 0

        try:
            children = list(folder.iterdir())
            subfolders = [child for child in children if child.is_dir()]
            file_count = sum(1 for child in children if child.is_file())
            toml_found = toml_path.is_file()

            if toml_found:
                try:
                    metadata = parse_toml_file(toml_path)
                    last_modified = datetime.fromtimestamp(toml_path.stat().st_mtime)
                except PermissionError as exc:
                    status = f"Permission denied reading TOML: {exc}"
                except tomllib.TOMLDecodeError as exc:
                    status = f"TOML parse error: {exc}"
                except OSError as exc:
                    status = f"TOML read error: {exc}"
            else:
                status = "TOML not found"

            pending.extend(reversed(subfolders))
        except PermissionError as exc:
            status = f"Permission denied: {exc}"
        except OSError as exc:
            status = f"Folder read error: {exc}"

        results.append(
            ScanResult(
                folder_path=folder,
                toml_found=toml_found,
                file_count=file_count,
                last_modified=last_modified,
                status=status,
                metadata=metadata,
            )
        )

    return results


def sync_to_server(scan_results: list[ScanResult]) -> None:
    # Future API sync should be added here after local scanning is stable.
    # The old CLI posted discovered file metadata with dataset_ref, name,
    # timestamps, and URL fields; keep that shape in mind for the API payload.
    # This placeholder intentionally does not upload anything yet.
    return None


def format_results_for_table(scan_results: list[ScanResult]) -> list[list[str | int]]:
    rows: list[list[str | int]] = []
    for result in scan_results:
        rows.append(
            [
                str(result.folder_path),
                "Yes" if result.toml_found else "No",
                result.file_count,
                result.last_modified.strftime("%Y-%m-%d %H:%M:%S")
                if result.last_modified
                else "",
                result.status,
            ]
        )
    return rows


def build_unc_path(computer: str, shared_folder: str) -> str:
    computer = computer.strip().strip("\\/")
    shared_folder = shared_folder.strip()

    if shared_folder.startswith("\\\\") or Path(shared_folder).is_absolute():
        return shared_folder
    if computer and shared_folder:
        cleaned_share = shared_folder.strip("\\/")
        return f"\\\\{computer}\\{cleaned_share}"
    return shared_folder


def build_window() -> sg.Window:
    sg.theme("SystemDefault")

    layout = [
        [
            sg.Text("Computer name or IP", size=(18, 1)),
            sg.Input(key="-COMPUTER-", expand_x=True),
        ],
        [
            sg.Text("Shared folder path", size=(18, 1)),
            sg.Input(key="-FOLDER-", expand_x=True),
            sg.FolderBrowse("Browse", target="-FOLDER-"),
        ],
        [
            sg.Text("Network username", size=(18, 1)),
            sg.Input(key="-USERNAME-", expand_x=True),
            sg.Text("Password"),
            sg.Input(key="-PASSWORD-", password_char="*", size=(24, 1)),
        ],
        [
            sg.Text("TOML filename", size=(18, 1)),
            sg.Input("folder_info.toml", key="-TOML-", expand_x=True),
            sg.Button("Scan", bind_return_key=True),
        ],
        [
            sg.Table(
                values=[],
                headings=[
                    "Folder path",
                    "TOML found",
                    "File count",
                    "Last modified",
                    "Status",
                ],
                key="-RESULTS-",
                auto_size_columns=False,
                col_widths=[48, 10, 10, 18, 28],
                justification="left",
                num_rows=14,
                expand_x=True,
                expand_y=True,
                enable_events=True,
                select_mode=sg.TABLE_SELECT_MODE_BROWSE,
            )
        ],
        [
            sg.Multiline(
                "",
                key="-PREVIEW-",
                size=(100, 12),
                expand_x=True,
                expand_y=True,
                disabled=True,
                autoscroll=True,
            )
        ],
        [sg.StatusBar("Ready", key="-STATUS-", expand_x=True)],
    ]

    return sg.Window(
        "Network Folder Metadata Scanner",
        layout,
        resizable=True,
        finalize=True,
    )


def main() -> None:
    window = build_window()
    scan_results: list[ScanResult] = []

    while True:
        event, values = window.read()

        if event == sg.WINDOW_CLOSED:
            break

        if event == "Scan":
            toml_filename = values["-TOML-"].strip() or "folder_info.toml"
            root_path = build_unc_path(values["-COMPUTER-"], values["-FOLDER-"])
            connected_share = False

            if not root_path:
                window["-STATUS-"].update("Enter a shared folder path or UNC path.")
                continue

            try:
                if root_path.startswith("\\\\") and (
                    values["-USERNAME-"].strip() or values["-PASSWORD-"]
                ):
                    window["-STATUS-"].update("Connecting to network share...")
                    window.refresh()
                    connected_share = connect_network_share(
                        root_path,
                        values["-USERNAME-"].strip(),
                        values["-PASSWORD-"],
                    )

                window["-STATUS-"].update(f"Scanning {root_path}...")
                window["-PREVIEW-"].update("")
                window.refresh()

                scan_results = scan_folders(root_path, toml_filename)
                window["-RESULTS-"].update(format_results_for_table(scan_results))
                window["-STATUS-"].update(
                    f"Scan complete: {len(scan_results)} folders scanned."
                )
            except NetworkConnectionError as exc:
                scan_results = [
                    ScanResult(
                        folder_path=Path(root_path),
                        toml_found=False,
                        file_count=0,
                        status=f"Network connection error: {exc}",
                    )
                ]
                window["-RESULTS-"].update(format_results_for_table(scan_results))
                window["-STATUS-"].update(str(exc))
            finally:
                if connected_share:
                    disconnect_network_share(root_path)

        if event == "-RESULTS-":
            selected_rows = values["-RESULTS-"]
            if not selected_rows:
                continue

            result = scan_results[selected_rows[0]]
            if result.metadata:
                preview = pformat(result.metadata, width=100, sort_dicts=False)
            elif result.toml_found:
                preview = result.status
            else:
                preview = "No TOML metadata found for this folder."

            window["-PREVIEW-"].update(preview)
            window["-STATUS-"].update(result.status)

    window.close()


if __name__ == "__main__":
    main()
