import json
import os
import tempfile
import urllib.request

from PySide6.QtCore import QThread, Signal

from version import APP_VERSION


GITHUB_OWNER = "keeran997"
GITHUB_REPOSITORY = "Databank_export_converter"

INSTALLER_FILENAME = "DatabankExportConverterSetup.exe"

LATEST_RELEASE_URL = (
    f"https://api.github.com/repos/"
    f"{GITHUB_OWNER}/{GITHUB_REPOSITORY}/releases/latest"
)


def version_tuple(version):
    """
    Convert a version such as 'v1.2.3' or '1.2.3'
    into a tuple such as (1, 2, 3).
    """
    cleaned_version = version.strip().lower().lstrip("v")

    try:
        return tuple(
            int(part)
            for part in cleaned_version.split(".")
        )
    except ValueError as error:
        raise ValueError(
            f"Invalid version number: {version}"
        ) from error


class UpdateCheckThread(QThread):
    update_available = Signal(dict)
    no_update = Signal()
    error = Signal(str)

    def run(self):
        try:
            request = urllib.request.Request(
                LATEST_RELEASE_URL,
                headers={
                    "Accept": "application/vnd.github+json",
                    "User-Agent": "Databank-Export-Converter",
                },
            )

            with urllib.request.urlopen(
                request,
                timeout=10,
            ) as response:
                release = json.loads(
                    response.read().decode("utf-8")
                )

            latest_version = release["tag_name"]

            if version_tuple(latest_version) <= version_tuple(APP_VERSION):
                self.no_update.emit()
                return

            installer_url = None

            for asset in release.get("assets", []):
                if asset.get("name") == INSTALLER_FILENAME:
                    installer_url = asset.get(
                        "browser_download_url"
                    )
                    break

            if not installer_url:
                raise ValueError(
                    "An update was found, but the installer "
                    f"'{INSTALLER_FILENAME}' was not attached "
                    "to the GitHub release."
                )

            update_information = {
                "version": latest_version.lstrip("v"),
                "notes": release.get(
                    "body",
                    "No release notes were provided.",
                ),
                "download_url": installer_url,
            }

            self.update_available.emit(update_information)

        except Exception as error:
            self.error.emit(str(error))


class UpdateDownloadThread(QThread):
    progress = Signal(int)
    downloaded = Signal(str)
    error = Signal(str)

    def __init__(self, download_url):
        super().__init__()
        self.download_url = download_url

    def run(self):
        try:
            request = urllib.request.Request(
                self.download_url,
                headers={
                    "User-Agent": "Databank-Export-Converter",
                },
            )

            with urllib.request.urlopen(
                request,
                timeout=60,
            ) as response:
                total_size = int(
                    response.headers.get("Content-Length", 0)
                )

                installer_path = os.path.join(
                    tempfile.gettempdir(),
                    INSTALLER_FILENAME,
                )

                downloaded_size = 0
                chunk_size = 1024 * 64

                with open(installer_path, "wb") as installer_file:
                    while True:
                        chunk = response.read(chunk_size)

                        if not chunk:
                            break

                        installer_file.write(chunk)
                        downloaded_size += len(chunk)

                        if total_size > 0:
                            percentage = int(
                                downloaded_size
                                / total_size
                                * 100
                            )
                            self.progress.emit(percentage)

            if not os.path.isfile(installer_path):
                raise FileNotFoundError(
                    "The update installer was not downloaded."
                )

            self.progress.emit(100)
            self.downloaded.emit(installer_path)

        except Exception as error:
            self.error.emit(str(error))