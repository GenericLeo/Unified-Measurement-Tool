"""GitHub release update manager for the Unified Measurement Tool."""

import json
import re
import ssl
import sys
import urllib.error
import urllib.request
import webbrowser
from typing import Dict, Optional

from version import __github_repo__, __version__


class UpdateManager:
    """Check GitHub releases and direct users to the latest installer/download."""

    def __init__(self, github_repo: Optional[str] = None) -> None:
        self.github_repo = github_repo or __github_repo__
        self.current_version = __version__
        self.api_url = f"https://api.github.com/repos/{self.github_repo}/releases/latest"
        self.platform_label = self._get_platform_label()

    def check_for_updates(self, timeout: int = 10) -> Dict[str, object]:
        """Check GitHub Releases for a newer version."""
        try:
            request = urllib.request.Request(
                self.api_url,
                headers={"Accept": "application/vnd.github.v3+json"},
            )

            with urllib.request.urlopen(
                request,
                timeout=timeout,
                context=ssl.create_default_context(),
            ) as response:
                release_data = json.loads(response.read().decode("utf-8"))

            assets = release_data.get("assets", [])
            latest_version = self._resolve_release_version(release_data, assets)
            download_asset = self._resolve_download_asset(assets)
            release_notes = release_data.get("body") or "No release notes available."

            return {
                "update_available": self._is_newer_version(latest_version),
                "latest_version": latest_version,
                "current_version": self.current_version,
                "download_url": download_asset.get("browser_download_url") if download_asset else None,
                "download_name": download_asset.get("name") if download_asset else None,
                "platform": self.platform_label,
                "release_notes": release_notes,
                "html_url": release_data.get("html_url"),
                "error": None,
            }
        except urllib.error.HTTPError as exc:
            if exc.code == 404:
                error = "Repository or release not found on GitHub."
            else:
                error = f"HTTP error {exc.code}: {exc.reason}"
        except urllib.error.URLError as exc:
            error = f"Network error: {exc.reason}"
        except json.JSONDecodeError:
            error = "Failed to parse GitHub release response."
        except Exception as exc:  # pragma: no cover - defensive UI path
            error = f"Unexpected error: {exc}"

        return {
            "update_available": False,
            "latest_version": None,
            "current_version": self.current_version,
            "download_url": None,
            "download_name": None,
            "platform": self.platform_label,
            "release_notes": None,
            "html_url": None,
            "error": error,
        }

    def open_download_page(self, url: Optional[str] = None) -> None:
        """Open a direct asset link or the latest release page."""
        webbrowser.open(url or f"https://github.com/{self.github_repo}/releases/latest")

    def _resolve_download_asset(self, assets: list[dict]) -> Optional[dict]:
        """Choose the most relevant release asset for the current platform."""
        candidate_checks: list = []
        if sys.platform == "darwin":
            candidate_checks = [
                lambda name: name.endswith(".dmg") and ("mac" in name or "macos" in name),
                lambda name: name.endswith(".dmg"),
                lambda name: name.endswith(".pkg"),
                lambda name: name.endswith(".zip") and ("mac" in name or "macos" in name),
            ]
        elif sys.platform == "win32":
            candidate_checks = [
                lambda name: name.endswith(".exe") and ("windows" in name or "win" in name),
                lambda name: name.endswith(".exe"),
                lambda name: name.endswith(".msi"),
                lambda name: name.endswith(".zip") and ("windows" in name or "win" in name),
            ]
        else:
            candidate_checks = [
                lambda name: name.endswith(".appimage"),
                lambda name: name.endswith(".tar.gz"),
                lambda name: name.endswith(".zip"),
            ]

        for check in candidate_checks:
            for asset in assets:
                asset_name = str(asset.get("name", "")).lower()
                if check(asset_name):
                    return asset

        return None

    @staticmethod
    def _get_platform_label() -> str:
        """Return a user-facing platform label for update prompts."""
        if sys.platform == "darwin":
            return "macOS"
        if sys.platform == "win32":
            return "Windows"
        return "this system"

    def _resolve_release_version(self, release_data: dict, assets: list[dict]) -> str:
        """Extract semantic version from tag, release name, or asset names."""
        for candidate in (
            release_data.get("tag_name", ""),
            release_data.get("name", ""),
            *[asset.get("name", "") for asset in assets],
        ):
            version = self._extract_semver(str(candidate))
            if version:
                return version
        return str(release_data.get("tag_name", "")).lstrip("v") or self.current_version

    def _is_newer_version(self, latest_version: str) -> bool:
        """Compare semantic-ish versions without external dependencies."""
        return self._version_key(latest_version) > self._version_key(self.current_version)

    @staticmethod
    def _extract_semver(text: str) -> Optional[str]:
        match = re.search(r"(\d+\.\d+\.\d+)", text)
        return match.group(1) if match else None

    @staticmethod
    def _version_key(version_text: str) -> tuple[int, ...]:
        match = re.findall(r"\d+", version_text)
        return tuple(int(part) for part in match[:3]) if match else (0, 0, 0)