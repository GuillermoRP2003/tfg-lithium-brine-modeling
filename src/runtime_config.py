from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path

try:
    import tomllib
except ModuleNotFoundError:  # pragma: no cover - only used on Python < 3.11
    import tomli as tomllib


ROOT = Path(__file__).resolve().parents[1]
INPUTS = ROOT / "input"
RUNS = ROOT / "runs"
RESULTS = ROOT / "results"
VIEWER_ASSETS = ROOT / "viewer_assets"

CONFIG_LOCAL = ROOT / "config.local.toml"
CONFIG_EXAMPLE = ROOT / "config.example.toml"


@dataclass(frozen=True)
class RuntimeConfig:
    root: Path
    inputs: Path
    runs: Path
    results: Path
    viewer_assets: Path
    phreeqc_exe: Path | None
    phreeqc_database: Path | None


def _load_local_config() -> dict:
    if not CONFIG_LOCAL.exists():
        return {}
    with CONFIG_LOCAL.open("rb") as file:
        return tomllib.load(file)


def _path_from_value(value: str | os.PathLike | None) -> Path | None:
    if value is None:
        return None
    text = os.path.expandvars(str(value).strip())
    if not text:
        return None
    path = Path(text).expanduser()
    if not path.is_absolute():
        path = ROOT / path
    return path


def _configured_path(env_var: str, config: dict, key: str) -> Path | None:
    env_value = os.environ.get(env_var)
    if env_value:
        return _path_from_value(env_value)
    return _path_from_value(config.get("phreeqc", {}).get(key))


def load_runtime_config() -> RuntimeConfig:
    config = _load_local_config()
    return RuntimeConfig(
        root=ROOT,
        inputs=INPUTS,
        runs=RUNS,
        results=RESULTS,
        viewer_assets=VIEWER_ASSETS,
        phreeqc_exe=_configured_path("PHREEQC_EXE", config, "exe"),
        phreeqc_database=_configured_path("PHREEQC_DATABASE", config, "database"),
    )


CONFIG = load_runtime_config()
PHREEQC_EXE = CONFIG.phreeqc_exe
DATABASE = CONFIG.phreeqc_database


def phreeqc_configuration_help() -> str:
    return (
        "Configure PHREEQC mediante variables de entorno PHREEQC_EXE y "
        "PHREEQC_DATABASE, o cree config.local.toml a partir de "
        "config.example.toml. PHREEQC no se incluye en el repositorio."
    )
