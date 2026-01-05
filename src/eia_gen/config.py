import os
from pathlib import Path

from pydantic import Field
from pydantic.aliases import AliasChoices
from pydantic_settings import BaseSettings, SettingsConfigDict


def _default_env_files() -> tuple[Path, ...]:
    """Return dotenv candidates for pydantic-settings (missing files are ignored).

    Goal: make `eia-gen` runnable from any CWD while still allowing per-CWD overrides.
    Precedence (later overrides earlier):
    1) repo-root `.env`, `.env.local` (where this package lives during dev)
    2) CWD `.env`, `.env.local` (user overrides)
    """

    def _repo_root() -> Path:
        here = Path(__file__).resolve()
        # Typical dev layout: <repo>/src/eia_gen/config.py
        for cand in [here.parent] + list(here.parents):
            if (cand / "pyproject.toml").exists() and (cand / "src").exists():
                return cand
        # Fallback: 2 levels up from package dir.
        try:
            return here.parents[2]
        except Exception:
            return here.parent

    repo_root = _repo_root()
    cwd = Path.cwd()
    return (
        repo_root / ".env",
        repo_root / ".env.local",
        cwd / ".env",
        cwd / ".env.local",
    )


class Settings(BaseSettings):
    # Prefer local env files when present (developer convenience).
    # - `.env` / `.env.local` are optional; missing files are ignored.
    # - Environment variables still override values from files.
    model_config = SettingsConfigDict(
        env_prefix="EIA_GEN_",
        extra="ignore",
        env_file=_default_env_files(),
        env_file_encoding="utf-8",
    )

    openai_api_key: str | None = Field(
        default=None,
        validation_alias=AliasChoices("EIA_GEN_OPENAI_API_KEY", "OPENAI_API_KEY"),
    )
    openai_model: str = "gpt-4.1-mini"

    # Public data / map API keys (optional)
    data_go_kr_service_key: str | None = Field(
        default=None,
        validation_alias=AliasChoices("EIA_GEN_DATA_GO_KR_SERVICE_KEY", "DATA_GO_KR_SERVICE_KEY"),
    )
    safemap_api_key: str | None = Field(
        default=None,
        validation_alias=AliasChoices("EIA_GEN_SAFEMAP_API_KEY", "SAFEMAP_API_KEY"),
    )
    vworld_api_key: str | None = Field(
        default=None,
        validation_alias=AliasChoices("EIA_GEN_VWORLD_API_KEY", "VWORLD_API_KEY"),
    )
    kma_api_key: str | None = Field(
        default=None,
        validation_alias=AliasChoices("EIA_GEN_KMA_API_KEY", "KMA_API_KEY"),
    )
    airkorea_api_key: str | None = Field(
        default=None,
        validation_alias=AliasChoices("EIA_GEN_AIRKOREA_API_KEY", "AIRKOREA_API_KEY"),
    )
    kosis_api_key: str | None = Field(
        default=None,
        validation_alias=AliasChoices("EIA_GEN_KOSIS_API_KEY", "KOSIS_API_KEY"),
    )

    # DOCX rendering
    # When false (default), we strip inline `〔SRC:...〕` blocks from visible text to
    # maximize "sample PDF" layout fidelity, while keeping traceability in source_register.xlsx.
    docx_render_citations: bool = False

    # When true, template mode becomes fail-fast:
    # - missing anchors/table placeholders become hard errors
    # - prevents silent fallback that can drift from the sample PDF style
    docx_strict_template: bool = False

    # Deliverables (optional)
    # When set, generated reports are copied to this directory for easy handoff (e.g., ~/Desktop).
    deliverables_dir: str | None = None
    # Optional tag used in deliverable filenames (fallback: case_id or case folder name).
    deliverables_tag: str | None = None


settings = Settings()

# Make `.env`-loaded keys visible to modules that still read from `os.environ`.
_EXPORTS = {
    "DATA_GO_KR_SERVICE_KEY": settings.data_go_kr_service_key,
    "SAFEMAP_API_KEY": settings.safemap_api_key,
    "VWORLD_API_KEY": settings.vworld_api_key,
    "KMA_API_KEY": settings.kma_api_key,
    "AIRKOREA_API_KEY": settings.airkorea_api_key,
    "KOSIS_API_KEY": settings.kosis_api_key,
}
for k, v in _EXPORTS.items():
    if v and not os.environ.get(k, "").strip():
        os.environ[k] = v
