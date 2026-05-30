"""
Configurações do projeto — lê credenciais de variáveis de ambiente ou .env.
"""
from dataclasses import dataclass, field
import os


def _env(key: str, default: str = "") -> str:
    return os.environ.get(key, default)


@dataclass
class VTEXConfig:
    """Credenciais e parâmetros da VTEX Orders API."""
    account:         str = field(default_factory=lambda: _env("VTEX_ACCOUNT", "sjdigital"))
    app_key:         str = field(default_factory=lambda: _env("VTEX_APP_KEY",   "vtexappkey-sjdigital-NBIBYX"))
    app_token:       str = field(default_factory=lambda: _env("VTEX_APP_TOKEN", "ZWWMCOPAPYMWRDDFJXJASHHUYAHMNWFDLQKYEFYTGNOHDWBDBJGDWDRAQKGALTKTJZUTNMSEOSARVFCIQDNTEVGACYJBFYYKDFRYJTFSQJTOANANWPYYWISDULGXVMON"))
    timeout_seconds: int = field(default_factory=lambda: int(_env("VTEX_TIMEOUT", "30")))


@dataclass
class AppConfig:
    """Configuração geral da aplicação."""
    vtex: VTEXConfig = field(default_factory=VTEXConfig)
