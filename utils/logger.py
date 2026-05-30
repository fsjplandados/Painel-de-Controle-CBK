"""
Logger padronizado para o projeto.
"""
import logging
import sys


def get_logger(name: str, level: int = logging.INFO) -> logging.Logger:
    """Retorna um logger com handler de console formatado."""
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger  # já configurado

    logger.setLevel(level)
    handler = logging.StreamHandler(sys.stdout)
    handler.setLevel(level)
    formatter = logging.Formatter(
        fmt="%(asctime)s [%(name)s] %(levelname)s — %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.propagate = False
    return logger
