"""
Decorator de retry com backoff exponencial.
"""
import time
import functools
from typing import Tuple, Type


def with_retry(
    max_attempts: int = 3,
    backoff_seconds: float = 2.0,
    exceptions: Tuple[Type[Exception], ...] = (Exception,),
):
    """
    Decorator que reexecuta a função em caso de falha.

    Args:
        max_attempts:    Número máximo de tentativas (incluindo a primeira).
        backoff_seconds: Tempo de espera base entre tentativas (dobra a cada retry).
        exceptions:      Tupla de exceções que acionam o retry.
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            last_exc = None
            for attempt in range(1, max_attempts + 1):
                try:
                    return func(*args, **kwargs)
                except exceptions as exc:
                    last_exc = exc
                    if attempt < max_attempts:
                        wait = backoff_seconds * (2 ** (attempt - 1))
                        time.sleep(wait)
            raise last_exc
        return wrapper
    return decorator
