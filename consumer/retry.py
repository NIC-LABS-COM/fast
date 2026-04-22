"""Mecanismo de retry com suporte a DLQ para callbacks RabbitMQ.

Equivalente Python do RetryInterceptorBuilder + RejectAndDontRequeueRecoverer
do Spring AMQP. Envolve callbacks com:
  - Manual ACK em caso de sucesso
  - Retry com backoff linear em caso de erro transitório
  - NACK (requeue=False) após esgotar tentativas → mensagem vai para DLQ
  - NACK imediato para erros permanentes (PermanentError)
"""

import time
import functools
import traceback

from .config import MAX_RETRY_ATTEMPTS, RETRY_BACKOFF_SECONDS
from .logger import log
from .errors import PermanentError


def with_retry_and_ack(max_retries=None, backoff=None):
    """Decorator que envolve um callback RabbitMQ com retry + manual ack/nack.

    O callback decorado DEVE ter a assinatura padrão do pika:
        callback(ch, method, properties, body)

    Fluxo:
        1. Executa o callback
        2. Se retornar normalmente → basic_ack (mensagem confirmada)
        3. Se lançar PermanentError → basic_nack(requeue=False) imediato (DLQ)
        4. Se lançar outra exceção → retry até max_retries
        5. Se todas as tentativas falharem → basic_nack(requeue=False) (DLQ)

    Parâmetros:
        max_retries: Número máximo de tentativas (padrão: MAX_RETRY_ATTEMPTS)
        backoff: Segundos base entre tentativas (padrão: RETRY_BACKOFF_SECONDS)
                 Backoff linear: wait = backoff * attempt_number
    """
    if max_retries is None:
        max_retries = MAX_RETRY_ATTEMPTS
    if backoff is None:
        backoff = RETRY_BACKOFF_SECONDS

    def decorator(callback_fn):
        @functools.wraps(callback_fn)
        def wrapper(ch, method, properties, body):
            delivery_tag = method.delivery_tag
            last_exception = None

            for attempt in range(1, max_retries + 1):
                try:
                    callback_fn(ch, method, properties, body)
                    # Sucesso — confirma a mensagem
                    ch.basic_ack(delivery_tag=delivery_tag)
                    return
                except PermanentError as exc:
                    # Erro permanente — rejeita sem requeue (vai direto para DLQ)
                    log(f"[RETRY] Erro PERMANENTE na tentativa "
                        f"{attempt}/{max_retries}: {exc}")
                    log(f"[RETRY] Mensagem rejeitada (DLQ). "
                        f"delivery_tag={delivery_tag}")
                    ch.basic_nack(delivery_tag=delivery_tag, requeue=False)
                    return
                except Exception as exc:
                    last_exception = exc
                    log(f"[RETRY] Tentativa {attempt}/{max_retries} falhou: "
                        f"{exc}")
                    if attempt < max_retries:
                        wait = backoff * attempt
                        log(f"[RETRY] Aguardando {wait}s antes da proxima "
                            f"tentativa...")
                        time.sleep(wait)

            # Todas as tentativas esgotadas — rejeita sem requeue (DLQ)
            # Equivalente ao RejectAndDontRequeueRecoverer do Spring
            log(f"[RETRY] TODAS as {max_retries} tentativas esgotadas. "
                f"Rejeitando mensagem (DLQ). delivery_tag={delivery_tag}")
            log(f"[RETRY] Ultimo erro: {last_exception}")
            log(f"[RETRY] Traceback: {traceback.format_exc()}")
            ch.basic_nack(delivery_tag=delivery_tag, requeue=False)

        return wrapper
    return decorator
