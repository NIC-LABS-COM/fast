"""Exceções de domínio para classificação de erros no consumo de mensagens.

Permitem distinguir falhas transitórias (retry) de falhas definitivas (DLQ).
Os handlers de mensagem podem lançar essas exceções para sinalizar ao
mecanismo de retry como tratar o erro — sem acoplamento com infraestrutura.
"""


class TransientError(Exception):
    """Erro transitório — pode ser resolvido com retry.

    Exemplos: timeout do SAP GUI, falha temporária de rede,
    sessão SAP ocupada.
    """


class PermanentError(Exception):
    """Erro definitivo — NÃO deve ser retentado.

    Exemplos: payload inválido, VBS ausente, routing key desconhecida,
    erro de negócio irrecuperável.
    A mensagem será enviada diretamente para a DLQ.
    """
