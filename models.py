from sqlalchemy import Column, Integer, String, DateTime, Float, Text
from database import Base
from datetime import datetime

class OrdemServicoFinanceiro(Base):
    __tablename__ = "ordens_servico_financeiro"

    id = Column(Integer, primary_key=True, index=True)
    tipo_ticket = Column(String)
    tipo_contrato = Column(String)
    locatario = Column(String)
    empreendimento_unidade = Column(String)
    numero_reserva = Column(String)
    responsavel = Column(String)
    solicitante = Column(String)
    status = Column(String, default="aberto")
    data_abertura = Column(DateTime, default=datetime.utcnow)
    data_captura = Column(DateTime, nullable=True)
    data_fechamento = Column(DateTime, nullable=True)
    data_ultima_edicao = Column(DateTime, nullable=True)
    ultimo_editor = Column(String, nullable=True)
    motivo_cancelamento = Column(Text, nullable=True)
    log_edicoes = Column(Text, nullable=True)
    historico_reaberturas = Column(Text, nullable=True)
    thread_ts = Column(String)
    canal_id = Column(String)
    capturado_por = Column(String, nullable=True)
