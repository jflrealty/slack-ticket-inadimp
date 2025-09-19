from sqlalchemy import Column, Integer, String, DateTime
from database import Base
from datetime import datetime

class OrdemServicoServicos(Base):
    __tablename__ = "tabela-os-servico"  # nome exato da tabela criada no Railway

    id = Column(Integer, primary_key=True, index=True)
    tipo_ticket = Column(String, nullable=False)
    locatario = Column(String, nullable=False)
    empreendimento_unidade = Column(String, nullable=False)
    responsavel = Column(String, nullable=False)
    status = Column(String, default="Aberto")
    criado_em = Column(DateTime, default=datetime.utcnow)
    atualizado_em = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    thread_ts = Column(String, nullable=True)
    canal_id = Column(String, nullable=True)
