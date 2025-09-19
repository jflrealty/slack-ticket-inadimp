from models import OrdemServicoServicos
from database import Base, engine

# Cria a tabela (somente se não existir)
Base.metadata.create_all(bind=engine)
print("✅ Tabela tabela-os-servico criada (ou já existia)")
