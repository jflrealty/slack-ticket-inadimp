from models import OrdemServicoFinanceiro
from database import Base, engine

# Cria a tabela (somente se não existir)
Base.metadata.create_all(bind=engine)
print("✅ Tabela ordens_servico_financeiro criada (ou já existia)")
