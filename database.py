from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
import os

# Pega a URL do banco do Railway
DATABASE_URL = os.getenv("DATABASE_URL")

# Conexão
engine = create_engine(DATABASE_URL)

# Base para os models
Base = declarative_base()

# Sessão
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
