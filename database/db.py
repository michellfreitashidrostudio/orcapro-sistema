import os
from datetime import datetime, date
import streamlit as st
from sqlalchemy import (
    create_engine, Column, Integer, String, Float, Boolean,
    Date, DateTime, ForeignKey, Text, UniqueConstraint, text
)
from sqlalchemy.orm import declarative_base, relationship, sessionmaker

# Busca o endereço do Neon nos Secrets da nuvem; se rodar local, usa o SQLite
try:
    DATABASE_URL = st.secrets["DATABASE_URL"]
except Exception:
    DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///orcamentos.db")

# Garante compatibilidade com SQLAlchemy 2.0
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

connect_args = {"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {}
engine = create_engine(
    DATABASE_URL,
    connect_args=connect_args,
    pool_pre_ping=True,  # Reconecta automaticamente quando o Neon acorda
    pool_recycle=300
)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

class BaseReferencia(Base):
    __tablename__ = "bases_referencia"
    id = Column(Integer, primary_key=True, index=True)
    nome = Column(String(100), nullable=False)
    descricao = Column(String(255))
    uf = Column(String(2), default="SP")
    data_base = Column(Date, default=date.today)
    desonerado = Column(Boolean, default=False)
    criado_em = Column(DateTime, default=datetime.utcnow)

    composicoes = relationship("Composicao", back_populates="base", cascade="all, delete-orphan")
    __table_args__ = (UniqueConstraint("nome", "uf", "data_base", "desonerado", name="uq_base_ref"),)

class Composicao(Base):
    __tablename__ = "composicoes"
    id = Column(Integer, primary_key=True, index=True)
    base_id = Column(Integer, ForeignKey("bases_referencia.id", ondelete="CASCADE"), nullable=False)
    codigo = Column(String(50), nullable=False, index=True)
    descricao = Column(Text, nullable=False)
    unidade = Column(String(20), nullable=False)
    preco_unitario = Column(Float, nullable=False)

    base = relationship("BaseReferencia", back_populates="composicoes")
    __table_args__ = (UniqueConstraint("base_id", "codigo", name="uq_base_codigo"),)

class PerfilBDI(Base):
    __tablename__ = "perfis_bdi"
    id = Column(Integer, primary_key=True, index=True)
    nome = Column(String(100), nullable=False, unique=True)
    ac = Column(Float, default=0.04)
    sg = Column(Float, default=0.008)
    r = Column(Float, default=0.0127)
    df = Column(Float, default=0.0123)
    l = Column(Float, default=0.074)
    t = Column(Float, default=0.0865)
    percentual_final = Column(Float, nullable=False)

class Orcamento(Base):
    __tablename__ = "orcamentos"
    id = Column(Integer, primary_key=True, index=True)
    titulo = Column(String(255), nullable=False)
    cliente = Column(String(255), default="")
    local_obra = Column(String(255), default="São Paulo - SP")
    bdi_padrao_id = Column(Integer, ForeignKey("perfis_bdi.id"), nullable=True)
    truncar_valores = Column(Boolean, default=True)
    desonerado = Column(Boolean, default=False)
    criado_em = Column(DateTime, default=datetime.utcnow)

    bdi_padrao = relationship("PerfilBDI")
    itens = relationship("OrcamentoItem", back_populates="orcamento", cascade="all, delete-orphan")

class OrcamentoItem(Base):
    __tablename__ = "orcamento_itens"
    id = Column(Integer, primary_key=True, index=True)
    orcamento_id = Column(Integer, ForeignKey("orcamentos.id", ondelete="CASCADE"), nullable=False)
    etapa = Column(String(150), nullable=False)
    item_eap = Column(String(30), nullable=False)
    fonte_base = Column(String(100), nullable=False)
    codigo = Column(String(50), nullable=False)
    descricao = Column(Text, nullable=False)
    unidade = Column(String(20), nullable=False)
    quantidade = Column(Float, nullable=False)
    preco_sem_bdi = Column(Float, nullable=False)
    bdi_aplicado = Column(Float, nullable=False)
    preco_com_bdi = Column(Float, nullable=False)
    preco_total = Column(Float, nullable=False)

    orcamento = relationship("Orcamento", back_populates="itens")

class Insumo(Base):
    __tablename__ = "insumos"
    id = Column(Integer, primary_key=True, index=True)
    codigo = Column(String(50), nullable=False, unique=True)
    descricao = Column(Text, nullable=False)
    tipo = Column(String(30), default="MATERIAL")
    unidade = Column(String(20), nullable=False)
    preco_unitario = Column(Float, nullable=False)
    fonte = Column(String(50), default="PROPRIA")

class Cadastro(Base):
    __tablename__ = "cadastros"
    id = Column(Integer, primary_key=True, index=True)
    tipo = Column(String(30), nullable=False)
    nome = Column(String(200), nullable=False)
    documento = Column(String(50), default="")
    contato = Column(String(150), default="")
    cidade_uf = Column(String(100), default="São Paulo - SP")

class Compra(Base):
    __tablename__ = "compras"
    id = Column(Integer, primary_key=True, index=True)
    orcamento_id = Column(Integer, ForeignKey("orcamentos.id", ondelete="CASCADE"), nullable=False)
    fornecedor = Column(String(200), default="")
    item_descricao = Column(Text, nullable=False)
    unidade = Column(String(20), default="UN")
    quantidade = Column(Float, default=1.0)
    valor_unitario = Column(Float, default=0.0)
    status = Column(String(30), default="Em Cotação")
    data_pedido = Column(Date, default=date.today)

class EncargoSocial(Base):
    __tablename__ = "encargos_sociais"
    id = Column(Integer, primary_key=True, index=True)
    nome = Column(String(120), nullable=False, unique=True)
    desonerado = Column(Boolean, default=True)
    horista_pct = Column(Float, nullable=False)
    mensalista_pct = Column(Float, nullable=False)
    observacoes = Column(Text, default="")

class MemoriaCalculo(Base):
    __tablename__ = "memoria_calculo"
    id = Column(Integer, primary_key=True, index=True)
    orcamento_item_id = Column(Integer, ForeignKey("orcamento_itens.id", ondelete="CASCADE"), nullable=False)
    descricao_trecho = Column(String(200), nullable=False)
    fator = Column(Float, default=1.0)
    comprimento = Column(Float, default=1.0)
    largura = Column(Float, default=1.0)
    altura = Column(Float, default=1.0)
    subtotal = Column(Float, nullable=False)

class CronogramaEtapa(Base):
    __tablename__ = "cronograma_etapas"
    id = Column(Integer, primary_key=True, index=True)
    orcamento_id = Column(Integer, ForeignKey("orcamentos.id", ondelete="CASCADE"), nullable=False)
    etapa = Column(String(150), nullable=False)
    mes = Column(Integer, nullable=False)
    percentual = Column(Float, default=0.0)
    __table_args__ = (UniqueConstraint("orcamento_id", "etapa", "mes", name="uq_crono_etapa_mes"),)

class DMTRegistro(Base):
    __tablename__ = "dmt_registros"
    id = Column(Integer, primary_key=True, index=True)
    orcamento_id = Column(Integer, ForeignKey("orcamentos.id", ondelete="CASCADE"), nullable=False)
    material = Column(String(150), nullable=False)
    origem = Column(String(150), default="")
    destino = Column(String(150), default="")
    km_pavimentado = Column(Float, default=0.0)
    km_nao_pavimentado = Column(Float, default=0.0)
    custo_km_pav = Column(Float, default=1.80)
    custo_km_nao_pav = Column(Float, default=2.60)
    volume_ton = Column(Float, default=1.0)

def init_db():
    Base.metadata.create_all(bind=engine)
    # Migração segura: adiciona a coluna 'desonerado' em bancos já existentes
    with engine.begin() as conn:
        try:
            conn.execute(text("ALTER TABLE orcamentos ADD COLUMN desonerado BOOLEAN DEFAULT 0"))
        except Exception:
            pass

    db = SessionLocal()
    if db.query(PerfilBDI).count() == 0:
        from core.calculations import calcular_bdi_tcu
        perc = calcular_bdi_tcu(0.04, 0.008, 0.0127, 0.0123, 0.074, 0.0865)
        db.add(PerfilBDI(nome="BDI Padrão - Obras Civis (TCU)", percentual_final=perc))
        db.commit()
    if db.query(EncargoSocial).count() == 0:
        db.add_all([
            EncargoSocial(nome="SINAPI SP - Desonerado", desonerado=True, horista_pct=84.73, mensalista_pct=47.76, observacoes="Padrão Caixa/SINAPI Desonerado"),
            EncargoSocial(nome="SINAPI SP - Não Desonerado", desonerado=False, horista_pct=114.21, mensalista_pct=71.36, observacoes="Padrão Caixa/SINAPI Não Desonerado"),
        ])
        db.commit()
    db.close()