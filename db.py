# -*- coding: utf-8 -*-
"""
db.py — Camada de acesso ao banco Supabase.
Se as credenciais não estiverem configuradas, cai para JSON local (core.py).
"""
import os, json
from pathlib import Path

try:
    from supabase import create_client, Client
    _SUPABASE_OK = True
except ImportError:
    _SUPABASE_OK = False

# Credenciais — lidas do ambiente ou do streamlit secrets
def _get_creds():
    url = os.environ.get("SUPABASE_URL", "")
    key = os.environ.get("SUPABASE_KEY", "")
    if not url or not key:
        try:
            import streamlit as st
            url = st.secrets.get("SUPABASE_URL", "")
            key = st.secrets.get("SUPABASE_KEY", "")
        except Exception:
            pass
    return url.strip(), key.strip()

def _client() -> "Client | None":
    if not _SUPABASE_OK:
        return None
    url, key = _get_creds()
    if not url or not key:
        return None
    try:
        return create_client(url, key)
    except Exception:
        return None

# ── CRUD principal ────────────────────────────────────────────────────────────
def db_load() -> list:
    """Carrega todas as viagens do Supabase. Fallback para JSON local."""
    sb = _client()
    if sb is None:
        from core import load_data
        return load_data()
    try:
        res = sb.table("viagens").select("*").order("data_registro", desc=True).execute()
        return res.data or []
    except Exception as e:
        print(f"[db] Erro ao carregar: {e} — usando JSON local")
        from core import load_data
        return load_data()

def db_save(reg: dict) -> dict:
    """Insere ou atualiza um registro no Supabase."""
    sb = _client()
    if sb is None:
        from core import save_data, load_data
        dados = load_data()
        idx = next((i for i,d in enumerate(dados) if d.get("id")==reg.get("id")), None)
        if idx is not None:
            dados[idx] = reg
        else:
            dados.append(reg)
        save_data(dados)
        return reg
    try:
        sb.table("viagens").upsert(reg).execute()
        return reg
    except Exception as e:
        print(f"[db] Erro ao salvar: {e}")
        return reg

def db_delete(reg_id: str) -> bool:
    """Remove um registro pelo id."""
    sb = _client()
    if sb is None:
        from core import load_data, save_data
        dados = [d for d in load_data() if d.get("id") != reg_id]
        return save_data(dados)
    try:
        sb.table("viagens").delete().eq("id", reg_id).execute()
        return True
    except Exception as e:
        print(f"[db] Erro ao deletar: {e}")
        return False

def db_migrar_json() -> int:
    """Migra registros do viagens.json local para o Supabase. Retorna quantidade migrada."""
    from core import load_data
    dados = load_data()
    if not dados:
        return 0
    sb = _client()
    if sb is None:
        return 0
    try:
        sb.table("viagens").upsert(dados).execute()
        return len(dados)
    except Exception as e:
        print(f"[db] Erro na migração: {e}")
        return 0

def db_disponivel() -> bool:
    return _client() is not None
