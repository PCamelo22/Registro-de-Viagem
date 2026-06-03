# -*- coding: utf-8 -*-
"""
db.py — Acesso ao Supabase via REST API (httpx).
Compatível com o novo formato de chaves sb_publishable_ / sb_secret_.
Fallback automático para JSON local se não configurado.
"""
import os, json

try:
    import httpx
    _HTTPX = True
except ImportError:
    _HTTPX = False

REST_SUFFIX = "/rest/v1"
HEADERS_BASE = {
    "Content-Type":  "application/json",
    "Prefer":        "return=representation",
}

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
    return url.strip().rstrip("/"), key.strip()

def _headers():
    _, key = _get_creds()
    return {**HEADERS_BASE, "apikey": key, "Authorization": f"Bearer {key}"}

def _base_url():
    url, _ = _get_creds()
    return f"{url}{REST_SUFFIX}"

def db_disponivel() -> bool:
    if not _HTTPX:
        return False
    url, key = _get_creds()
    return bool(url and key)

# ── CRUD ──────────────────────────────────────────────────────────────────────
def db_load() -> list:
    if not db_disponivel():
        from core import load_data
        return load_data()
    try:
        r = httpx.get(
            f"{_base_url()}/viagens",
            headers=_headers(),
            params={"order": "data_registro.desc"},
            timeout=10,
        )
        r.raise_for_status()
        return r.json() or []
    except Exception as e:
        print(f"[db] Erro ao carregar: {e} — usando JSON local")
        from core import load_data
        return load_data()

def db_save(reg: dict) -> dict:
    if not db_disponivel():
        from core import load_data, save_data
        dados = load_data()
        idx = next((i for i, d in enumerate(dados) if d.get("id") == reg.get("id")), None)
        if idx is not None:
            dados[idx] = reg
        else:
            dados.append(reg)
        save_data(dados)
        return reg
    try:
        hdrs = {**_headers(), "Prefer": "resolution=merge-duplicates,return=representation"}
        r = httpx.post(
            f"{_base_url()}/viagens",
            headers=hdrs,
            content=json.dumps(reg),
            timeout=10,
        )
        r.raise_for_status()
        return reg
    except Exception as e:
        print(f"[db] Erro ao salvar: {e}")
        return reg

def db_delete(reg_id: str) -> bool:
    if not db_disponivel():
        from core import load_data, save_data
        dados = [d for d in load_data() if d.get("id") != reg_id]
        return save_data(dados)
    try:
        r = httpx.delete(
            f"{_base_url()}/viagens",
            headers=_headers(),
            params={"id": f"eq.{reg_id}"},
            timeout=10,
        )
        r.raise_for_status()
        return True
    except Exception as e:
        print(f"[db] Erro ao deletar: {e}")
        return False

def db_migrar_json() -> int:
    from core import load_data
    dados = load_data()
    if not dados:
        return 0
    if not db_disponivel():
        return 0
    try:
        hdrs = {**_headers(), "Prefer": "resolution=merge-duplicates,return=minimal"}
        r = httpx.post(
            f"{_base_url()}/viagens",
            headers=hdrs,
            content=json.dumps(dados),
            timeout=30,
        )
        r.raise_for_status()
        return len(dados)
    except Exception as e:
        print(f"[db] Erro na migração: {e}")
        return 0

def db_testar() -> tuple[bool, str]:
    """Testa a conexão e retorna (ok, mensagem)."""
    if not _HTTPX:
        return False, "httpx não instalado (pip install httpx)"
    url, key = _get_creds()
    if not url or not key:
        return False, "SUPABASE_URL ou SUPABASE_KEY não configurados"
    try:
        r = httpx.get(
            f"{url}{REST_SUFFIX}/viagens",
            headers=_headers(),
            params={"limit": "1"},
            timeout=8,
        )
        if r.status_code == 200:
            return True, f"Conectado — {len(r.json())} registro(s) encontrado(s)"
        return False, f"HTTP {r.status_code}: {r.text[:200]}"
    except Exception as e:
        return False, str(e)
