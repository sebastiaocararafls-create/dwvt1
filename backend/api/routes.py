from __future__ import annotations

from typing import Any, Dict, List, Tuple, Optional

import pandas as pd
from fastapi import APIRouter, HTTPException, Request

from ..core.excel_repo import append_produto_excel
from ..core import solar_core as sc
from .schemas import CalcularRequest, CadastroRequest, CalcularResponse

router = APIRouter()


def _get_store(request: Request):
    store = getattr(request.app.state, "store", None)
    if store is None:
        raise HTTPException(status_code=500, detail="Store não inicializado no app.")
    return store


def _get_dfs(request: Request) -> Tuple[pd.DataFrame, pd.DataFrame]:
    store = _get_store(request)
    store.ensure_loaded()
    if store.df_inv is None or store.df_mod is None:
        raise HTTPException(status_code=500, detail=store.last_error or "Falha ao carregar Excel.")
    return store.df_inv, store.df_mod


def _inv_by_modelo(df_inv: pd.DataFrame, modelo: str) -> Dict[str, Any]:
    df = df_inv[df_inv[sc.CI_MODELO] == modelo]
    if df.empty:
        raise HTTPException(status_code=404, detail="Inversor não encontrado.")
    return df.iloc[0].to_dict()


def _mod_by_modelo(df_mod: pd.DataFrame, modelo: str) -> Dict[str, Any]:
    df = df_mod[df_mod[sc.CM_MODELO] == modelo]
    if df.empty:
        raise HTTPException(status_code=404, detail="Módulo não encontrado.")
    return df.iloc[0].to_dict()


@router.get("/health")
def health(request: Request):
    store = _get_store(request)
    # não força load aqui; só mostra status
    return {
        "ok": True,
        "excel_path": str(store.excel_path),
        "loaded": (store.df_inv is not None and store.df_mod is not None),
        "last_error": store.last_error,
    }


@router.post("/reload")
def reload_data(request: Request):
    store = _get_store(request)
    store.load()
    return {"ok": store.last_error is None, "last_error": store.last_error}


# ------------------------
# Listagens
# ------------------------

@router.get("/modulos/fabricantes")
def mod_fabricantes(request: Request) -> List[str]:
    _, df_mod = _get_dfs(request)
    return sorted(df_mod[sc.CM_FABRIC].dropna().unique().tolist())


@router.get("/modulos")
def mod_modelos(request: Request, fabricante: str) -> List[str]:
    _, df_mod = _get_dfs(request)
    df = df_mod[df_mod[sc.CM_FABRIC] == fabricante]
    return df[sc.CM_MODELO].dropna().tolist()


@router.get("/inversores/fabricantes")
def inv_fabricantes(request: Request) -> List[str]:
    df_inv, _ = _get_dfs(request)
    return sorted(df_inv[sc.CI_FABRIC].dropna().unique().tolist())


@router.get("/inversores")
def inv_modelos(request: Request, fabricante: str) -> List[str]:
    df_inv, _ = _get_dfs(request)
    df = df_inv[df_inv[sc.CI_FABRIC] == fabricante]
    return df[sc.CI_MODELO].dropna().tolist()


# ------------------------
# Cálculo
# ------------------------

@router.post("/calcular", response_model=CalcularResponse)
def calcular(request: Request, req: CalcularRequest):
    if req.tmin >= req.tmax:
        raise HTTPException(status_code=400, detail="Temperaturas inválidas: tmin deve ser < tmax.")

    df_inv, df_mod = _get_dfs(request)

    # (Opcional) valida se o modelo pertence ao fabricante selecionado
    # sem travar caso a planilha tenha inconsistências:
    inv = _inv_by_modelo(df_inv, req.modelo_inv)
    mod = _mod_by_modelo(df_mod, req.modelo_mod)

    ent_total = sc.entradas_total_inversor(inv)

    ent_usadas = req.entradas_usadas if req.entradas_usadas else ent_total
    ent_usadas = max(1, min(int(ent_usadas), ent_total))

    tam_forcado = int(req.tam_string) if req.tam_string else None

    mod_wp = sc.to_float(mod.get(sc.CM_POT, 0), 0.0)
    if mod_wp <= 0:
        raise HTTPException(status_code=400, detail="Módulo sem potência válida (Wp).")

    inv_pout_w = sc.to_float(inv.get(sc.CI_POT_SAI, 0), 0.0)
    if inv_pout_w <= 0:
        raise HTTPException(status_code=400, detail="Inversor sem potência de saída válida.")

    ovmax = sc.overload_max(inv)
    kwp_alvo_por_inv = req.kwp_sis / float(req.qtd_inv)

    corr = sc.calcular_correcoes(mod, req.tmin, req.tmax)
    lim = sc.limites_string(inv, corr)

    if lim["n_min"] > lim["n_max"]:
        return CalcularResponse(
            ok=False,
            motivo="Sem intervalo de string viável para esse módulo/inversor nas temperaturas informadas.",
            intervalo_string=lim,
        )

    combos = sc.gerar_combinacoes(
        kwp_alvo_por_inv=kwp_alvo_por_inv,
        mod_wp=mod_wp,
        inv_pout_w=inv_pout_w,
        ov_max=ovmax,
        entradas_total=ent_total,
        entradas_usadas=ent_usadas,
        n_min=lim["n_min"],
        n_max=lim["n_max"],
        tam_forcado=tam_forcado,
        max_items=8000,
    )

    if not combos:
        return CalcularResponse(
            ok=False,
            motivo="Nenhuma combinação válida encontrada com as restrições atuais.",
            intervalo_string=lim,
        )

    melhor = combos[0]
    criterios_melhor = sc.avaliar_criterios(
        inv=inv,
        corr=corr,
        n_min=lim["n_min"],
        n_max=lim["n_max"],
        combo=melhor,
        entradas_total=ent_total,
        entradas_usadas=ent_usadas,
    )

    return CalcularResponse(
        ok=True,
        melhor=melhor,
        combos=combos[:1200],
        criterios_melhor=criterios_melhor,
        intervalo_string=lim,
        ent_total=ent_total,
        ent_usadas=ent_usadas,
        ovmax=ovmax,
        kwp_alvo_por_inv=kwp_alvo_por_inv,
        pot_sis_melhor=melhor["pot"] * req.qtd_inv,
        inv=inv,          # dicionário completo do inversor (linha do Excel)
        mod=mod,          # dicionário completo do módulo (linha do Excel)
        correcoes={
            "voc_corrigida": corr["voc_tmin"],   # Voc corrigida (Tmin)
            "isc_corrigida": corr["isc_tmax"],   # Isc corrigida (Tmax)
            "vmp_corrigida": corr["vmp_tmax"],   # Vmp corrigida (Tmax) — boa p/ partida
        },
    )


# ------------------------
# Cadastro (escreve no Excel)
# ------------------------

@router.post("/cadastro/modulo")
def cadastro_modulo(request: Request, req: CadastroRequest):
    store = _get_store(request)
    try:
        append_produto_excel(store.excel_path, "MODULO", req.dados)
        store.load()
        return {"ok": True}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.post("/cadastro/inversor")
def cadastro_inversor(request: Request, req: CadastroRequest):
    store = _get_store(request)
    try:
        append_produto_excel(store.excel_path, "INVERSOR", req.dados)
        store.load()
        return {"ok": True}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))