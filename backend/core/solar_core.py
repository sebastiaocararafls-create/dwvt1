from __future__ import annotations

import math
from typing import Optional, Dict, Any, List, Tuple

TEMP_STC = 25.0

# INVERSORES (nomes conforme Excel)
CI_MODELO = "MODELO"
CI_FABRIC = "FABRICANTE"
CI_POT_ENT = "MAX. POT. ENTRADA (KWP)"
CI_N_ENT = "Nº ENTRADAS INVERSOR"
CI_N_MPPT = "Nº MPPTS"
CI_POT_SAI = "MAX. POT. SAÍDA (KW)"
CI_VAC = "TENSÃO SAÍDA (V)"
CI_VMAX = "TENSÃO MÁX. ENTRADA (V)"
CI_VMPP_MAX = "TENSÃO MÁX. MPP (V)"
CI_VMIN_PART = "TENSÃO MÍN PARTIDA. (V)"
CI_IMAX = "CORRENTE MÁX. ENTRADA FV (A)"
CI_ICC = "CORRENTE MÁX. CURTO (A)"
CI_CFG_MPPT = "CONFIGURAÇÃO MPPTs"

# MODULOS
CM_MODELO = "MODELO"
CM_FABRIC = "FABRICANTE"
CM_POT = "POTÊNCIA (KWP)"  # na planilha: Wp (ex.: 600)
CM_TIPO = "TIPO CÉLULA"
CM_VOC = "TENSÃO ABERTO (Voc)"
CM_VMP = "TENSÃO OP. STC (Vmp)"
CM_IMP = "CORRENTE (A)"
CM_ISC = "CORRENTE CURTO (A)"
CM_EFIC = "EFICIÊNCIA (%)"
CM_COEF_ISC = "COEF. TEMP. CORR. CURTO (%/°C)"
CM_COEF_VOC = "COEF. TEMP. TENSÃO ABERTO (%/°C)"
CM_BIFACIAL = "MONO/BIFACIAL"


# ============================================================
# UTILITÁRIOS
# ============================================================

def to_float(v, default=0.0) -> float:
    try:
        if v is None:
            return float(default)
        if isinstance(v, (int, float)):
            f = float(v)
            return float(default) if math.isnan(f) else f
        s = str(v).strip()
        if s == "" or s.lower() == "nan":
            return float(default)
        s = s.replace(",", ".")
        # pega o primeiro número em strings tipo "20/20 22,5/22,5"
        for sep in ("/", " "):
            if sep in s:
                s = s.split(sep)[0].strip()
                break
        return float(s)
    except Exception:
        return float(default)

def to_int(v, default=0) -> int:
    try:
        return int(float(str(v).replace(",", ".")))
    except Exception:
        return int(default)

def parse_config_mppt(config_str: str) -> List[int]:
    """
    Retorna lista de entradas por MPPT.
    Ex.: "2/2" -> [2,2] (total 4 entradas)
         "1/1/1" -> [1,1,1]
         "14" -> [14]
    """
    if not config_str:
        return []
    s = str(config_str).strip()
    if s == "" or s.lower() == "nan":
        return []
    out = []
    for part in s.split("/"):
        part = part.strip()
        if part == "":
            continue
        if " " in part:
            part = part.split(" ")[0].strip()
        part = part.replace(",", ".")
        try:
            n = int(float(part))
            if n > 0:
                out.append(n)
        except Exception:
            pass
    return out

def entradas_total_inversor(inv: Dict[str, Any]) -> int:
    cfg = str(inv.get(CI_CFG_MPPT, "")).strip()
    parsed = parse_config_mppt(cfg)
    if parsed:
        return max(1, sum(parsed))
    n = to_int(inv.get(CI_N_ENT, 0), 0)
    return n if n > 0 else 1

def overload_max(inv: Dict[str, Any]) -> float:
    pin = to_float(inv.get(CI_POT_ENT, 0), 0.0)
    pout = to_float(inv.get(CI_POT_SAI, 0), 0.0)
    if pin <= 0 or pout <= 0:
        return 1.8
    # ambos tipicamente em W/Wp na sua planilha
    return max(pin / pout, 1.0)


# ============================================================
# CÁLCULOS
# ============================================================

def calcular_correcoes(mod: Dict[str, Any], t_min: float, t_max: float) -> Dict[str, float]:
    voc = to_float(mod.get(CM_VOC, 0), 0.0)
    vmp = to_float(mod.get(CM_VMP, 0), 0.0)
    isc = to_float(mod.get(CM_ISC, 0), 0.0)
    imp = to_float(mod.get(CM_IMP, 0), 0.0)

    coef_voc = to_float(mod.get(CM_COEF_VOC, -0.25), -0.25) / 100.0
    coef_isc = to_float(mod.get(CM_COEF_ISC, 0.043), 0.043) / 100.0

    dt_min = t_min - TEMP_STC
    dt_max = t_max - TEMP_STC

    voc_tmin = voc * (1 + coef_voc * dt_min)
    vmp_tmin = vmp * (1 + coef_voc * dt_min)
    vmp_tmax = vmp * (1 + coef_voc * dt_max)

    isc_tmax = isc * (1 + coef_isc * dt_max)
    imp_corr = imp * (1 + coef_isc * dt_max)

    return {
        "voc": voc, "vmp": vmp, "isc": isc, "imp": imp,
        "voc_tmin": voc_tmin,
        "vmp_tmin": vmp_tmin,
        "vmp_tmax": vmp_tmax,
        "isc_tmax": isc_tmax,
        "imp_corr": imp_corr,
        "dt_min": dt_min,
        "dt_max": dt_max
    }

def limites_string(inv: Dict[str, Any], corr: Dict[str, float]) -> Dict[str, int]:
    vmax = to_float(inv.get(CI_VMAX, 1000), 1000)
    vmpp_max = to_float(inv.get(CI_VMPP_MAX, 1100), 1100)
    vstart = to_float(inv.get(CI_VMIN_PART, 60), 60)

    voc_tmin = corr["voc_tmin"]
    vmp_tmin = corr["vmp_tmin"]
    vmp_tmax = corr["vmp_tmax"]

    n_max_voc = math.floor(vmax / voc_tmin) if voc_tmin > 0 else 0
    n_max_mpp = math.floor(vmpp_max / vmp_tmin) if vmp_tmin > 0 else 0
    n_min_mpp = math.ceil(vstart / vmp_tmax) if vmp_tmax > 0 else 1

    n_min = max(1, n_min_mpp)
    n_max = min(n_max_voc, n_max_mpp)

    return {
        "n_min": int(n_min),
        "n_max": int(n_max),
        "n_max_voc": int(n_max_voc),
        "n_max_mpp": int(n_max_mpp),
        "n_min_mpp": int(n_min_mpp),
    }

def gerar_combinacoes(
    kwp_alvo_por_inv: float,
    mod_wp: float,
    inv_pout_w: float,
    ov_max: float,
    entradas_total: int,
    entradas_usadas: int,
    n_min: int,
    n_max: int,
    tam_forcado: Optional[int] = None,
    max_items: int = 8000,
) -> List[Dict[str, float]]:
    """
    Gera combinações por inversor:
      - tamanho (mód/string)
      - qtd (strings = entradas utilizadas)
      - total módulos
      - pot (kWp DC por inversor)
      - dif (kWp)
      - overload (fator)
    """
    mod_kwp = mod_wp / 1000.0
    if mod_kwp <= 0:
        return []

    ent = max(1, min(entradas_usadas, entradas_total))

    if tam_forcado is not None:
        if tam_forcado < n_min or tam_forcado > n_max:
            return []
        tamanhos = [tam_forcado]
    else:
        tamanhos = list(range(n_min, n_max + 1))

    out = []
    for tam in tamanhos:
        for qtd in range(1, ent + 1):
            total = tam * qtd
            pot = total * mod_kwp
            dif = abs(pot - kwp_alvo_por_inv)
            overload = (pot * 1000.0) / inv_pout_w if inv_pout_w > 0 else 0.0
            if overload > ov_max + 1e-9:
                continue
            out.append({
                "tamanho": float(tam),
                "qtd": float(qtd),
                "total": float(total),
                "pot": float(pot),
                "dif": float(dif),
                "overload": float(overload),
            })
            if len(out) >= max_items:
                break
        if len(out) >= max_items:
            break

    out.sort(key=lambda r: (r["dif"], -r["overload"], r["total"]))
    return out

def avaliar_criterios(
    inv: Dict[str, Any],
    corr: Dict[str, float],
    n_min: int,
    n_max: int,
    combo: Dict[str, float],
    entradas_total: int,
    entradas_usadas: int,
) -> List[Tuple[str, str]]:
    """
    Retorna lista de (nome, status) onde status ∈ {"APROVADO","ATENÇÃO","REPROVADO"}
    """
    tam = int(combo["tamanho"])
    qtd = int(combo["qtd"])

    vmax = to_float(inv.get(CI_VMAX, 0), 0.0)
    vmpp_max = to_float(inv.get(CI_VMPP_MAX, 0), 0.0)
    vstart = to_float(inv.get(CI_VMIN_PART, 0), 0.0)
    imax = to_float(inv.get(CI_IMAX, 0), 0.0)
    icc = to_float(inv.get(CI_ICC, 0), 0.0)

    voc_max = corr["voc_tmin"] * tam
    vmp_min_oper = corr["vmp_tmax"] * tam
    vmp_max_oper = corr["vmp_tmin"] * tam
    imp_corr = corr["imp_corr"]
    isc_corr = corr["isc_tmax"]

    ovmax = overload_max(inv)
    inv_pout_w = to_float(inv.get(CI_POT_SAI, 0), 0.0)
    pot_kwp = combo["pot"]
    ov = (pot_kwp * 1000.0) / inv_pout_w if inv_pout_w > 0 else 9e9

    ent = max(1, min(entradas_usadas, entradas_total))
    ok_strings = qtd <= ent

    def ok(cond: bool) -> str:
        return "APROVADO" if cond else "REPROVADO"

    def ok_current(i: float, lim: float) -> str:
        if lim <= 0:
            return "ATENÇÃO"
        if i <= lim * 0.95:
            return "APROVADO"
        if i <= lim:
            return "ATENÇÃO"
        return "REPROVADO"

    criterios = []
    criterios.append(("Tensão Máxima CC", ok(voc_max <= vmax if vmax > 0 else True)))
    criterios.append(("Tensão de partida CC", ok(vmp_min_oper >= vstart if vstart > 0 else True)))
    criterios.append(("Faixa de tensão máxima MPPT", ok(vmp_max_oper <= vmpp_max if vmpp_max > 0 else True)))
    criterios.append(("Faixa de tensão mínima MPPT", ok(vmp_min_oper >= vstart if vstart > 0 else True)))
    criterios.append(("Corrente de operação", ok_current(imp_corr, imax)))
    criterios.append(("Corrente de curto", ok(isc_corr <= icc if icc > 0 else True)))
    criterios.append(("Overload Utilizado", ok(ov <= ovmax)))
    criterios.append(("Número de módulos por entrada", ok(n_min <= tam <= n_max)))
    criterios.append(("Número de strings/entradas", ok(ok_strings)))
    return criterios