from __future__ import annotations

from typing import Optional, Any, Dict, List, Tuple
from pydantic import BaseModel, Field
from typing import Literal

class CalcularRequest(BaseModel):
    kwp_sis: float = Field(gt=0, description="Potência desejada do sistema (kWp)")
    qtd_inv: int = Field(ge=1, description="Quantidade de inversores")

    fabricante_mod: str
    modelo_mod: str

    fabricante_inv: str
    modelo_inv: str

    tmin: float
    tmax: float

    tam_string: Optional[int] = Field(default=None, description="Tamanho da string (módulos) opcional")
    entradas_usadas: Optional[int] = Field(default=None, description="Qtd de entradas a usar (opcional)")


class CadastroRequest(BaseModel):
    dados: Dict[str, Any]


class CalcularResponse(BaseModel):
    ok: bool
    motivo: Optional[str] = None

    melhor: Optional[Dict[str, float]] = None
    combos: Optional[List[Dict[str, float]]] = None
    criterios_melhor: Optional[List[Tuple[str, str]]] = None

    intervalo_string: Optional[Dict[str, int]] = None
    ent_total: Optional[int] = None
    ent_usadas: Optional[int] = None
    ovmax: Optional[float] = None
    kwp_alvo_por_inv: Optional[float] = None
    pot_sis_melhor: Optional[float] = None

    inv: Optional[Dict[str, Any]] = None
    mod: Optional[Dict[str, Any]] = None
    correcoes: Optional[Dict[str, float]] = None