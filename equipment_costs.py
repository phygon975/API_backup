"""
Equipment cost module (Turton-based)

This module provides a flexible API to estimate equipment costs for
pumps, compressors, turbines, and fans using Turton's correlations.

Design goals:
- SI inputs (power in Watt) and SI-aware outputs
- Adjustable CEPCI and base year/index as inputs (defaults provided)
- Returns purchased, bare-module, and installed costs for TIC workflows
- Supports material and pressure factors (F_M, F_P) hooks
- Device subtypes with sensible defaults:
  - Pump: centrifugal (default)
  - Compressor: centrifugal (default)
  - Turbine: axial (default)
  - Fan: four types; default is centrifugal radial; limited to ~0.16 barg

NOTE:
- Turton correlation coefficients and BM/installation factors must be
  filled in the indicated placeholders during the next step.
"""

from dataclasses import dataclass
from typing import Literal, Optional, Dict, Tuple, List

# 분리된 모듈들 import
from aspen_data_extractor import get_aspen_cache, extract_mcompr_stage_data, get_unit_type_value
from unit_converter import (
    convert_power_to_target_unit, convert_flow_to_target_unit,
    convert_power_to_kw, convert_pressure_to_bar, convert_flow_to_m3_s,
    check_minimum_size_limit, get_max_size_limit, get_cepi_index
)
from config import (
    DEFAULT_MATERIAL, DEFAULT_INSTALL_FACTOR, DEFAULT_TARGET_YEAR,
    DEFAULT_CEPCI_BASE_INDEX, MCOMPR_BM_FACTOR_MULTIPLIER,
    DEFAULT_HEAT_EXCHANGER_BM_FACTOR, FAN_MAX_PRESSURE_RISE,
    TURBINE_MIN_PRESSURE_DROP, ENABLE_DEBUG_OUTPUT
)


# --------------------------------------------------------------------------------------
# Types
# --------------------------------------------------------------------------------------

PumpType = Literal["centrifugal", "reciprocating"]
CompressorType = Literal["centrifugal", "axial", "reciprocating"]
TurbineType = Literal["axial", "radial"]
FanType = Literal[
    "centrifugal_radial",  # default
    "centrifugal_backward_curved",
    "axial_tubeaxial",
    "axial_vaneless"
]

# Heat Exchanger types (first pass; can be extended)
HeatExchangerType = Literal[
    "double_pipe",
    "multiple_pipe",
    "fixed_tube",
    "floating_head",
    "bayonet",
    "kettle_reboiler",
    "scraped_wall",
    "teflon_tube",
    "air_cooler",
    "spiral_tube_shell",
    "spiral_plate",
    "flat_plate",
]

MaterialType = Literal[
    "CS",          # carbon steel
    "SS",          # stainless steel
    "Ni",          # nickel alloy
    "Cu",          # copper
    "Cl",          # cast iron (for pumps table)
    "Ti",          # titanium
    "Fiberglass",  # for fans table
    "Al"           # aluminum (air cooler tubes)
]


@dataclass
class CEPCIOptions:
    base_year: int = 2017
    base_index: float = 567.5  # CEPCI 2017 (Turton tables base)
    target_index: Optional[float] = None  # if None, equals base_index


@dataclass
class CostInputs:
    power_kilowatt: float = 0.0  # shaft power in kW (SI). Fans may use volumetric flow instead
    volumetric_flow_m3_s: Optional[float] = None  # for fans (m^3/s)
    material_factor: float = 1.0  # F_M
    pressure_factor: float = 1.0  # F_P (rating/MAWP correction)
    pressure_bar: Optional[float] = None  # optional pressure (abs or outlet, category-dependent)
    pressure_delta_bar: Optional[float] = None  # optional pressure rise/drop magnitude (bar)
    # Heat exchanger specific
    heat_duty_W: Optional[float] = None  # Q [W]
    overall_U_W_m2K: Optional[float] = None  # U [W/(m^2*K)]
    lmtd_K: Optional[float] = None  # LMTD [K]
    area_m2: Optional[float] = None  # if provided, overrides Q/(U*LMTD)
    # Optional meta
    notes: Optional[str] = None


# --------------------------------------------------------------------------------------
# Utilities
# --------------------------------------------------------------------------------------

# convert_flow_to_target_unit 함수는 unit_converter.py에서 import


# convert_power_to_target_unit 함수는 unit_converter.py에서 import


def adjust_cost_to_index(cost_at_base_index: float, base_index: float, target_index: Optional[float], debug_name: str = "") -> float:
    """
    Scale a cost reported at CEPCI base_index to target_index.
    If target_index is None, returns the original cost.
    """
    if target_index is None or target_index == 0:
        if debug_name:
            print(f"{debug_name}: CEPCI adjustment skipped (target_index is None or 0)")
        return cost_at_base_index
    if base_index is None or base_index == 0:
        if debug_name:
            print(f"{debug_name}: CEPCI adjustment skipped (base_index is None or 0)")
        return cost_at_base_index
    
    adjusted_cost = cost_at_base_index * (target_index / base_index)
    if debug_name:
        print(f"{debug_name}: CEPCI index adjustment")
        print(f"  Cost at base index: {cost_at_base_index:.2f} USD")
        print(f"  Base CEPCI index: {base_index:.1f}")
        print(f"  Target CEPCI index: {target_index:.1f}")
        print(f"  Adjustment factor: {target_index:.1f} / {base_index:.1f} = {target_index/base_index:.4f}")
        print(f"  Adjusted cost: {cost_at_base_index:.2f} * {target_index/base_index:.4f} = {adjusted_cost:.2f} USD")
    
    return adjusted_cost


# watt_to_kilowatt 함수 제거 - 이제 직접 kW로 변환하므로 불필요


# --------------------------------------------------------------------------------------
# Placeholder correlation helpers (to be completed in next steps)
# --------------------------------------------------------------------------------------

@dataclass
class LogQuadraticCoeff:
    """
    Represents Turton-style purchased cost correlation of the form:
      log10(Cp) = k1 + k2*log10(S) + k3*(log10(S))^2
    where S is the sizing variable (e.g., kW, hp, or m3/s for fans).
    """
    k1: float
    k2: float
    k3: float
    size_basis: Literal["kW", "hp", "m3/s"] = "kW"


# In-memory registries for coefficients per equipment subtype
_PUMP_COEFFS: Dict[PumpType, LogQuadraticCoeff] = {}
_COMP_COEFFS: Dict[CompressorType, LogQuadraticCoeff] = {}
_TURB_COEFFS: Dict[TurbineType, LogQuadraticCoeff] = {}
_FAN_COEFFS: Dict[FanType, LogQuadraticCoeff] = {}


def register_pump_correlation(pump_type: PumpType, coeff: LogQuadraticCoeff) -> None:
    _PUMP_COEFFS[pump_type] = coeff


def register_compressor_correlation(comp_type: CompressorType, coeff: LogQuadraticCoeff) -> None:
    _COMP_COEFFS[comp_type] = coeff


def register_turbine_correlation(turbine_type: TurbineType, coeff: LogQuadraticCoeff) -> None:
    _TURB_COEFFS[turbine_type] = coeff


def register_fan_correlation(fan_type: FanType, coeff: LogQuadraticCoeff) -> None:
    _FAN_COEFFS[fan_type] = coeff


def _eval_log_quadratic_cost(size_value: float, coeff: LogQuadraticCoeff, debug_name: str = "") -> float:
    import math
    if size_value <= 0:
        raise ValueError("Size value must be positive for cost correlation")
    
    logS = math.log10(size_value)
    logC = coeff.k1 + coeff.k2 * logS + coeff.k3 * (logS ** 2)
    cost = 10.0 ** logC
    
    # 디버그 출력
    if debug_name:
        print(f"{debug_name}: logS = math.log10({size_value:.4f}({coeff.size_basis}))")
        print(f"logC = {coeff.k1:.4f} + {coeff.k2:.4f} * logS + {coeff.k3:.4f} * logS**2")
        print(f"logC = {coeff.k1:.4f} + {coeff.k2:.4f} * {logS:.4f} + {coeff.k3:.4f} * {logS**2:.4f}")
        print(f"logC = {logC:.4f}")
        print(f"log_quadratic_cost = 10 ** {logC:.4f} = {cost:.2f}")
    
    return cost


def _turton_purchased_cost_pump_kw(power_kw: float, pump_type: PumpType) -> float:
    """
    Turton purchased cost correlation for pumps (base material, base pressure).
    RETURN: purchased cost at base CEPCI (in USD)

    Uses registered coefficients with flexible unit conversion.
    """
    coeff = _PUMP_COEFFS.get(pump_type)
    if coeff is None:
        raise NotImplementedError(f"Pump coefficients not registered for type: {pump_type}")
    
    # power_kw는 이미 kW 단위이므로 변환 없이 직접 사용
    size_val = power_kw
    print(f"PUMP {pump_type.upper()}: Input power = {power_kw:.4f} kW")
    print(f"Using power directly: {power_kw:.4f} kW (no conversion needed)")
    
    cost = _eval_log_quadratic_cost(size_val, coeff, f"PUMP {pump_type.upper()}")
    print(f"Base purchased cost = {cost:.2f} USD")
    return cost


def _turton_purchased_cost_compressor_kw(power_kw: float, comp_type: CompressorType) -> float:
    """Turton purchased cost correlation for compressors with flexible unit conversion."""
    coeff = _COMP_COEFFS.get(comp_type)
    if coeff is None:
        raise NotImplementedError(f"Compressor coefficients not registered for type: {comp_type}")
    
    # power_kw는 이미 kW 단위이므로 변환 없이 직접 사용
    size_val = power_kw
    print(f"COMPRESSOR {comp_type.upper()}: Input power = {power_kw:.4f} kW")
    print(f"Using power directly: {power_kw:.4f} kW (no conversion needed)")
    
    cost = _eval_log_quadratic_cost(size_val, coeff, f"COMPRESSOR {comp_type.upper()}")
    print(f"Base purchased cost = {cost:.2f} USD")
    return cost


def _turton_purchased_cost_turbine_kw(power_kw: float, turbine_type: TurbineType) -> float:
    """Turton purchased cost correlation for turbines with flexible unit conversion."""
    coeff = _TURB_COEFFS.get(turbine_type)
    if coeff is None:
        raise NotImplementedError(f"Turbine coefficients not registered for type: {turbine_type}")
    
    # power_kw는 이미 kW 단위이므로 변환 없이 직접 사용
    size_val = power_kw
    print(f"TURBINE {turbine_type.upper()}: Input power = {power_kw:.4f} kW")
    print(f"Using power directly: {power_kw:.4f} kW (no conversion needed)")
    
    cost = _eval_log_quadratic_cost(size_val, coeff, f"TURBINE {turbine_type.upper()}")
    print(f"Base purchased cost = {cost:.2f} USD")
    return cost




def _turton_purchased_cost_fan_flow(q_m3_s: float, fan_type: FanType) -> float:
    """
    Turton purchased cost correlation for fans sized by volumetric flow (m^3/s).
    Uses registered coefficients with flexible unit conversion.
    """
    coeff = _FAN_COEFFS.get(fan_type)
    if coeff is None:
        raise NotImplementedError(f"Fan coefficients not registered for type: {fan_type}")
    
    # q_m3_s는 이미 m³/s 단위이므로 변환 없이 직접 사용
    size_val = q_m3_s
    print(f"FAN {fan_type.upper()}: Input flow = {q_m3_s:.6f} m³/s")
    print(f"Using flow directly: {q_m3_s:.6f} m³/s (no conversion needed)")
    
    cost = _eval_log_quadratic_cost(size_val, coeff, f"FAN {fan_type.upper()}")
    print(f"Base purchased cost = {cost:.2f} USD")
    return cost


def _apply_material_pressure_factors(purchased_cost: float, F_M: float, F_P: float, debug_name: str = "") -> float:
    result = purchased_cost * F_M * F_P
    if debug_name:
        print(f"{debug_name}: Material & Pressure factors applied")
        print(f"  Base cost: {purchased_cost:.2f} USD")
        print(f"  Material factor (F_M): {F_M:.3f}")
        print(f"  Pressure factor (F_P): {F_P:.3f}")
        print(f"  Adjusted cost: {purchased_cost:.2f} * {F_M:.3f} * {F_P:.3f} = {result:.2f} USD")
    return result


def _to_bare_module_cost(purchased_cost: float, bm_factor: float, debug_name: str = "") -> float:
    result = purchased_cost * bm_factor
    if debug_name:
        print(f"{debug_name}: Bare Module cost calculation")
        print(f"  Purchased cost: {purchased_cost:.2f} USD")
        print(f"  BM factor: {bm_factor:.3f}")
        print(f"  Bare module cost: {purchased_cost:.2f} * {bm_factor:.3f} = {result:.2f} USD")
    return result


# _to_installed_cost 함수는 더 이상 사용되지 않음 (Bare Module Cost에 설치비 포함)
# def _to_installed_cost(bare_module_cost: float, install_factor: float, debug_name: str = "") -> float:
#     result = bare_module_cost * install_factor
#     if debug_name:
#         print(f"{debug_name}: Installed cost calculation")
#         print(f"  Bare module cost: {bare_module_cost:.2f} USD")
#         print(f"  Install factor: {install_factor:.3f}")
#         print(f"  Installed cost: {bare_module_cost:.2f} * {install_factor:.3f} = {result:.2f} USD")
#     return result


# --------------------------------------------------------------------------------------
# Public API
# --------------------------------------------------------------------------------------

# Wmax(kW) limits from provided Turton tables
_WMAX_KW: Dict[str, Dict[str, float]] = {
    "pump": {
        "centrifugal": 300.0,
        "reciprocating": 200.0,
    },
    "compressor": {
        "centrifugal": 3000.0,
        "axial": 3000.0,
        "reciprocating": 3000.0,
    },
    "turbine": {
        "axial": 4000.0,
        "radial": 1500.0,
    },
}

def _sum_costs(parts):
    total = {"purchased": 0.0, "purchased_adj": 0.0, "bare_module": 0.0, "installed": 0.0}
    for c in parts:
        for k in total:
            total[k] += float(c.get(k, 0.0))
    return total

def estimate_pump_cost(
    inputs: CostInputs,
    cepci: CEPCIOptions = CEPCIOptions(),
    pump_type: PumpType = "centrifugal",
    bm_factor: Optional[float] = None,  # if None, computed from Turton B1,B2 and material Fm
    material: MaterialType = "CS",
    install_factor: float = 1.0,  # placeholder
    _split: bool = True
) -> Dict[str, float]:
    """
    Returns a dict with keys: purchased, purchased_adj, bare_module, installed
    - purchased: at base CEPCI
    - purchased_adj: adjusted to target CEPCI (if provided)
    """
    # Split by Wmax if needed
    power_kw = inputs.power_kilowatt
    
    # 최소 크기 제한 체크
    is_valid, error_msg = check_minimum_size_limit("pump", pump_type, power_kw, "kW")
    if not is_valid:
        raise ValueError(f"Pump {pump_type} size {power_kw:.3f} kW is {error_msg}")
    
    if _split:
        max_kw = _WMAX_KW.get("pump", {}).get(pump_type)
        if max_kw and power_kw > max_kw:
            import math
            n = int(math.ceil(power_kw / max_kw))
            per_kw = power_kw / n
            per_inputs = CostInputs(
                power_kilowatt=per_kw,
                material_factor=inputs.material_factor,
                pressure_factor=inputs.pressure_factor,
                pressure_bar=inputs.pressure_bar,
                notes=inputs.notes,
            )
            parts = [
                estimate_pump_cost(per_inputs, cepci=cepci, pump_type=pump_type, bm_factor=bm_factor, material=material, install_factor=install_factor, _split=False)
                for _ in range(n)
            ]
            return _sum_costs(parts)
    purchased_base = _turton_purchased_cost_pump_kw(power_kw, pump_type)
    # material factor: pumps table provides Fm*; multiply purchased
    fm = _resolve_material_factor("pump", pump_type, material)
    fp = _resolve_pressure_factor("pump", pump_type, inputs.pressure_bar)
    purchased_base = _apply_material_pressure_factors(purchased_base, fm, fp, f"PUMP {pump_type.upper()}")
    purchased_adj = adjust_cost_to_index(purchased_base, cepci.base_index, cepci.target_index or cepci.base_index, f"PUMP {pump_type.upper()}")
    # Pump BM from Turton: FBM = B1 + B2 * Fm(material)
    if bm_factor is None:
        B1, B2 = _resolve_pump_b1b2(pump_type)
        effective_bm = B1 + B2 * fm
        print(f"PUMP {pump_type.upper()}: BM calculation (B1 + B2 * Fm)")
        print(f"  B1 = {B1:.3f}, B2 = {B2:.3f}, Fm = {fm:.3f}")
        print(f"  Effective BM = {B1:.3f} + {B2:.3f} * {fm:.3f} = {effective_bm:.3f}")
    else:
        effective_bm = bm_factor
        print(f"PUMP {pump_type.upper()}: Using provided BM factor = {effective_bm:.3f}")
    bare = _to_bare_module_cost(purchased_adj, effective_bm, f"PUMP {pump_type.upper()}")
    return {
        "purchased": purchased_base,
        "purchased_adj": purchased_adj,
        "bare_module": bare,
    }


def estimate_compressor_cost(
    inputs: CostInputs,
    cepci: CEPCIOptions = CEPCIOptions(),
    comp_type: CompressorType = "centrifugal",
    bm_factor: Optional[float] = None,
    material: MaterialType = "CS",
    install_factor: float = 1.0,
    _split: bool = True
) -> Dict[str, float]:
    power_kw = inputs.power_kilowatt
    
    # 최소 크기 제한 체크
    is_valid, error_msg = check_minimum_size_limit("compressor", comp_type, power_kw, "kW")
    if not is_valid:
        raise ValueError(f"Compressor {comp_type} size {power_kw:.3f} kW is {error_msg}")
    
    if _split:
        max_kw = _WMAX_KW.get("compressor", {}).get(comp_type)
        if max_kw and power_kw > max_kw:
            import math
            n = int(math.ceil(power_kw / max_kw))
            per_kw = power_kw / n
            per_inputs = CostInputs(
                power_kilowatt=per_kw,
                material_factor=inputs.material_factor,
                pressure_factor=inputs.pressure_factor,
                pressure_bar=inputs.pressure_bar,
                notes=inputs.notes,
            )
            parts = [
                estimate_compressor_cost(per_inputs, cepci=cepci, comp_type=comp_type, bm_factor=bm_factor, material=material, install_factor=install_factor, _split=False)
                for _ in range(n)
            ]
            return _sum_costs(parts)
    purchased_base = _turton_purchased_cost_compressor_kw(power_kw, comp_type)
    # compressor table provides FBM by material; material factor defaults to 1
    purchased_base = _apply_material_pressure_factors(purchased_base, 1.0, 1.0, f"COMPRESSOR {comp_type.upper()}")
    purchased_adj = adjust_cost_to_index(purchased_base, cepci.base_index, cepci.target_index or cepci.base_index, f"COMPRESSOR {comp_type.upper()}")
    effective_bm = bm_factor if bm_factor is not None else _resolve_bm("compressor", comp_type, material)
    if bm_factor is None:
        print(f"COMPRESSOR {comp_type.upper()}: Using BM factor from table = {effective_bm:.3f}")
    else:
        print(f"COMPRESSOR {comp_type.upper()}: Using provided BM factor = {effective_bm:.3f}")
    bare = _to_bare_module_cost(purchased_adj, effective_bm, f"COMPRESSOR {comp_type.upper()}")
    return {
        "purchased": purchased_base,
        "purchased_adj": purchased_adj,
        "bare_module": bare,
    }


def estimate_turbine_cost(
    inputs: CostInputs,
    cepci: CEPCIOptions = CEPCIOptions(),
    turbine_type: TurbineType = "axial",
    bm_factor: Optional[float] = None,
    material: MaterialType = "CS",
    install_factor: float = 1.0,
    _split: bool = True
) -> Dict[str, float]:
    # Turbine WNET is typically reported as negative (power produced).
    # Use absolute power for cost correlations.
    power_kw = abs(inputs.power_kilowatt)
    
    # 최소 크기 제한 체크
    is_valid, error_msg = check_minimum_size_limit("turbine", turbine_type, power_kw, "kW")
    if not is_valid:
        raise ValueError(f"Turbine {turbine_type} size {power_kw:.3f} kW is {error_msg}")
    
    if _split:
        max_kw = _WMAX_KW.get("turbine", {}).get(turbine_type)
        if max_kw and power_kw > max_kw:
            import math
            n = int(math.ceil(power_kw / max_kw))
            per_kw = power_kw / n
            per_inputs = CostInputs(
                power_kilowatt=per_kw,
                material_factor=inputs.material_factor,
                pressure_factor=inputs.pressure_factor,
                pressure_bar=inputs.pressure_bar,
                notes=inputs.notes,
            )
            parts = [
                estimate_turbine_cost(per_inputs, cepci=cepci, turbine_type=turbine_type, bm_factor=bm_factor, material=material, install_factor=install_factor, _split=False)
                for _ in range(n)
            ]
            return _sum_costs(parts)
    purchased_base = _turton_purchased_cost_turbine_kw(power_kw, turbine_type)
    purchased_base = _apply_material_pressure_factors(purchased_base, 1.0, 1.0, f"TURBINE {turbine_type.upper()}")
    purchased_adj = adjust_cost_to_index(purchased_base, cepci.base_index, cepci.target_index or cepci.base_index, f"TURBINE {turbine_type.upper()}")
    effective_bm = bm_factor if bm_factor is not None else _resolve_bm("turbine", turbine_type, material)
    if bm_factor is None:
        print(f"TURBINE {turbine_type.upper()}: Using BM factor from table = {effective_bm:.3f}")
    else:
        print(f"TURBINE {turbine_type.upper()}: Using provided BM factor = {effective_bm:.3f}")
    bare = _to_bare_module_cost(purchased_adj, effective_bm, f"TURBINE {turbine_type.upper()}")
    return {
        "purchased": purchased_base,
        "purchased_adj": purchased_adj,
        "bare_module": bare,
    }


def estimate_mcompr_cost(
    inputs: CostInputs,
    cepci: CEPCIOptions = CEPCIOptions(),
    bm_factor: Optional[float] = None,
    material: MaterialType = "CS",
    install_factor: float = 1.0
) -> Dict[str, float]:
    """
    MCompr (Multi-stage Compressor) 비용 추정
    MCompr는 일반적으로 더 복잡한 구조를 가지므로 별도의 비용 계산이 필요할 수 있음
    현재는 일반 압축기와 동일한 방식으로 처리하되, 향후 MCompr 전용 상관관계 추가 가능
    """
    power_kw = inputs.power_kilowatt
    
    # MCompr는 일반적으로 더 큰 용량이므로 분할 처리
    max_kw = 5000.0  # MCompr 전용 최대 용량 설정
    if power_kw > max_kw:
        import math
        n = int(math.ceil(power_kw / max_kw))
        per_kw = power_kw / n
        per_inputs = CostInputs(
            power_kilowatt=per_kw,
            material_factor=inputs.material_factor,
            pressure_factor=inputs.pressure_factor,
            pressure_bar=inputs.pressure_bar,
            notes=inputs.notes,
        )
        parts = [
            estimate_mcompr_cost(per_inputs, cepci=cepci, bm_factor=bm_factor, material=material, install_factor=install_factor)
            for _ in range(n)
        ]
        return _sum_costs(parts)
    
    # MCompr는 일반적으로 원심 압축기로 가정
    purchased_base = _turton_purchased_cost_compressor_kw(power_kw, "centrifugal")
    purchased_base = _apply_material_pressure_factors(purchased_base, 1.0, 1.0, "MCOMPR CENTRIFUGAL")
    purchased_adj = adjust_cost_to_index(purchased_base, cepci.base_index, cepci.target_index or cepci.base_index, "MCOMPR CENTRIFUGAL")
    
    # MCompr는 더 복잡한 구조이므로 BM 인수가 일반 압축기보다 높을 수 있음
    effective_bm = bm_factor if bm_factor is not None else _resolve_bm("compressor", "centrifugal", material) * 1.2  # 20% 추가
    if bm_factor is None:
        base_bm = _resolve_bm("compressor", "centrifugal", material)
        print(f"MCOMPR CENTRIFUGAL: Using BM factor from table = {base_bm:.3f} * 1.2 = {effective_bm:.3f}")
    else:
        print(f"MCOMPR CENTRIFUGAL: Using provided BM factor = {effective_bm:.3f}")
    bare = _to_bare_module_cost(purchased_adj, effective_bm, "MCOMPR CENTRIFUGAL")
    
    return {
        "purchased": purchased_base,
        "purchased_adj": purchased_adj,
        "bare_module": bare,
    }


def estimate_fan_cost(
    inputs: CostInputs,
    cepci: CEPCIOptions = CEPCIOptions(),
    fan_type: FanType = "centrifugal_radial",
    bm_factor: Optional[float] = None,
    material: MaterialType = "CS",
    install_factor: float = 1.0
) -> Dict[str, float]:
    # Fans are sized by volumetric gas flow. Expect m^3/s in inputs.
    if inputs.volumetric_flow_m3_s is None:
        raise ValueError("Fan cost calculation requires volumetric flow (m^3/s). Power-based calculation is not supported.")
    q_m3_s = max(0.0, float(inputs.volumetric_flow_m3_s))
    
    # 최소 크기 제한 체크
    is_valid, error_msg = check_minimum_size_limit("fan", fan_type, q_m3_s, "m³/s")
    if not is_valid:
        raise ValueError(f"Fan {fan_type} size {q_m3_s:.6f} m³/s is {error_msg}")
    purchased_base = _turton_purchased_cost_fan_flow(q_m3_s, fan_type)
    if ENABLE_DEBUG_OUTPUT:
        if inputs.pressure_delta_bar is None:
            print("FAN WARNING: pressure_delta_bar is None -> FP defaults to 1.0")
        else:
            print(f"FAN ΔP provided: {inputs.pressure_delta_bar*100.0:.3f} kPa")
    fp = _resolve_pressure_factor("fan", fan_type, inputs.pressure_bar, inputs.pressure_delta_bar)
    purchased_base = _apply_material_pressure_factors(purchased_base, 1.0, fp, f"FAN {fan_type.upper()}")
    purchased_adj = adjust_cost_to_index(purchased_base, cepci.base_index, cepci.target_index or cepci.base_index, f"FAN {fan_type.upper()}")
    effective_bm = bm_factor if bm_factor is not None else _resolve_bm("fan", fan_type, material)
    if bm_factor is None:
        print(f"FAN {fan_type.upper()}: Using BM factor from table = {effective_bm:.3f}")
    else:
        print(f"FAN {fan_type.upper()}: Using provided BM factor = {effective_bm:.3f}")
    bare = _to_bare_module_cost(purchased_adj, effective_bm, f"FAN {fan_type.upper()}")
    return {
        "purchased": purchased_base,
        "purchased_adj": purchased_adj,
        "bare_module": bare,
    }


# Convenience function for TIC breakdown (sum or per-equipment handled externally)
def format_cost_breakdown(costs: Dict[str, float]) -> str:
    return (
        f"Purchased (base): {costs['purchased']:.2f}\n"
        f"Purchased (adj):  {costs['purchased_adj']:.2f}\n"
        f"Bare module:      {costs['bare_module']:.2f}\n"
        f"Installed:        {costs['installed']:.2f}"
    )


# --------------------------------------------------------------------------------------
# Default correlation and BM-factor registration from provided tables
# --------------------------------------------------------------------------------------

_BM_MAP: Dict[str, Dict[str, Dict[str, float]]] = {
    # category -> subtype -> material -> BM
    "compressor": {
        "centrifugal": {"CS": 2.7, "SS": 5.8, "Ni": 11.5},
        "axial": {"CS": 3.8, "SS": 8.0, "Ni": 15.9},
        # rotary and reciprocating present in data; subtype names differ from our CompressorType
        # they can be added later if needed
        "reciprocating": {"CS": 3.4, "SS": 7.0, "Ni": 13.9},
    },
    "fan": {
        "centrifugal_radial": {"CS": 2.7, "Fiberglass": 5.0, "SS": 5.8, "Ni": 11.5},
        "centrifugal_backward_curved": {"CS": 2.7, "Fiberglass": 5.0, "SS": 5.8, "Ni": 11.5},
        "axial_tubeaxial": {"CS": 2.7, "Fiberglass": 5.0, "SS": 5.8, "Ni": 11.5},
        "axial_vaneless": {"CS": 2.7, "Fiberglass": 5.0, "SS": 5.8, "Ni": 11.5},
    },
    "turbine": {
        "axial": {"CS": 3.5, "SS": 6.1, "Ni": 11.7},
        "radial": {"CS": 3.5, "SS": 6.1, "Ni": 11.7},
    },
    # Pumps table shows material correction factors (Fm*): use as material factors, not BM
}


def _resolve_bm(category: str, subtype: str, material: MaterialType = "CS") -> float:
    by_cat = _BM_MAP.get(category, {})
    by_type = by_cat.get(subtype, {})
    # default to CS if not found
    return by_type.get(material, by_type.get("CS", 1.0))


_FM_MAP: Dict[str, Dict[str, Dict[str, float]]] = {
    # category -> subtype -> material -> F_M
    "pump": {
        "centrifugal": {"Cl": 1.0, "CS": 1.6, "SS": 2.3, "Ni": 4.4},
        "reciprocating": {"Cl": 1.0, "CS": 1.5, "SS": 2.4, "Ni": 4.0, "Ti": 6.4},
    }
}


def _resolve_material_factor(category: str, subtype: str, material: MaterialType = "CS") -> float:
    by_cat = _FM_MAP.get(category, {})
    by_type = by_cat.get(subtype, {})
    return by_type.get(material, 1.0)


_PUMP_B1B2: Dict[PumpType, Tuple[float, float]] = {
    # From provided Pump Data table (B1, B2 columns)
    "centrifugal": (1.89, 1.35),
    "reciprocating": (1.89, 1.35),  # table shows same B1,B2 for listed pump types
}


def _resolve_pump_b1b2(pump_type: PumpType) -> Tuple[float, float]:
    return _PUMP_B1B2.get(pump_type, (1.89, 1.35))


def register_default_correlations() -> None:
    """Register k1,k2,k3 coefficients parsed from the provided tables (kW basis)."""
    # Compressors (without electric motors) — power basis kW
    register_compressor_correlation("centrifugal", LogQuadraticCoeff(k1=2.2891, k2=1.3604, k3=-0.1027, size_basis="kW"))
    register_compressor_correlation("axial", LogQuadraticCoeff(k1=2.2891, k2=1.3604, k3=-0.1027, size_basis="kW"))
    register_compressor_correlation("reciprocating", LogQuadraticCoeff(k1=2.2891, k2=1.3604, k3=-0.1027, size_basis="kW"))
    # (Rotary available but not in enum; can be added if needed)

    # Fans (include electric motors) — treat size basis as kW per requirement
    register_fan_correlation("centrifugal_radial", LogQuadraticCoeff(k1=3.5391, k2=-0.3533, k3=0.4477, size_basis="m3/s"))
    register_fan_correlation("centrifugal_backward_curved", LogQuadraticCoeff(k1=3.3471, k2=-0.0734, k3=0.3090, size_basis="m3/s"))
    register_fan_correlation("axial_tubeaxial", LogQuadraticCoeff(k1=3.0414, k2=-0.3375, k3=0.4722, size_basis="m3/s"))
    register_fan_correlation("axial_vaneless", LogQuadraticCoeff(k1=3.1761, k2=-0.1373, k3=0.3414, size_basis="m3/s"))

    # Pumps (including electric drives)
    register_pump_correlation("centrifugal", LogQuadraticCoeff(k1=3.3892, k2=0.0536, k3=0.1538, size_basis="kW"))
    register_pump_correlation("reciprocating", LogQuadraticCoeff(k1=3.8696, k2=0.3161, k3=0.1220, size_basis="kW"))
    # Positive displacement available in table; add later if needed

    # Turbines
    register_turbine_correlation("axial", LogQuadraticCoeff(k1=2.7051, k2=1.4398, k3=-0.1776, size_basis="kW"))
    register_turbine_correlation("radial", LogQuadraticCoeff(k1=2.2476, k2=1.4965, k3=-0.1618, size_basis="kW"))


# Pressure factor resolution based on table constraints
_PRESSURE_LIMIT_BAR: Dict[str, float] = {
    # Kept for legacy checks (selection), not used in C1/C2/C3 model
    "fan": 0.16,
}


# ----------------------------------------------------------------------------
# Pressure factor by device using Turton form:
#   log10(Fp) = C1 + C2*log10(P) + C3*(log10(P))^2
# P units: unless specified otherwise in table, use barg (gauge bar).
# For fans, table specifies ΔP in kPa.
# ----------------------------------------------------------------------------

# Coefficient tables: category -> subtype -> list of (range_checker, unit, C1,C2,C3)
# For ranges, provide lambdas taking a numeric P (already in target unit).
_PF_TABLE: Dict[str, Dict[str, List[Tuple]] ] = {
    # heat exchangers: pressure factor will be added later using C1,C2,C3 from provided table
    "heat_exchanger": {},
    "pump": {
        # P in barg (gauge bar)
        "centrifugal": [
            (lambda p_barg: p_barg < 10.0, "barg", 0.0, 0.0, 0.0),
            (lambda p_barg: 10.0 < p_barg < 100.0, "barg", -0.3935, 0.3957, -0.00226),
        ],
        "reciprocating": [
            (lambda p_barg: p_barg < 10.0, "barg", 0.0, 0.0, 0.0),
            (lambda p_barg: 10.0 < p_barg < 100.0, "barg", -0.245382, 0.259016, -0.01363),
        ],
        # positive displacement not explicitly in enum; if added later, reuse reciprocating coeffs
    },
    "fan": {
        # Centrifugal radial and backward curved
        "centrifugal_radial": [
            (lambda dp_kpa: dp_kpa < 1.0, "kPa", 0.0, 0.0, 0.0),
            (lambda dp_kpa: 1.0 <= dp_kpa < 16.0, "kPa", 0.0, 0.20899, -0.0328),
        ],
        "centrifugal_backward_curved": [
            (lambda dp_kpa: dp_kpa < 1.0, "kPa", 0.0, 0.0, 0.0),
            (lambda dp_kpa: 1.0 <= dp_kpa < 16.0, "kPa", 0.0, 0.20899, -0.0328),
        ],
        # Axial vane and axial tube
        "axial_tubeaxial": [
            (lambda dp_kpa: dp_kpa < 1.0, "kPa", 0.0, 0.0, 0.0),
            (lambda dp_kpa: 1.0 <= dp_kpa < 4.0, "kPa", 0.0, 0.20899, -0.0328),
        ],
        "axial_vaneless": [
            (lambda dp_kpa: dp_kpa < 1.0, "kPa", 0.0, 0.0, 0.0),
            (lambda dp_kpa: 1.0 <= dp_kpa < 4.0, "kPa", 0.0, 0.20899, -0.0328),
        ],
    },
    # Compressors and turbines: pressure factors not applied per table (C1=C2=C3=0)
    "compressor": {
        "centrifugal": [],
        "axial": [],
        "reciprocating": [],
    },
    "turbine": {
        "axial": [],
        "radial": [],
    },
}


def _calc_fp_from_coeffs(P_value: float, C1: float, C2: float, C3: float) -> float:
    import math
    if P_value is None or P_value <= 0:
        return 1.0
    logP = math.log10(P_value)
    logFp = C1 + C2 * logP + C3 * (logP ** 2)
    Fp = 10.0 ** logFp
    if ENABLE_DEBUG_OUTPUT:
        print(f"FP calculation: log10(Fp) = {C1:.6f} + {C2:.6f}*log10(P) + {C3:.6f}*(log10(P))^2")
        print(f"  P = {P_value:.6g}")
        print(f"  log10(P) = {logP:.6f}")
        print(f"  log10(Fp) = {logFp:.6f}")
        print(f"  Fp = 10**{logFp:.6f} = {Fp:.6f}")
    return Fp


def _resolve_pressure_factor(category: str, subtype: str, pressure_bar: Optional[float], pressure_delta_bar: Optional[float] = None) -> float:
    table = _PF_TABLE.get(category, {}).get(subtype)
    if table is None:
        return 1.0
    # No entries → factor 1
    if len(table) == 0:
        return 1.0
    # Determine P in required unit
    if category == "fan":
        # For fans, the table uses ΔP in kPa
        if pressure_delta_bar is None:
            # fall back: if only absolute provided, assume low ΔP → FP=1
            return 1.0
        dp_kpa = float(pressure_delta_bar) * 100.0  # bar -> kPa
        if ENABLE_DEBUG_OUTPUT:
            print(f"RESOLVE FP [fan/{subtype}] using ΔP = {dp_kpa:.6g} kPa")
        for checker, unit, C1, C2, C3 in table:
            if unit == "kPa" and checker(dp_kpa):
                return _calc_fp_from_coeffs(dp_kpa, C1, C2, C3)
        # If out of table range, clamp to last valid segment
        # choose the last segment's coeffs
        checker, unit, C1, C2, C3 = table[-1]
        return _calc_fp_from_coeffs(dp_kpa, C1, C2, C3)
    
    if category == "pump":
        if pressure_bar is None:
            return 1.0
        # convert abs bar to barg (assume abs passed from upstream)
        p_g_barg = max(0.0, float(pressure_bar) - 1.01325)
        if ENABLE_DEBUG_OUTPUT:
            print(f"RESOLVE FP [pump/{subtype}] using P_g = {p_g_barg:.6g} barg")
        for checker, unit, C1, C2, C3 in table:
            if unit == "barg" and checker(p_g_barg):
                return _calc_fp_from_coeffs(p_g_barg, C1, C2, C3)
        # out of range → clamp to last segment
        checker, unit, C1, C2, C3 = table[-1]
        return _calc_fp_from_coeffs(p_g_barg, C1, C2, C3)

    # Default: use gauge bar if unit unspecified (most equipment)
    if pressure_bar is None:
        return 1.0
    P_g_barg = max(0.0, float(pressure_bar))  # callers should pass gauge; but keep non-negative guard
    # There is no detailed table yet for other equipment; return 1.0 by default
    return 1.0


# --------------------------------------------------------------------------------------
# Heat Exchangers: purchased/BM cost (size basis = area m^2)
# --------------------------------------------------------------------------------------

# K1,K2,K3 for purchased cost: log10(Cp) = K1 + K2*log10(A) + K3*(log10(A))^2
# Note: Fill these from the provided table per heat exchanger subtype.
_HX_COEFFS: Dict[str, LogQuadraticCoeff] = {
    # K1,K2,K3 from provided Heat Exchanger Data table (size basis: area m2)
    "double_pipe": LogQuadraticCoeff(k1=3.3444, k2=0.2745, k3=-0.0472, size_basis="m2"),
    "multiple_pipe": LogQuadraticCoeff(k1=2.7652, k2=0.7282, k3=0.0783, size_basis="m2"),
    "fixed_tube": LogQuadraticCoeff(k1=4.3247, k2=-0.3030, k3=0.1634, size_basis="m2"),
    "floating_head": LogQuadraticCoeff(k1=4.8306, k2=-0.8509, k3=0.3187, size_basis="m2"),
    "bayonet": LogQuadraticCoeff(k1=4.2768, k2=-0.0495, k3=0.1431, size_basis="m2"),
    "kettle_reboiler": LogQuadraticCoeff(k1=4.4646, k2=-0.5277, k3=0.3955, size_basis="m2"),
    "scraped_wall": LogQuadraticCoeff(k1=3.7803, k2=0.8569, k3=0.0349, size_basis="m2"),
    "teflon_tube": LogQuadraticCoeff(k1=3.8062, k2=0.8924, k3=-0.1671, size_basis="m2"),
    "air_cooler": LogQuadraticCoeff(k1=4.0336, k2=0.2341, k3=0.0497, size_basis="m2"),
    "spiral_tube_shell": LogQuadraticCoeff(k1=3.9912, k2=0.0668, k3=0.2430, size_basis="m2"),
    "spiral_plate": LogQuadraticCoeff(k1=4.6561, k2=-0.2947, k3=0.2207, size_basis="m2"),
    "flat_plate": LogQuadraticCoeff(k1=4.6656, k2=-0.1557, k3=0.1547, size_basis="m2"),
}

# BM factors: FBM = B1 + B2 * Fm (composite material factor)
_HX_B1B2: Dict[str, Tuple[float, float]] = {
    # B1,B2 from provided Heat Exchanger Data table
    "double_pipe": (1.74, 1.55),
    "multiple_pipe": (1.74, 1.55),
    "fixed_tube": (1.63, 1.66),
    "floating_head": (1.63, 1.66),
    "bayonet": (1.63, 1.66),
    "kettle_reboiler": (1.63, 1.66),
    "scraped_wall": (1.74, 1.55),
    "teflon_tube": (1.63, 1.66),
    "air_cooler": (0.96, 1.21),
    "spiral_tube_shell": (1.74, 1.55),
    "spiral_plate": (0.96, 1.21),
    "flat_plate": (0.96, 1.21),
}

# Material factors by side (shell/tube). To be populated from material factor matrix.
_HX_FM_SHELL: Dict[str, Dict[str, float]] = {}

_HX_FM_TUBE: Dict[str, Dict[str, float]] = {}

# Optional full combination matrix: hx_type -> shell -> tube -> Fm (if provided, used directly)
_HX_FM_COMBO: Dict[str, Dict[str, Dict[str, float]]] = {
    # 조합 표 (쉘→튜브): 제공된 조합만 유효
    "double_pipe": {"CS": {"CS": 1.00, "Cu": 1.35, "SS": 1.81, "Ni": 2.68, "Ti": 4.63}, "Cu": {"Cu": 1.69}, "SS": {"SS": 2.73}, "Ni": {"Ni": 3.73}, "Ti": {"Ti": 11.38}},
    "multiple_pipe": {"CS": {"CS": 1.00, "Cu": 1.35, "SS": 1.81, "Ni": 2.68, "Ti": 4.63}, "Cu": {"Cu": 1.69}, "SS": {"SS": 2.73}, "Ni": {"Ni": 3.73}, "Ti": {"Ti": 11.38}},
    "fixed_tube": {"CS": {"CS": 1.00, "Cu": 1.35, "SS": 1.81, "Ni": 2.68, "Ti": 4.63}, "Cu": {"Cu": 1.69}, "SS": {"SS": 2.73}, "Ni": {"Ni": 3.73}, "Ti": {"Ti": 11.38}},
    "floating_head": {"CS": {"CS": 1.00, "Cu": 1.35, "SS": 1.81, "Ni": 2.68, "Ti": 4.63}, "Cu": {"Cu": 1.69}, "SS": {"SS": 2.73}, "Ni": {"Ni": 3.73}, "Ti": {"Ti": 11.38}},
    "bayonet": {"CS": {"CS": 1.00, "Cu": 1.35, "SS": 1.81, "Ni": 2.68, "Ti": 4.63}, "Cu": {"Cu": 1.69}, "SS": {"SS": 2.73}, "Ni": {"Ni": 3.73}, "Ti": {"Ti": 11.38}},
    "kettle_reboiler": {"CS": {"CS": 1.00, "Cu": 1.35, "SS": 1.81, "Ni": 2.68, "Ti": 4.63}, "Cu": {"Cu": 1.69}, "SS": {"SS": 2.73}, "Ni": {"Ni": 3.73}, "Ti": {"Ti": 11.38}},
    "scraped_wall": {"CS": {"CS": 1.00, "Cu": 1.35, "SS": 1.81, "Ni": 2.68, "Ti": 4.63}, "Cu": {"Cu": 1.69}, "SS": {"SS": 2.73}, "Ni": {"Ni": 3.73}, "Ti": {"Ti": 11.38}},
    "spiral_tube_shell": {"CS": {"CS": 1.00, "Cu": 1.35, "SS": 1.81, "Ni": 2.68, "Ti": 4.63}, "Cu": {"Cu": 1.69}, "SS": {"SS": 2.73}, "Ni": {"Ni": 3.73}, "Ti": {"Ti": 11.38}},
    # Plate-type exchangers: 표가 "Material In Contact with Process Fluid" 기준이므로
    # 쉘 키를 'CS'로 고정하고, 튜브 쪽에 재질을 매핑하여 선택하도록 처리
    "spiral_plate": {"CS": {"CS": 1.00, "Cu": 1.35, "SS": 2.45, "Ni": 2.68, "Ti": 4.63}},
    "flat_plate": {"CS": {"CS": 1.00, "Cu": 1.35, "SS": 2.45, "Ni": 2.68, "Ti": 4.63}},
}

# Special-case tables
_HX_FM_TEF_SHELL: Dict[str, float] = {  # Teflon Tube Exchanger: shell-side only
    "CS": 1.00, "Cu": 1.20, "SS": 1.30, "Ni": 2.68, "Ti": 3.30
}

_HX_FM_AIR_TUBE: Dict[str, float] = {  # Air cooler tube material
    "CS": 1.00, "Al": 1.42, "SS": 2.93
}

def _resolve_hx_material_factor(hx_type: str, shell_material: MaterialType, tube_material: Optional[MaterialType]) -> float:
    # Teflon tube: tube fixed (Teflon), shell selects per dedicated table
    if hx_type == "teflon_tube":
        return _HX_FM_TEF_SHELL.get(shell_material, 1.0)
    # Air cooler: only tube side counts
    if hx_type == "air_cooler":
        return _HX_FM_AIR_TUBE.get(tube_material or "CS", 1.0)
    # If full combo matrix exists, use it
    by_shell = _HX_FM_COMBO.get(hx_type)
    if by_shell and shell_material in by_shell:
        by_tube = by_shell[shell_material]
        if tube_material is None:
            raise ValueError(f"열교환기 '{hx_type}': 튜브 재질을 선택하세요. 가능한 튜브 재질: {', '.join(sorted(by_tube.keys()))}")
        fm_combo = by_tube.get(tube_material)
        if fm_combo is not None:
            return fm_combo
        raise ValueError(
            f"열교환기 '{hx_type}': 지원되지 않는 재질 조합입니다 (shell={shell_material}, tube={tube_material}). "
            f"shell={shell_material}에서 가능한 튜브 재질: {', '.join(sorted(by_tube.keys()))}"
        )
    # No matrix defined for this hx_type
    raise ValueError(
        f"열교환기 '{hx_type}': 재질 조합 매트릭스가 정의되지 않았습니다. 지원되는 타입을 선택하거나 매트릭스를 제공하세요."
    )


def get_hx_material_options(hx_type: HeatExchangerType) -> Dict[str, List[str]]:
    """열교환기 타입별 재질 선택 가이드를 반환합니다.
    반환 형식: {"shell": [...], "tube": [...], "notes": [str,...]}
    """
    guide: Dict[str, List[str] or str] = {"shell": [], "tube": [], "notes": []}
    if hx_type == "teflon_tube":
        guide["shell"] = sorted(_HX_FM_TEF_SHELL.keys())
        guide["tube"] = ["Teflon (fixed)"]
        guide["notes"] = ["Teflon Tube: 튜브는 테플론 고정, 쉘 재질만 선택"]
        return guide
    if hx_type == "air_cooler":
        guide["shell"] = ["(선택 불가)"]
        guide["tube"] = sorted(_HX_FM_AIR_TUBE.keys())
        guide["notes"] = ["Air Cooler: 공기측 장치이므로 튜브 재질만 선택(CS/Al/SS)"]
        return guide
    if hx_type in ("spiral_plate", "flat_plate"):
        # plate류: 프로세스 유체 재질만 선택 → 튜브 목록으로 제공
        plate = _HX_FM_COMBO.get(hx_type, {}).get("CS", {})
        guide["shell"] = ["(내부 고정)"]
        guide["tube"] = sorted(plate.keys())
        guide["notes"] = ["Plate류: 표의 'Material In Contact with Process Fluid' 재질을 선택"]
        return guide
    # 일반 쉘&튜브형: 조합표 기준
    by_shell = _HX_FM_COMBO.get(hx_type)
    if by_shell:
        guide["shell"] = sorted(by_shell.keys())
        # shell=CS 기준 가능한 튜브 목록(참고용)
        cs_tube = sorted(by_shell.get("CS", {}).keys())
        guide["tube"] = cs_tube
        guide["notes"] = [
            "쉘 재질을 먼저 선택하면 해당 쉘에서 가능한 튜브 재질 목록이 제시됩니다.",
            "일부 쉘 재질은 동일 재질 튜브만 허용(Cu→Cu, SS→SS, Ni→Ni, Ti→Ti)",
        ]
        return guide
    guide["notes"] = ["해당 타입의 재질 조합 표가 아직 등록되지 않았습니다."]
    return guide


def _hx_compute_area(inputs: CostInputs) -> float:
    if inputs.area_m2 is not None and inputs.area_m2 > 0:
        return float(inputs.area_m2)
    if inputs.heat_duty_W is not None and inputs.overall_U_W_m2K is not None and inputs.lmtd_K is not None:
        q = float(inputs.heat_duty_W)
        u = float(inputs.overall_U_W_m2K)
        dt = float(inputs.lmtd_K)
        if u <= 0 or dt <= 0:
            raise ValueError("Invalid U or LMTD for area calculation")
        return q / (u * dt)
    raise ValueError("Heat exchanger area cannot be determined. Provide area_m2 or (heat_duty_W, overall_U_W_m2K, lmtd_K)")


def estimate_heat_exchanger_cost(
    inputs: CostInputs,
    cepci: CEPCIOptions = CEPCIOptions(),
    hx_type: HeatExchangerType = "fixed_tube",
    material_shell: MaterialType = "CS",
    material_tube: Optional[MaterialType] = None,
) -> Dict[str, float]:
    # 1) Area (m2)
    area_m2 = _hx_compute_area(inputs)
    if ENABLE_DEBUG_OUTPUT:
        print(f"HX INPUTS: area={area_m2:.4f} m2, Q={inputs.heat_duty_W}, U={inputs.overall_U_W_m2K}, LMTD={inputs.lmtd_K}")
    # 2) Purchased cost via HX K1,K2,K3
    coeff = _HX_COEFFS.get(hx_type)
    if coeff is None:
        raise NotImplementedError(f"Heat exchanger coefficients not registered for type: {hx_type}")
    purchased_base = _eval_log_quadratic_cost(area_m2, coeff, f"HX {hx_type.upper()}")
    # 3) Material factor: combine shell/tube (multiplicative assumption)
    fm = _resolve_hx_material_factor(hx_type, material_shell, material_tube)
    # 4) Pressure factor: pending separate table -> 1.0 for now
    fp = 1.0
    purchased_base = _apply_material_pressure_factors(purchased_base, fm, fp, f"HX {hx_type.upper()}")
    # 5) CEPCI adjustment
    purchased_adj = adjust_cost_to_index(purchased_base, cepci.base_index, cepci.target_index or cepci.base_index, f"HX {hx_type.upper()}")
    # 6) Bare module via B1,B2
    b1b2 = _HX_B1B2.get(hx_type)
    if b1b2 is None:
        raise NotImplementedError(f"Heat exchanger B1,B2 not registered for type: {hx_type}")
    B1, B2 = b1b2
    effective_bm = B1 + B2 * fm
    if ENABLE_DEBUG_OUTPUT:
        print(f"HX {hx_type.upper()}: BM (B1 + B2*Fm) = {B1:.3f} + {B2:.3f}*{fm:.3f} = {effective_bm:.3f}")
    bare = _to_bare_module_cost(purchased_adj, effective_bm, f"HX {hx_type.upper()}")
    return {
        "purchased": purchased_base,
        "purchased_adj": purchased_adj,
        "bare_module": bare,
        "area_m2": area_m2,
        "Fm": fm,
    }


# --------------------------------------------------------------------------------------
# Data Caching System
# --------------------------------------------------------------------------------------

class AspenDataCache:
    """Aspen Plus 데이터를 캐싱하여 중복 읽기를 방지하는 클래스"""
    
    def __init__(self):
        self._block_data = {}
        self._unit_data = {}
        self._pressure_data = {}
        self._power_data = {}
    
    def get_block_data(self, Application, block_name: str) -> Dict[str, any]:
        """블록의 모든 데이터를 한 번에 읽어서 캐싱"""
        if block_name not in self._block_data:
            self._block_data[block_name] = self._extract_block_data(Application, block_name)
        return self._block_data[block_name]
    
    def _extract_block_data(self, Application, block_name: str) -> Dict[str, any]:
        """블록에서 모든 필요한 데이터를 한 번에 추출"""
        data = {
            'power_kilowatt': None,
            'volumetric_flow_m3_s': None,
            'inlet_pressure_bar': None,
            'outlet_pressure_bar': None,
            'available_nodes': {}
        }
        
        # Output 섹션에서만 노드 정보 수집
        try:
            section_node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Output")
            if section_node and hasattr(section_node, 'Elements'):
                for element in section_node.Elements:
                    try:
                        data['available_nodes'][element.Name] = element.Value
                    except:
                        pass
        except:
            pass
        
        return data
    
    def get_power_data(self, Application, block_name: str, power_unit: Optional[str]) -> Optional[float]:
        """전력 데이터를 캐싱하여 반환"""
        cache_key = f"{block_name}_{power_unit}"
        if cache_key not in self._power_data:
            self._power_data[cache_key] = self._extract_power_data(Application, block_name, power_unit)
        return self._power_data[cache_key]
    
    def _extract_power_data(self, Application, block_name: str, power_unit: Optional[str]) -> Optional[float]:
        """전력 데이터 추출"""
        try:
            node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Output\\WNET")
            if node is None or node.Value is None:
                return None
            return convert_power_to_kw(float(node.Value), power_unit)
        except Exception:
            return None

    def _extract_fan_flow(self, Application, block_name: str, flow_unit: Optional[str]) -> Optional[float]:
        """팬 유량(부피유량) 데이터 추출: FEED_VFLOW → m^3/s 로 변환"""
        try:
            node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Output\\FEED_VFLOW")
            if node is None or node.Value is None:
                return None
            return convert_flow_to_m3_s(float(node.Value), flow_unit)
        except Exception:
            return None
    
    def get_pressure_data(self, Application, block_name: str, pressure_unit: Optional[str], which: str) -> Optional[float]:
        """압력 데이터를 캐싱하여 반환"""
        cache_key = f"{block_name}_{pressure_unit}_{which}"
        if cache_key not in self._pressure_data:
            self._pressure_data[cache_key] = self._extract_pressure_data(Application, block_name, pressure_unit, which)
        return self._pressure_data[cache_key]
    
    def _extract_pressure_data(self, Application, block_name: str, pressure_unit: Optional[str], which: str) -> Optional[float]:
        """압력 데이터 추출"""
        node_name = 'IN_PRES' if which == 'inlet' else 'POC'
        
        try:
            node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Output\\{node_name}")
            if node is None or node.Value is None:
                return None
                
            sval = str(node.Value).strip()
            if not sval:
                return None
                
            return convert_pressure_to_bar(float(sval), pressure_unit)
            
        except Exception:
            return None
    
    def clear_cache(self):
        """캐시 초기화"""
        self._block_data.clear()
        self._unit_data.clear()
        self._pressure_data.clear()
        self._power_data.clear()

# 전역 캐시 인스턴스
_aspen_cache = AspenDataCache()

def clear_aspen_cache():
    """Aspen Plus 데이터 캐시를 초기화하는 함수"""
    _aspen_cache.clear_cache()

def get_cache_stats():
    """캐시 통계를 반환하는 함수"""
    return {
        'block_data_count': len(_aspen_cache._block_data),
        'power_data_count': len(_aspen_cache._power_data),
        'pressure_data_count': len(_aspen_cache._pressure_data),
        'unit_data_count': len(_aspen_cache._unit_data)
    }

# --------------------------------------------------------------------------------------
# Aspen helpers (optional): read power/pressure and convert internally
# --------------------------------------------------------------------------------------

# Power conversion constants moved to unit_converter.py

# _convert_power_to_kilowatt 함수는 unit_converter.py의 convert_power_to_kw로 대체됨

# Volumetric flow conversion to m^3/s
# Flow conversion constants moved to unit_converter.py

# _convert_flow_to_m3_s 함수는 unit_converter.py의 convert_flow_to_m3_s로 대체됨

# Heat duty conversion to W
_HEAT_TO_W: Dict[str, float] = {
    "W": 1.0,
    "Watt": 1.0,
    "kW": 1000.0,
    "MW": 1000000.0,
    "Btu/hr": 0.293071,
    "kcal/hr": 1.163,
    "J/s": 1.0,
}

def _convert_heat_to_w(value: float, unit: Optional[str]) -> Optional[float]:
    """열량 단위를 W로 변환"""
    if unit is None:
        return float(value)
    u = unit.strip()
    factor = _HEAT_TO_W.get(u, None)
    if factor is None:
        # try case-insensitive
        factor = _HEAT_TO_W.get(u.lower(), None)
    if factor is None:
        return float(value)
    return float(value) * factor

# Overall heat transfer coefficient conversion to W/m2-K
_U_TO_W_M2K: Dict[str, float] = {
    "W/m2-K": 1.0,
    "W/m2-C": 1.0,  # K와 C는 온도 차이이므로 동일
    "Btu/hr-ft2-F": 5.678,
    "kcal/hr-m2-C": 1.163,
}

def _convert_u_to_w_m2k(value: float, unit: Optional[str]) -> Optional[float]:
    """전열계수 단위를 W/m2-K로 변환"""
    if unit is None:
        return float(value)
    u = unit.strip()
    factor = _U_TO_W_M2K.get(u, None)
    if factor is None:
        # try case-insensitive
        factor = _U_TO_W_M2K.get(u.lower(), None)
    if factor is None:
        return float(value)
    return float(value) * factor

# Temperature difference conversion to K
_TEMP_DIFF_TO_K: Dict[str, float] = {
    "K": 1.0,
    "C": 1.0,  # 온도 차이는 K와 C가 동일
    "F": 0.555556,  # 5/9
    "R": 0.555556,  # R과 F는 동일한 온도 차이
}

def _convert_temp_diff_to_k(value: float, unit: Optional[str]) -> Optional[float]:
    """온도 차이 단위를 K로 변환"""
    if unit is None:
        return float(value)
    u = unit.strip()
    factor = _TEMP_DIFF_TO_K.get(u, None)
    if factor is None:
        # try case-insensitive
        factor = _TEMP_DIFF_TO_K.get(u.lower(), None)
    if factor is None:
        return float(value)
    return float(value) * factor

# Pressure conversion constants moved to unit_converter.py

# _convert_pressure_to_bar 함수는 unit_converter.py의 convert_pressure_to_bar로 대체됨


def _is_gauge_pressure_unit(unit: Optional[str]) -> bool:
    if not unit:
        return False
    u = unit.strip().lower()
    return u in {"barg", "pag", "kpag", "mpag", "psig", "atmg", "mbarg"}


def _read_pressure_bar(Application, block_name: str, pressure_unit: Optional[str], which: str) -> Optional[float]:
    """캐시를 사용하여 압력 데이터를 읽어옴"""
    return _aspen_cache.get_pressure_data(Application, block_name, pressure_unit, which)


def estimate_pump_cost_from_aspen(
    Application,
    block_name: str,
    power_unit: Optional[str],
    cepci: CEPCIOptions,
    pump_type: PumpType = "centrifugal",
    material: MaterialType = "CS",
    install_factor: float = 1.0
) -> Dict[str, float]:
    try:
        power_kilowatt = _aspen_cache.get_power_data(Application, block_name, power_unit)
        if power_kilowatt is None:
            raise ValueError("WNET not found")
        return estimate_pump_cost(CostInputs(power_kilowatt=power_kilowatt), cepci=cepci, pump_type=pump_type, material=material, install_factor=install_factor)
    except Exception as e:
        raise e


def estimate_compressor_cost_from_aspen(
    Application,
    block_name: str,
    power_unit: Optional[str],
    pressure_unit: Optional[str],
    cepci: CEPCIOptions,
    comp_type: CompressorType = "centrifugal",
    material: MaterialType = "CS",
    install_factor: float = 1.0,
    auto_use_fan_when_low_pressure: bool = True,
    fan_type: FanType = "centrifugal_radial"
) -> Dict[str, float]:
    try:
        power_kilowatt = _aspen_cache.get_power_data(Application, block_name, power_unit)
        if power_kilowatt is None:
            raise ValueError("WNET not found")
        
        pressure_bar = _aspen_cache.get_pressure_data(Application, block_name, pressure_unit, 'outlet')
        
        # If within fan operating pressure (by outlet gauge pressure), compute fan cost
        if auto_use_fan_when_low_pressure and pressure_bar is not None:
            limit = _PRESSURE_LIMIT_BAR.get("fan", 0.16)
            outlet_gauge_bar = pressure_bar - 1.01325 if not _is_gauge_pressure_unit(pressure_unit) else pressure_bar
            if outlet_gauge_bar <= limit:
                # ΔP(bar) 계산 (가능하면)
                inlet_bar = _aspen_cache.get_pressure_data(Application, block_name, pressure_unit, 'inlet')
                dp_bar = None
                if inlet_bar is not None and pressure_bar is not None:
                    try:
                        dp_bar = max(0.0, float(pressure_bar) - float(inlet_bar))
                    except Exception:
                        dp_bar = None
                if ENABLE_DEBUG_OUTPUT:
                    print(f"AUTO FAN SWITCH: inlet_abs={inlet_bar}, outlet_abs={pressure_bar}, ΔP_bar={dp_bar}")
                return estimate_fan_cost(
                    CostInputs(power_kilowatt=power_kilowatt, pressure_bar=pressure_bar, pressure_delta_bar=dp_bar),
                    cepci=cepci,
                    fan_type=fan_type,
                    material=material,
                    install_factor=install_factor,
                )
        return estimate_compressor_cost(CostInputs(power_kilowatt=power_kilowatt, pressure_bar=pressure_bar), cepci=cepci, comp_type=comp_type, material=material, install_factor=install_factor)
    except Exception as e:
        raise e


def estimate_fan_cost_from_aspen(
    Application,
    block_name: str,
    power_unit: Optional[str],
    pressure_unit: Optional[str],
    cepci: CEPCIOptions,
    fan_type: FanType = "centrifugal_radial",
    material: MaterialType = "CS",
    install_factor: float = 1.0
) -> Dict[str, float]:
    try:
        # Fans require volumetric flow (FEED_VFLOW) for cost calculation
        flow_m3_s = _aspen_cache._extract_fan_flow(Application, block_name, power_unit)
        if flow_m3_s is None:
            raise ValueError("FEED_VFLOW not found - fan cost calculation requires volumetric flow data")
        return estimate_fan_cost(CostInputs(volumetric_flow_m3_s=flow_m3_s), cepci=cepci, fan_type=fan_type, material=material, install_factor=install_factor)
    except Exception as e:
        raise e


def estimate_turbine_cost_from_aspen(
    Application,
    block_name: str,
    power_unit: Optional[str],
    cepci: CEPCIOptions,
    turbine_type: TurbineType = "axial",
    material: MaterialType = "CS",
    install_factor: float = 1.0
) -> Dict[str, float]:
    try:
        power_kilowatt = _aspen_cache.get_power_data(Application, block_name, power_unit)
        if power_kilowatt is None:
            raise ValueError("WNET not found")
        # Ensure positive power magnitude for cost estimation
        return estimate_turbine_cost(CostInputs(power_kilowatt=abs(power_kilowatt)), cepci=cepci, turbine_type=turbine_type, material=material, install_factor=install_factor)
    except Exception as e:
        raise e


def _extract_mcompr_stage_data(
    Application,
    block_name: str,
    power_unit: Optional[str],
    pressure_unit: Optional[str],
) -> Dict[str, Dict[str, Optional[float]]]:
    """
    MCompr 블록의 각 단계별 데이터를 추출하는 함수
    Elements를 통해 B_PRES와 BRAKE_POWER의 하위 노드를 파악하여 단계 수를 자동 결정
    """
    stage_data = {}
    
    try:
        # B_PRES 노드에서 하위 요소들 확인
        bpres_node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Output\\B_PRES")
        if bpres_node is None:
            return {}
        
        # Elements에서 단계 번호들 추출
        stage_numbers = []
        try:
            elements = bpres_node.Elements
            for i in range(elements.Count):
                element_name = elements.Item(i).Name
                if element_name.isdigit():
                    stage_numbers.append(int(element_name))
        except Exception as e:
            return {}
        
        if not stage_numbers:
            return {}
        
        stage_numbers.sort()  # 단계 번호 정렬
        
        # 각 단계별 토출 압력 추출
        for stage_num in stage_numbers:
            node_path = f"\\Data\\Blocks\\{block_name}\\Output\\B_PRES\\{stage_num}"
            node = Application.Tree.FindNode(node_path)
            
            if node is not None and node.Value is not None:
                pressure_bar = convert_pressure_to_bar(float(node.Value), pressure_unit)
                stage_data[stage_num] = {'outlet_pressure_bar': pressure_bar}
        
        # BRAKE_POWER 노드에서 하위 요소들 확인
        brake_power_node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Output\\BRAKE_POWER")
        if brake_power_node is not None:
            try:
                elements = brake_power_node.Elements
                for i in range(elements.Count):
                    element_name = elements.Item(i).Name
                    if element_name.isdigit():
                        stage_num = int(element_name)
                        if stage_num in stage_data:  # B_PRES에서 확인된 단계만 처리
                            node_path = f"\\Data\\Blocks\\{block_name}\\Output\\BRAKE_POWER\\{stage_num}"
                            node = Application.Tree.FindNode(node_path)
                            
                            if node is not None and node.Value is not None:
                                power_kilowatt = convert_power_to_kw(float(node.Value), power_unit)
                                stage_data[stage_num]['power_kilowatt'] = power_kilowatt
            except Exception as e:
                # Error reading BRAKE_POWER Elements
                pass
        
        return stage_data
        
    except Exception as e:
        return {}


def _calculate_mcompr_stage_costs(
    stage_data: Dict[str, Dict[str, Optional[float]]],
    cepci: CEPCIOptions,
    material: MaterialType = "CS",
    install_factor: float = 1.0,
    include_intercoolers: bool = True
) -> Dict[str, Dict[str, float]]:
    """
    각 단계별 압축기 비용을 계산하는 함수
    Intercooler 비용도 포함하여 계산
    """
    stage_costs = {}
    intercooler_costs = {}
    total_costs = {"purchased": 0.0, "purchased_adj": 0.0, "bare_module": 0.0, "installed": 0.0}
    
    for stage_num, data in stage_data.items():
        power_kilowatt = data.get('power_kilowatt')
        outlet_pressure_bar = data.get('outlet_pressure_bar')
        
        if power_kilowatt is None:
            continue
            
        try:
            # 각 단계를 개별 압축기로 계산
            compressor_costs = estimate_compressor_cost(
                CostInputs(
                    power_kilowatt=power_kilowatt,
                    pressure_bar=outlet_pressure_bar
                ),
                cepci=cepci,
                comp_type="centrifugal",  # 일반적으로 원심 압축기로 가정
                material=material,
                install_factor=install_factor
            )
            
            stage_costs[f"Stage_{stage_num}"] = compressor_costs
            
            # 총 비용에 압축기 비용 추가
            for key in total_costs:
                total_costs[key] += compressor_costs.get(key, 0.0)
            
            # Intercooler 비용 계산 (마지막 단계 제외)
            if include_intercoolers and int(stage_num) < len(stage_data):
                # MCompr에서는 입구 압력이 없으므로 토출 압력만 사용
                intercooler_cost = _estimate_intercooler_cost(
                    None, outlet_pressure_bar, cepci, material
                )
                if intercooler_cost:
                    intercooler_costs[f"Intercooler_{stage_num}"] = intercooler_cost
                    
                    # 총 비용에 Intercooler 비용 추가
                    for key in total_costs:
                        total_costs[key] += intercooler_cost.get(key, 0.0)
                
        except Exception as e:
            # Error calculating cost for stage
            stage_costs[f"Stage_{stage_num}"] = {"error": str(e)}
    
    return stage_costs, intercooler_costs, total_costs


def _estimate_intercooler_cost(
    inlet_pressure_bar: Optional[float],
    outlet_pressure_bar: Optional[float],
    cepci: CEPCIOptions,
    material: MaterialType = "CS"
) -> Optional[Dict[str, float]]:
    """
    Intercooler 비용을 추정하는 함수 (Placeholder)
    향후 heat exchanger 모듈과 연동될 예정
    MCompr의 경우 입구 압력이 없을 수 있으므로 토출 압력만 사용
    """
    if outlet_pressure_bar is None:
        return None
    
    try:
        # TODO: 향후 heat exchanger 모듈과 연동
        # 현재는 간단한 추정치 제공
        
        # 압력 기반 간단한 추정 (실제로는 열교환기 설계 필요)
        if inlet_pressure_bar is not None:
            pressure_avg = (inlet_pressure_bar + outlet_pressure_bar) / 2
        else:
            # MCompr의 경우 토출 압력만 사용
            pressure_avg = outlet_pressure_bar
        
        # 간단한 추정 공식 (실제 구현 시 heat exchanger 모듈 사용)
        # 이는 임시 추정치이며, 실제로는 열교환기 설계 데이터가 필요
        estimated_cost = 10000 * (pressure_avg / 10) ** 0.6  # 임시 공식
        
        purchased_adj = adjust_cost_to_index(estimated_cost, cepci.base_index, cepci.target_index or cepci.base_index)
        
        # 간단한 BM 인수 적용
        bm_factor = 2.5  # 열교환기 일반적인 BM 인수
        bare = _to_bare_module_cost(purchased_adj, bm_factor)
        
        return {
            "purchased": estimated_cost,
            "purchased_adj": purchased_adj,
            "bare_module": bare,
            "note": "Intercooler cost - placeholder implementation"
        }
        
    except Exception as e:
        # Error estimating intercooler cost
        return None


def estimate_intercooler_cost_from_heat_exchanger_module(
    heat_exchanger_data: Dict[str, float],
    cepci: CEPCIOptions,
    material: MaterialType = "CS"
) -> Dict[str, float]:
    """
    향후 heat exchanger 모듈과 연동될 Intercooler 비용 계산 함수
    heat_exchanger_data: {"area_m2": float, "pressure_bar": float, "temperature_c": float, ...}
    """
    # TODO: heat_exchanger_module.py와 연동
    # 현재는 placeholder 구현
    
    try:
        area_m2 = heat_exchanger_data.get("area_m2", 100)  # 기본값
        pressure_bar = heat_exchanger_data.get("pressure_bar", 10)  # 기본값
        
        # 간단한 추정 공식 (실제로는 heat exchanger 모듈 사용)
        estimated_cost = 500 * area_m2 * (pressure_bar / 10) ** 0.6
        
        purchased_adj = adjust_cost_to_index(estimated_cost, cepci.base_index, cepci.target_index or cepci.base_index)
        
        bm_factor = 2.5
        bare = _to_bare_module_cost(purchased_adj, bm_factor)
        
        return {
            "purchased": estimated_cost,
            "purchased_adj": purchased_adj,
            "bare_module": bare,
            "note": "Intercooler cost from heat exchanger module"
        }
        
    except Exception as e:
        # Error calculating intercooler cost from heat exchanger module
        return {"error": str(e)}


def estimate_mcompr_cost_from_aspen(
    Application,
    block_name: str,
    power_unit: Optional[str],
    pressure_unit: Optional[str],
    cepci: CEPCIOptions,
    material: MaterialType = "CS",
    install_factor: float = 1.0
) -> Dict[str, float]:
    """
    MCompr 블록의 비용을 추정하는 함수
    각 단계별로 개별 압축기 비용을 계산하여 합산
    """
    try:
        # 단계별 데이터 추출
        stage_data = _extract_mcompr_stage_data(Application, block_name, power_unit, pressure_unit)
        
        if not stage_data:
            # 단계별 데이터가 없으면 기존 방식으로 처리
            power_kilowatt = _aspen_cache.get_power_data(Application, block_name, power_unit)
            if power_kilowatt is None:
                raise ValueError("WNET not found")
            
            inlet_bar = _aspen_cache.get_pressure_data(Application, block_name, pressure_unit, 'inlet')
            outlet_bar = _aspen_cache.get_pressure_data(Application, block_name, pressure_unit, 'outlet')
            
            return estimate_mcompr_cost(
                CostInputs(power_kilowatt=power_kilowatt, pressure_bar=outlet_bar), 
                cepci=cepci, 
                material=material, 
                install_factor=install_factor
            )
        
        # 각 단계별 비용 계산 (Intercooler 포함)
        stage_costs, intercooler_costs, total_costs = _calculate_mcompr_stage_costs(
            stage_data, cepci, material, install_factor, include_intercoolers=True
        )
        
        # 결과에 단계별 상세 정보 포함
        result = {
            "purchased": total_costs["purchased"],
            "purchased_adj": total_costs["purchased_adj"],
            "bare_module": total_costs["bare_module"],
            "installed": total_costs["installed"],
            "stage_details": stage_costs,
            "intercooler_details": intercooler_costs,
            "stage_count": len(stage_data),
            "intercooler_count": len(intercooler_costs)
        }
        
        return result
        
    except Exception as e:
        raise e



# High-level batch helper to compute costs for pressure-driven equipment modeled
# as compressors in Aspen (auto-detect fan/turbine where applicable)
def calculate_pressure_device_costs_from_aspen(
    Application,
    block_info: Dict[str, str],
    power_unit: Optional[str],
    pressure_unit: Optional[str],
    material: MaterialType = "CS",
    cepci: CEPCIOptions = CEPCIOptions(target_index=None),
    material_overrides: Optional[Dict[str, MaterialType]] = None,
):
    """
    Parameters
    -----------
    Application: Aspen Plus COM object
    block_info: dict mapping block_name -> category (as parsed in main)
    power_unit: current unit set's POWER unit string
    pressure_unit: current unit set's PRESSURE unit string

    Returns
    -------
    (results, totals): list of per-device dicts, and totals dict
    """
    compressor_blocks = [name for name, cat in block_info.items() if cat in ('Compr', 'MCompr')]
    pump_blocks = [name for name, cat in block_info.items() if cat in ('Pump',)]
    results = []
    totals = {"purchased": 0.0, "purchased_adj": 0.0, "bare_module": 0.0, "installed": 0.0}

    # Pumps
    for name in pump_blocks:
        try:
            m = (material_overrides or {}).get(name, material)
            costs = estimate_pump_cost_from_aspen(Application, name, power_unit, cepci, pump_type="centrifugal", material=m)
            results.append({"name": name, "type": "pump", **costs})
            for k in totals:
                totals[k] += float(costs.get(k, 0.0))
        except Exception as e:
            results.append({"name": name, "type": "error", "error": str(e)})

    # Compressors / Turbines (auto fan)
    for name in compressor_blocks:
        try:
            # inlet/outlet absolute pressures to classify
            inlet_bar = _read_pressure_bar(Application, name, pressure_unit, 'inlet')
            outlet_bar = _read_pressure_bar(Application, name, pressure_unit, 'outlet')

            m = (material_overrides or {}).get(name, material)
            if inlet_bar is not None and outlet_bar is not None and inlet_bar > outlet_bar:
                costs = estimate_turbine_cost_from_aspen(Application, name, power_unit, cepci, turbine_type='axial', material=m)
                dtype = 'turbine'
            else:
                costs = estimate_compressor_cost_from_aspen(Application, name, power_unit, pressure_unit, cepci, comp_type='centrifugal', material=m)
                dtype = 'compressor/fan'

            results.append({"name": name, "type": dtype, **costs})
            for k in totals:
                totals[k] += float(costs.get(k, 0.0))
        except Exception as e:
            results.append({"name": name, "type": "error", "error": str(e)})

    return results, totals


# Convenience: auto-resolve POWER/PRESSURE units from a Unit Set name,
# then compute costs for compressor-like blocks
def _get_unit_type_value(Application, unit_set_name: str, unit_type: str) -> Optional[str]:
    try:
        node = Application.Tree.FindNode(f"\\Data\\Setup\\Units-Sets\\{unit_set_name}\\Unit-Types\\{unit_type}")
        return node.Value if node is not None else None
    except Exception:
        return None


def _extract_all_pressure_device_data(
    Application,
    block_info: Dict[str, str],
    power_unit: Optional[str],
    pressure_unit: Optional[str],
) -> Dict[str, Dict[str, Optional[float]]]:
    """모든 압력 장치의 데이터를 추출"""
    data = {}
    for name, cat in block_info.items():
        if cat in ('Pump', 'Compr'):
            power_kilowatt = _aspen_cache.get_power_data(Application, name, power_unit)
            inlet_bar = _aspen_cache.get_pressure_data(Application, name, pressure_unit, 'inlet')
            outlet_bar = _aspen_cache.get_pressure_data(Application, name, pressure_unit, 'outlet')
            data[name] = {
                'power_kilowatt': power_kilowatt,
                'inlet_bar': inlet_bar,
                'outlet_bar': outlet_bar,
            }
        elif cat == 'MCompr':
            # MCompr는 단계별 데이터를 추출
            stage_data = _extract_mcompr_stage_data(Application, name, power_unit, pressure_unit)
            
            if stage_data:
                # 총 파워 계산
                total_power = sum(data.get('power_kilowatt', 0) for data in stage_data.values() if data.get('power_kilowatt') is not None)
                # 최종 토출 압력 (마지막 단계)
                final_outlet = None
                if stage_data:
                    max_stage = max(stage_data.keys())
                    final_outlet = stage_data[max_stage].get('outlet_pressure_bar')
                
                data[name] = {
                    'power_kilowatt': total_power if total_power > 0 else None,
                    'inlet_bar': None,  # MCompr는 입구 압력 없음
                    'outlet_bar': final_outlet,
                    'stage_data': stage_data  # 단계별 데이터도 포함
                }
            else:
                # 단계별 데이터가 없으면 기본 방식으로 처리
                power_kilowatt = _aspen_cache.get_power_data(Application, name, power_unit)
                outlet_bar = _aspen_cache.get_pressure_data(Application, name, pressure_unit, 'outlet')
                data[name] = {
                    'power_kilowatt': power_kilowatt,
                    'inlet_bar': None,
                    'outlet_bar': outlet_bar,
                }
    return data


def calculate_pressure_device_costs_auto(
    Application,
    block_info: Dict[str, str],
    current_unit_set: Optional[str],
    material: MaterialType = "CS",
    cepci: CEPCIOptions = CEPCIOptions(target_index=None),
    material_overrides: Optional[Dict[str, MaterialType]] = None,
):
    power_unit = None
    pressure_unit = None
    if current_unit_set:
        power_unit = _get_unit_type_value(Application, current_unit_set, 'POWER')
        pressure_unit = _get_unit_type_value(Application, current_unit_set, 'PRESSURE')
    return calculate_pressure_device_costs_from_aspen(
        Application,
        block_info,
        power_unit,
        pressure_unit,
        material=material,
        cepci=cepci,
        material_overrides=material_overrides,
    )


# ----------------------------------------------------------------------------
# Calculation using pre-extracted data (no COM reads inside)
# ----------------------------------------------------------------------------

def calculate_pressure_device_costs_with_data(
    pre_extracted: Dict[str, Dict[str, Optional[float]]],
    block_info: Dict[str, str],
    material: MaterialType = "CS",
    cepci: CEPCIOptions = CEPCIOptions(target_index=None),
    material_overrides: Optional[Dict[str, MaterialType]] = None,
    type_overrides: Optional[Dict[str, str]] = None,
    subtype_overrides: Optional[Dict[str, str]] = None,
):
    results = []
    totals = {"purchased": 0.0, "purchased_adj": 0.0, "bare_module": 0.0, "installed": 0.0}
    for name, cat in block_info.items():
        pdata = pre_extracted.get(name, {})
        pw = pdata.get('power_kilowatt')
        inlet_bar = pdata.get('inlet_bar')
        outlet_bar = pdata.get('outlet_bar')
        pressure_delta_bar = pdata.get('pressure_delta_bar')
        if pressure_delta_bar is None and inlet_bar is not None and outlet_bar is not None:
            try:
                pressure_delta_bar = max(0.0, float(outlet_bar) - float(inlet_bar))
                if ENABLE_DEBUG_OUTPUT:
                    print(f"WITH_DATA: computed ΔP_bar for {name} = {pressure_delta_bar}")
            except Exception:
                pressure_delta_bar = None
        m = (material_overrides or {}).get(name, material)
        # error 필드가 있는 경우 건너뛰기
        if 'error' in pdata:
            results.append({"name": name, "type": "error", "error": pdata['error']})
            continue
            
        try:
            if cat == 'Pump':
                if pw is None:
                    raise ValueError('Missing power_kilowatt for pump')
                
                print(f"\n🔧 PUMP CALCULATION: {name}")
                selected_subtype = (subtype_overrides or {}).get(name, 'centrifugal')
                print(f"   Type: {selected_subtype}, Material: {m}, Power: {pw:.2f} kW")
                
                costs = estimate_pump_cost(
                    CostInputs(power_kilowatt=pw, pressure_bar=outlet_bar),
                    cepci=cepci,
                    pump_type=selected_subtype,
                    material=m,
                )
                dtype = 'pump'
                
                print(f"   ✅ Cost: ${costs['purchased']:,.2f} → ${costs['bare_module']:,.2f}")
            elif cat == 'Compr':
                if pw is None:
                    raise ValueError('Missing power_kilowatt for compressor')
                
                print(f"\n🔧 COMPRESSOR CALCULATION: {name}")
                
                # 사용자가 선택한 타입이 있으면 우선 사용
                selected_type = (type_overrides or {}).get(name)
                selected_subtype = (subtype_overrides or {}).get(name)
                
                if selected_type:
                    print(f"   User-selected type: {selected_type} ({selected_subtype})")
                    
                    if selected_type == 'fan':
                        fan_type = selected_subtype if selected_subtype else 'centrifugal_radial'
                        vflow = pdata.get('volumetric_flow_m3_s')
                        if vflow is None:
                            raise ValueError(f"Fan {name} requires volumetric flow data (FEED_VFLOW) for cost calculation")
                        print(f"   Type: Fan ({fan_type}), Material: {m}, Flow: {vflow:.2f} m³/s")
                        costs = estimate_fan_cost(CostInputs(volumetric_flow_m3_s=vflow, pressure_bar=outlet_bar, pressure_delta_bar=pressure_delta_bar), cepci=cepci, fan_type=fan_type, material=m)
                        dtype = 'fan'
                        print(f"   ✅ Cost: ${costs['purchased']:,.2f} → ${costs['bare_module']:,.2f}")
                    elif selected_type == 'compressor':
                        comp_type = selected_subtype if selected_subtype else 'centrifugal'
                        print(f"   Type: Compressor ({comp_type}), Material: {m}, Power: {pw:.2f} kW")
                        costs = estimate_compressor_cost(CostInputs(power_kilowatt=pw, pressure_bar=outlet_bar), cepci=cepci, comp_type=comp_type, material=m)
                        dtype = 'compressor'
                        print(f"   ✅ Cost: ${costs['purchased']:,.2f} → ${costs['bare_module']:,.2f}")
                    elif selected_type == 'turbine':
                        turbine_type = selected_subtype if selected_subtype else 'axial'
                        print(f"   Type: Turbine ({turbine_type}), Material: {m}, Power: {pw:.2f} kW")
                        costs = estimate_turbine_cost(CostInputs(power_kilowatt=pw), cepci=cepci, turbine_type=turbine_type, material=m)
                        dtype = 'turbine'
                        print(f"   ✅ Cost: ${costs['purchased']:,.2f} → ${costs['bare_module']:,.2f}")
                    else:
                        # 잘못된 타입이면 기본 로직 사용
                        selected_type = None
                
                # 사용자 선택이 없으면 기존 자동 분류 로직 사용
                if not selected_type:
                    print(f"   Auto-classification based on pressure conditions")
                    
                    if inlet_bar is not None and outlet_bar is not None and inlet_bar > outlet_bar:
                        print(f"   Detected: Turbine (pressure drop: {inlet_bar:.2f} → {outlet_bar:.2f} bar)")
                        print(f"   Type: Turbine (axial), Material: {m}, Power: {pw:.2f} kW")
                        costs = estimate_turbine_cost(CostInputs(power_kilowatt=pw), cepci=cepci, turbine_type='axial', material=m)
                        dtype = 'turbine'
                        print(f"   ✅ Cost: ${costs['purchased']:,.2f} → ${costs['bare_module']:,.2f}")
                    else:
                        # 압력 상승이 낮으면 팬, 높으면 압축기로 분류
                        if outlet_bar is not None and inlet_bar is not None:
                            pressure_rise = outlet_bar - inlet_bar
                            if pressure_rise <= 0.16:  # 팬 범위
                                print(f"   Detected: Fan (pressure rise: {pressure_rise:.3f} bar ≤ 0.16 bar)")
                                vflow = pdata.get('volumetric_flow_m3_s')
                                if vflow is None:
                                    raise ValueError(f"Fan {name} requires volumetric flow data (FEED_VFLOW) for cost calculation")
                                print(f"   Type: Fan (centrifugal), Material: {m}, Flow: {vflow:.2f} m³/s")
                                costs = estimate_fan_cost(CostInputs(volumetric_flow_m3_s=vflow, pressure_bar=outlet_bar, pressure_delta_bar=pressure_delta_bar), cepci=cepci, material=m)
                                dtype = 'fan'
                                print(f"   ✅ Cost: ${costs['purchased']:,.2f} → ${costs['bare_module']:,.2f}")
                            else:
                                print(f"   Detected: Compressor (pressure rise: {pressure_rise:.3f} bar > 0.16 bar)")
                                print(f"   Type: Compressor (centrifugal), Material: {m}, Power: {pw:.2f} kW")
                                costs = estimate_compressor_cost(CostInputs(power_kilowatt=pw, pressure_bar=outlet_bar), cepci=cepci, comp_type='centrifugal', material=m)
                                dtype = 'compressor'
                                print(f"   ✅ Cost: ${costs['purchased']:,.2f} → ${costs['bare_module']:,.2f}")
                        else:
                            # 압력 정보가 없으면 압축기로 가정
                            print(f"   Detected: Compressor (no pressure data, default assumption)")
                            print(f"   Type: Compressor (centrifugal), Material: {m}, Power: {pw:.2f} kW")
                            costs = estimate_compressor_cost(CostInputs(power_kilowatt=pw, pressure_bar=outlet_bar), cepci=cepci, comp_type='centrifugal', material=m)
                            dtype = 'compressor'
                            print(f"   ✅ Cost: ${costs['purchased']:,.2f} → ${costs['bare_module']:,.2f}")
            elif cat == 'MCompr':
                if pw is None:
                    raise ValueError('Missing power_kilowatt for multi-stage compressor')
                
                print(f"\n{'='*80}")
                print(f"MULTI-STAGE COMPRESSOR CALCULATION: {name}")
                print(f"{'='*80}")
                
                # MCompr: 단계별 데이터가 있으면 단계별 계산, 없으면 기본 계산
                stage_data = pdata.get('stage_data')
                
                if stage_data:
                    print(f"📊 Stage-by-stage calculation (Total stages: {len(stage_data)})")
                    print(f"🔧 Material: {m}, CEPCI: {cepci.target_index}")
                    
                    # 단계별 비용 계산
                    stage_costs, intercooler_costs, total_costs = _calculate_mcompr_stage_costs(
                        stage_data, cepci, m, install_factor=1.0, include_intercoolers=True
                    )
                    costs = {
                        "purchased": total_costs["purchased"],
                        "purchased_adj": total_costs["purchased_adj"],
                        "bare_module": total_costs["bare_module"],
                        "installed": total_costs["installed"],
                        "stage_details": stage_costs,
                        "intercooler_details": intercooler_costs,
                        "stage_count": len(stage_data),
                        "intercooler_count": len(intercooler_costs)
                    }
                    dtype = 'multi-stage compressor (stage-by-stage)'
                    
                    print(f"✅ Calculation completed for {name}")
                    print(f"💰 Total purchased cost: ${total_costs['purchased']:,.2f}")
                    print(f"🏗️  Total bare module cost: ${total_costs['bare_module']:,.2f}")
                else:
                    print(f"📊 Simplified calculation (No stage data available)")
                    print(f"🔧 Material: {m}, CEPCI: {cepci.target_index}")
                    
                    # 단계별 데이터가 없으면 기본 계산
                    costs = estimate_mcompr_cost(CostInputs(power_kilowatt=pw, pressure_bar=outlet_bar), cepci=cepci, material=m)
                    dtype = 'multi-stage compressor (simplified)'
                    
                    print(f"✅ Calculation completed for {name}")
                    print(f"💰 Purchased cost: ${costs['purchased']:,.2f}")
                    print(f"🏗️  Bare module cost: ${costs['bare_module']:,.2f}")
                
                print(f"{'='*80}")
            else:
                continue
            results.append({"name": name, "type": dtype, **costs})
            for k in totals:
                totals[k] += float(costs.get(k, 0.0))
        except ValueError as e:
            error_msg = str(e)
            if "under limit" in error_msg:
                # 최소 크기 제한을 벗어나는 경우 특별 처리
                results.append({
                    "name": name, 
                    "type": f"{cat.lower()}(under limit)", 
                    "error": error_msg
                })
            else:
                # 다른 오류는 기존 방식으로 처리
                results.append({"name": name, "type": "error", "error": error_msg})
        except Exception as e:
            results.append({"name": name, "type": "error", "error": str(e)})
    return results, totals


def calculate_heat_exchanger_costs_with_data(
    pre_extracted: Dict[str, Dict],
    block_info: Dict[str, str],
    material: str = 'CS',
    cepci: CEPCIOptions = CEPCIOptions(),
    material_overrides: Optional[Dict[str, str]] = None,
    type_overrides: Optional[Dict[str, str]] = None,
    subtype_overrides: Optional[Dict[str, str]] = None,
) -> Tuple[List[Dict], Dict[str, float]]:
    """
    미리 추출된 데이터를 사용하여 열교환기 비용을 계산합니다.
    
    Args:
        pre_extracted: 미리 추출된 장치 데이터 딕셔너리
        block_info: 블록 정보 딕셔너리
        material: 기본 재질
        cepci: CEPCI 옵션
        material_overrides: 재질 오버라이드
        type_overrides: 타입 오버라이드
        subtype_overrides: 세부 타입 오버라이드
    
    Returns:
        Tuple[List[Dict], Dict[str, float]]: (비용 결과 리스트, 총계 딕셔너리)
    """
    results = []
    totals = {
        "purchased": 0.0,
        "purchased_adj": 0.0,
        "bare_module": 0.0,
        "installed": 0.0,
    }
    
    # 열교환기 카테고리 정의
    hx_cats = {"Heater", "Cooler", "HeatX", "Condenser"}
    
    # 열교환기 블록 찾기
    hx_blocks = {name: cat for name, cat in block_info.items() if cat in hx_cats}
    
    if not hx_blocks:
        # 열교환기가 없으면 빈 결과 반환
        return results, totals
    
    for name, cat in hx_blocks.items():
        try:
            pdata = pre_extracted.get(name, {})
            if not pdata:
                raise ValueError(f"No pre-extracted data found for {name}")
            
            # 기본값 설정
            m = material_overrides.get(name) if material_overrides else material
            hx_type = type_overrides.get(name) if type_overrides else 'fixed_tube'
            hx_subtype = subtype_overrides.get(name) if subtype_overrides else 'fixed_tube'
            
            # 열교환기 입력 데이터 추출
            heat_duty_W = pdata.get('heat_duty_W')
            overall_U_W_m2K = pdata.get('overall_U_W_m2K')
            lmtd_K = pdata.get('lmtd_K')
            area_m2 = pdata.get('area_m2')
            
            if heat_duty_W is None:
                raise ValueError(f'Missing heat_duty_W for heat exchanger {name}')
            
            # 면적 계산 (필요한 경우)
            if area_m2 is None and overall_U_W_m2K is not None and lmtd_K is not None:
                area_m2 = heat_duty_W / (overall_U_W_m2K * lmtd_K)
            
            if area_m2 is None:
                raise ValueError(f'Missing area_m2 for heat exchanger {name}')
            
            # 비용 계산
            costs = estimate_heat_exchanger_cost(
                inputs=CostInputs(
                    heat_duty_W=heat_duty_W,
                    overall_U_W_m2K=overall_U_W_m2K,
                    lmtd_K=lmtd_K,
                    area_m2=area_m2,
                ),
                cepci=cepci,
                hx_type=hx_type,
                material_shell=m,
                material_tube=m,  # 기본값으로 동일한 재질 사용
            )
            
            dtype = f'{hx_type} heat exchanger'
            results.append({"name": name, "type": dtype, **costs})
            
            for k in totals:
                totals[k] += float(costs.get(k, 0.0))
                
        except ValueError as e:
            error_msg = str(e)
            results.append({"name": name, "type": "error", "error": error_msg})
        except Exception as e:
            results.append({"name": name, "type": "error", "error": str(e)})
    
    return results, totals


# ----------------------------------------------------------------------------
# Preview helpers (to show extracted data before cost calculation)
# ----------------------------------------------------------------------------

def preview_pressure_devices_from_aspen(
    Application,
    block_info: Dict[str, str],
    power_unit: Optional[str],
    pressure_unit: Optional[str],
    flow_unit: Optional[str] = None,
) -> List[Dict]:
    """
    Aspen에서 압력 관련 장치 정보를 추출하고 미리보기를 생성합니다.

    Args:
        Application: Aspen COM 객체
        block_info (Dict[str, str]): {장치 이름: 카테고리} 형태의 딕셔너리
        power_unit (Optional[str]): 전력 단위
        pressure_unit (Optional[str]): 압력 단위
        flow_unit (Optional[str]): 유량 단위

    Returns:
        List[Dict]: 장치 정보가 담긴 딕셔너리 리스트
    """
    # 장치 이름 오름차순으로 정렬
    all_blocks = sorted(block_info.items(), key=lambda x: x[0])
    preview = []

    for name, cat in all_blocks:
        if cat == 'Pump':
            try:
                pw = _aspen_cache.get_power_data(Application, name, power_unit)
                inlet_bar = _aspen_cache.get_pressure_data(Application, name, pressure_unit, 'inlet')
                outlet_bar = _aspen_cache.get_pressure_data(Application, name, pressure_unit, 'outlet')
                dp_bar = None
                if inlet_bar is not None and outlet_bar is not None:
                    try:
                        dp_bar = max(0.0, float(outlet_bar) - float(inlet_bar))
                    except Exception:
                        dp_bar = None
                preview.append({
                    "name": name,
                    "category": "Pump",
                    "power_kilowatt": pw,
                    "inlet_bar": inlet_bar,
                    "outlet_bar": outlet_bar,
                    "pressure_delta_bar": dp_bar,
                    "suggested": "pump",
                    "material": "CS",  # 기본 재질
                    "selected_type": "pump",  # 선택된 타입
                    "selected_subtype": "centrifugal",  # 선택된 세부 타입
                })
            except Exception as e:
                preview.append({"name": name, "category": "Pump", "error": f"failed to read: {str(e)}"})

        elif cat == 'Compr':
            try:
                pw = _aspen_cache.get_power_data(Application, name, power_unit)
                inlet_bar = _aspen_cache.get_pressure_data(Application, name, pressure_unit, 'inlet')
                outlet_bar = _aspen_cache.get_pressure_data(Application, name, pressure_unit, 'outlet')
                suggested = None
                if inlet_bar is not None and outlet_bar is not None:
                    if outlet_bar > inlet_bar:
                        pressure_rise = outlet_bar - inlet_bar
                        if pressure_rise <= 0.16:  # 팬 범위
                            suggested = 'fan'
                        else:
                            suggested = 'compressor'
                    elif inlet_bar > outlet_bar:
                        suggested = 'turbine'

                # 팬일 가능성이 있는 경우 유량도 시도
                vflow = None
                if suggested == 'fan':
                    vflow = _aspen_cache._extract_fan_flow(Application, name, flow_unit)

                preview.append({
                    "name": name,
                    "category": "Compr",
                    "power_kilowatt": pw,
                    "volumetric_flow_m3_s": vflow,
                    "inlet_bar": inlet_bar,
                    "outlet_bar": outlet_bar,
                    "pressure_delta_bar": (outlet_bar - inlet_bar) if (inlet_bar is not None and outlet_bar is not None) else None,
                    "suggested": suggested,
                    "material": "CS",  # 기본 재질
                    "selected_type": suggested,  # 선택된 타입 (기본값은 suggested)
                    "selected_subtype": (
                        "centrifugal" if suggested == "compressor"
                        else "centrifugal_radial" if suggested == "fan"
                        else "axial"
                    ),  # 기본 세부 타입
                })
            except Exception as e:
                preview.append({
                    "name": name,
                    "category": "Compr",
                    "error": f"failed to read: {str(e)}"
                })

        elif cat == 'MCompr':
            try:
                # 기본 정보
                pw = _aspen_cache.get_power_data(Application, name, power_unit)

                # 단계별 정보 추출
                stage_data = _extract_mcompr_stage_data(Application, name, power_unit, pressure_unit)

                # 최종 토출 압력 계산 (가장 마지막 단계의 토출 압력)
                final_outlet_bar = None
                if stage_data:
                    max_stage = max(stage_data.keys())
                    final_outlet_bar = stage_data[max_stage].get('outlet_pressure_bar')

                preview.append({
                    "name": name,
                    "category": "MCompr",
                    "power_kilowatt": pw,
                    "suggested": "multi-stage compressor",
                    "stage_count": len(stage_data),
                    "final_outlet_bar": final_outlet_bar,
                    "stage_data": stage_data,  # stage_data도 포함
                    "material": "CS",  # 기본 재질
                    "selected_type": "multi-stage compressor",  # 선택된 타입
                    "selected_subtype": "centrifugal",  # 선택된 세부 타입
                })
            except Exception as e:
                preview.append({"name": name, "category": "MCompr", "error": f"failed to read: {str(e)}"})

    return preview


def preview_pressure_devices_auto(
    Application,
    block_info: Dict[str, str],
    current_unit_set: Optional[str],
):
    power_unit = None
    pressure_unit = None
    flow_unit = None
    if current_unit_set:
        power_unit = _get_unit_type_value(Application, current_unit_set, 'POWER')
        pressure_unit = _get_unit_type_value(Application, current_unit_set, 'PRESSURE')
        flow_unit = _get_unit_type_value(Application, current_unit_set, 'VOLUME-FLOW')
    return preview_pressure_devices_from_aspen(Application, block_info, power_unit, pressure_unit, flow_unit)


def get_device_type_options(category: str) -> Dict[str, List[str]]:
    """장치 카테고리별로 선택 가능한 타입과 세부 타입 목록을 반환합니다."""
    if category == 'Pump':
        return {
            'pump': ['centrifugal', 'reciprocating']
        }
    elif category == 'Compr':
        return {
            'fan': ['centrifugal_radial', 'centrifugal_backward', 'centrifugal_forward', 'axial'],
            'compressor': ['centrifugal', 'axial', 'reciprocating'],
            'turbine': ['axial', 'radial']
        }
    elif category == 'MCompr':
        return {
            'multi-stage compressor': ['centrifugal', 'axial']
        }
    else:
        return {}


def print_preview_results(preview: list, Application, power_unit: Optional[str], pressure_unit: Optional[str]):
    """
    프리뷰 결과를 포맷팅하여 출력하는 함수
    """
    print("\n" + "="*60)
    print("PREVIEW: PRESSURE-DRIVEN DEVICES (extracted data)")
    print("="*60)
    
    for p in preview:
        name = p.get('name')
        cat = p.get('category')
        
        # 오류가 있는 장치 처리
        if 'error' in p:
            error_msg = p.get('error', 'Unknown error')
            print(f"{name:20s} | {cat:12s} | ERROR: {error_msg}")
            continue
            
        pw = p.get('power_kilowatt')
        inlet_bar = p.get('inlet_bar')
        outlet_bar = p.get('outlet_bar')
        sug = p.get('suggested')
        material = p.get('material', 'CS')  # 기본값 CS
        selected_type = p.get('selected_type', sug)  # 기본값은 suggested
        selected_subtype = p.get('selected_subtype', 'N/A')  # 세부 타입
        
        if cat == 'MCompr':
            # MCompr의 경우 preview 데이터에서 단계 수와 최종 토출 압력 사용
            stage_count = p.get('stage_count', 0)
            final_outlet_bar = p.get('final_outlet_bar')
            pw_str = f"{pw:,.2f}" if pw is not None else "NA"
            final_outlet_str = f"{final_outlet_bar:,.2f}" if final_outlet_bar is not None else "NA"
            print(f"{name:20s} | {cat:12s} | P={pw_str} kW | Stages={stage_count} | Pout_final={final_outlet_str} bar | Material={material} | Type={selected_type} | Subtype={selected_subtype}")
        else:
            # Pump, Compr의 경우 기존 형식 (게이지 압력 제외)
            vflow = p.get('volumetric_flow_m3_s')
            if selected_type == 'fan' and vflow is not None:
                vflow_str = f"{vflow:,.2f}" if vflow is not None else "NA"
                inlet_str = f"{inlet_bar:,.2f}" if inlet_bar is not None else "NA"
                outlet_str = f"{outlet_bar:,.2f}" if outlet_bar is not None else "NA"
                print(f"{name:20s} | {cat:12s} | Q={vflow_str} m3/s | Pin={inlet_str} bar | Pout={outlet_str} bar | Material={material} | Type={selected_type} | Subtype={selected_subtype}")
            else:
                pw_str = f"{pw:,.2f}" if pw is not None else "NA"
                inlet_str = f"{inlet_bar:,.2f}" if inlet_bar is not None else "NA"
                outlet_str = f"{outlet_bar:,.2f}" if outlet_bar is not None else "NA"
                print(f"{name:20s} | {cat:12s} | P={pw_str} kW | Pin={inlet_str} bar | Pout={outlet_str} bar | Material={material} | Type={selected_type} | Subtype={selected_subtype}")


# ----------------------------------------------------------------------------
# Heat exchanger preview helpers
# ----------------------------------------------------------------------------

def _read_float_node(Application, path: str) -> Optional[float]:
    try:
        node = Application.Tree.FindNode(path)
        if node is None or node.Value is None:
            return None
        sval = str(node.Value).strip()
        if not sval:
            return None
        return float(sval)
    except Exception:
        return None

def _read_float_node_with_unit(Application, path: str) -> Tuple[Optional[float], Optional[str]]:
    """Aspen 노드에서 float 값과 단위를 함께 읽어옵니다."""
    try:
        node = Application.Tree.FindNode(path)
        if node is None or node.Value is None:
            return None, None
        sval = str(node.Value).strip()
        if not sval:
            return None, None
        
        # 단위 정보 확인
        unit = None
        try:
            unit_node = Application.Tree.FindNode(path + "\\Unit")
            if unit_node and unit_node.Value:
                unit = unit_node.Value
        except Exception:
            pass
        
        return float(sval), unit
    except Exception:
        return None, None


def debug_aspen_units(Application, block_name: str) -> None:
    """Aspen에서 특정 블록의 단위 정보를 디버깅합니다."""
    print(f"\n=== 단위 디버깅: {block_name} ===")
    
    # 전력 단위 확인
    try:
        power_node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Output\\POWER")
        power_unit_node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Output\\POWER\\Unit")
        print(f"POWER: {power_node.Value if power_node else 'None'} [{power_unit_node.Value if power_unit_node else 'No Unit'}]")
    except Exception as e:
        print(f"POWER 단위 확인 실패: {e}")
    
    # 압력 단위 확인
    try:
        inlet_pressure_node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Output\\INLET_PRESSURE")
        inlet_pressure_unit_node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Output\\INLET_PRESSURE\\Unit")
        print(f"INLET_PRESSURE: {inlet_pressure_node.Value if inlet_pressure_node else 'None'} [{inlet_pressure_unit_node.Value if inlet_pressure_unit_node else 'No Unit'}]")
    except Exception as e:
        print(f"INLET_PRESSURE 단위 확인 실패: {e}")
    
    try:
        outlet_pressure_node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Output\\OUTLET_PRESSURE")
        outlet_pressure_unit_node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Output\\OUTLET_PRESSURE\\Unit")
        print(f"OUTLET_PRESSURE: {outlet_pressure_node.Value if outlet_pressure_node else 'None'} [{outlet_pressure_unit_node.Value if outlet_pressure_unit_node else 'No Unit'}]")
    except Exception as e:
        print(f"OUTLET_PRESSURE 단위 확인 실패: {e}")
    
    # 유량 단위 확인 (팬의 경우)
    try:
        flow_node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Output\\FEED_VFLOW")
        flow_unit_node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Output\\FEED_VFLOW\\Unit")
        print(f"FEED_VFLOW: {flow_node.Value if flow_node else 'None'} [{flow_unit_node.Value if flow_unit_node else 'No Unit'}]")
    except Exception as e:
        print(f"FEED_VFLOW 단위 확인 실패: {e}")
    
    # 열교환기 관련 단위 확인
    try:
        duty_node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Output\\HX_DUTY")
        duty_unit_node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Output\\HX_DUTY\\Unit")
        print(f"HX_DUTY: {duty_node.Value if duty_node else 'None'} [{duty_unit_node.Value if duty_unit_node else 'No Unit'}]")
    except Exception as e:
        print(f"HX_DUTY 단위 확인 실패: {e}")
    
    try:
        u_node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Input\\U")
        u_unit_node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Input\\U\\Unit")
        print(f"U: {u_node.Value if u_node else 'None'} [{u_unit_node.Value if u_unit_node else 'No Unit'}]")
    except Exception as e:
        print(f"U 단위 확인 실패: {e}")
    
    try:
        lmtd_node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Output\\HX_DTLM")
        lmtd_unit_node = Application.Tree.FindNode(f"\\Data\\Blocks\\{block_name}\\Output\\HX_DTLM\\Unit")
        print(f"HX_DTLM: {lmtd_node.Value if lmtd_node else 'None'} [{lmtd_unit_node.Value if lmtd_unit_node else 'No Unit'}]")
    except Exception as e:
        print(f"HX_DTLM 단위 확인 실패: {e}")
    
    print("=" * 50)

def preview_heat_exchangers_from_aspen(
    Application,
    block_info: Dict[str, str],
    unit_set: Optional[str],
) -> List[Dict]:
    """열교환기 후보 블록을 스캔하여 Q/U/LMTD/Area 미리보기를 생성합니다."""
    hx_cats = {"Heater", "Cooler", "HeatX", "Condenser"}
    all_blocks = sorted(block_info.items(), key=lambda x: x[0])
    preview: List[Dict] = []
    for name, cat in all_blocks:
        if cat not in hx_cats:
            continue
        q_w = _read_float_node(Application, f"\\Data\\Blocks\\{name}\\Output\\HX_DUTY")
        u = _read_float_node(Application, f"\\Data\\Blocks\\{name}\\Input\\U")
        lmtd = _read_float_node(Application, f"\\Data\\Blocks\\{name}\\Output\\HX_DTLM")
        
        # 면적은 항상 계산 (HX_AREA 노드는 존재하지 않음)
        area = None
        if q_w is not None and u is not None and lmtd is not None:
            area = q_w / (u * lmtd)
        
        # 간이 제안 타입: 면적/온도 유무로 fixed_tube 기본
        suggested = "fixed_tube"
        preview.append({
            "name": name,
            "category": "HeatExchanger",
            "q_w": q_w,
            "u_W_m2K": u,
            "lmtd_K": lmtd,
            "area_m2": area,
            "suggested": suggested,
        })
    return preview


def preview_heat_exchangers_auto(Application, block_info: Dict[str, str], current_unit_set: Optional[str]) -> List[Dict]:
    return preview_heat_exchangers_from_aspen(Application, block_info, current_unit_set)


def print_preview_hx_results(preview: List[Dict]) -> None:
    print("\n" + "="*60)
    print("PREVIEW: HEAT EXCHANGERS (extracted data)")
    print("="*60)
    for p in preview:
        name = p.get('name')
        q = p.get('q_w')
        u = p.get('u_W_m2K')
        lmtd = p.get('lmtd_K')
        area = p.get('area_m2')
        q_str = f"{q:,.2f}" if q is not None else "NA"
        u_str = f"{u:,.2f}" if u is not None else "NA"
        lmtd_str = f"{lmtd:,.2f}" if lmtd is not None else "NA"
        area_str = f"{area:,.2f}" if area is not None else "NA"
        print(f"{name:20s} | Q={q_str} W | U={u_str} W/m2-K | LMTD={lmtd_str} K | A={area_str} m2")


# ----------------------------------------------------------------------------
# Unified preview functions (pressure devices + heat exchangers)
# ----------------------------------------------------------------------------

def preview_all_devices_from_aspen(
    Application,
    block_info: Dict[str, str],
    power_unit: Optional[str],
    pressure_unit: Optional[str],
    flow_unit: Optional[str] = None,
    heat_unit: Optional[str] = None,
    temperature_unit: Optional[str] = None,
) -> Tuple[List[Dict], List[Dict]]:
    """
    압력 장치와 열교환기를 모두 포함한 통합 프리뷰를 생성합니다.
    
    Returns:
        Tuple[List[Dict], List[Dict]]: (압력장치 프리뷰, 열교환기 프리뷰)
    """
    # 장치 이름 오름차순으로 정렬
    all_blocks = sorted(block_info.items(), key=lambda x: x[0])
    
    pressure_preview = []
    hx_preview = []
    
    # 열교환기 카테고리 정의
    hx_cats = {"Heater", "Cooler", "HeatX", "Condenser"}
    
    for name, cat in all_blocks:
        # 압력 관련 장치 처리
        if cat in ['Pump', 'Compr', 'MCompr']:
            try:
                if cat == 'Pump':
                    pw = _aspen_cache.get_power_data(Application, name, power_unit)
                    inlet_bar = _aspen_cache.get_pressure_data(Application, name, pressure_unit, 'inlet')
                    outlet_bar = _aspen_cache.get_pressure_data(Application, name, pressure_unit, 'outlet')
                    dp_bar = None
                    if inlet_bar is not None and outlet_bar is not None:
                        try:
                            dp_bar = max(0.0, float(outlet_bar) - float(inlet_bar))
                        except Exception:
                            dp_bar = None
                    pressure_preview.append({
                        "name": name,
                        "category": "Pump",
                        "power_kilowatt": pw,
                        "inlet_bar": inlet_bar,
                        "outlet_bar": outlet_bar,
                        "pressure_delta_bar": dp_bar,
                        "suggested": "pump",
                        "material": "CS",
                        "selected_type": "pump",
                        "selected_subtype": "centrifugal",
                    })
                elif cat == 'Compr':
                    pw = _aspen_cache.get_power_data(Application, name, power_unit)
                    inlet_bar = _aspen_cache.get_pressure_data(Application, name, pressure_unit, 'inlet')
                    outlet_bar = _aspen_cache.get_pressure_data(Application, name, pressure_unit, 'outlet')
                    suggested = None
                    if inlet_bar is not None and outlet_bar is not None:
                        if outlet_bar > inlet_bar:
                            pressure_rise = outlet_bar - inlet_bar
                            if pressure_rise <= 0.16:  # 팬 범위
                                suggested = 'fan'
                            else:
                                suggested = 'compressor'
                        elif inlet_bar > outlet_bar:
                            suggested = 'turbine'

                    # 팬일 가능성이 있는 경우 유량도 시도
                    vflow = None
                    if suggested == 'fan':
                        vflow = _aspen_cache._extract_fan_flow(Application, name, flow_unit)
                    pressure_preview.append({
                        "name": name,
                        "category": "Compr",
                        "power_kilowatt": pw,
                        "inlet_bar": inlet_bar,
                        "outlet_bar": outlet_bar,
                        "volumetric_flow_m3_s": vflow,
                        "pressure_delta_bar": (outlet_bar - inlet_bar) if (inlet_bar is not None and outlet_bar is not None) else None,
                        "suggested": suggested,
                        "material": "CS",
                        "selected_type": suggested,
                        "selected_subtype": ("centrifugal" if suggested == "compressor" else "centrifugal_radial" if suggested == "fan" else "axial"),
                    })
                elif cat == 'MCompr':
                    pw = _aspen_cache.get_power_data(Application, name, power_unit)
                    stage_data = _extract_mcompr_stage_data(Application, name, power_unit, pressure_unit)
                    final_outlet_bar = None
                    if stage_data:
                        max_stage = max(stage_data.keys())
                        final_outlet_bar = stage_data[max_stage].get('outlet_pressure_bar')
                    pressure_preview.append({
                        "name": name,
                        "category": "MCompr",
                        "power_kilowatt": pw,
                        "suggested": "multi-stage compressor",
                        "stage_count": len(stage_data),
                        "final_outlet_bar": final_outlet_bar,
                        "stage_data": stage_data,
                        "material": "CS",
                        "selected_type": "multi-stage compressor",
                        "selected_subtype": "centrifugal",
                    })
            except Exception as e:
                # 예외 발생 시 기존 방식대로 error 필드만 있는 딕셔너리 생성
                if cat == 'Pump':
                    pressure_preview.append({"name": name, "category": "Pump", "error": f"failed to read: {str(e)}"})
                elif cat == 'Compr':
                    pressure_preview.append({"name": name, "category": "Compr", "error": f"failed to read: {str(e)}"})
                elif cat == 'MCompr':
                    pressure_preview.append({"name": name, "category": "MCompr", "error": f"failed to read: {str(e)}"})
        
        # 열교환기 처리
        elif cat in hx_cats:
            try:
                q_w_raw = _read_float_node(Application, f"\\Data\\Blocks\\{name}\\Output\\HX_DUTY")
                u_raw = _read_float_node(Application, f"\\Data\\Blocks\\{name}\\Input\\U")
                lmtd_raw = _read_float_node(Application, f"\\Data\\Blocks\\{name}\\Output\\HX_DTLM")
                
                # 단위 변환 적용
                q_w = _convert_heat_to_w(q_w_raw, heat_unit) if q_w_raw is not None else None
                u = _convert_u_to_w_m2k(u_raw, None) if u_raw is not None else None  # U는 보통 단위가 없음
                
                # HX_DTLM은 온도 차이이므로 온도 차이 전용 변환 사용
                # 온도 차이는 K와 C가 동일하므로 단위가 없어도 안전
                lmtd = lmtd_raw  # 온도 차이는 단위 변환이 필요 없음 (K = C)
                
                # 면적은 항상 계산 (HX_AREA 노드는 존재하지 않음)
                area = None
                if q_w is not None and u is not None and lmtd is not None:
                    area = q_w / (u * lmtd)
                
                suggested = "fixed_tube"
                hx_preview.append({
                    "name": name,
                    "category": "HeatExchanger",
                    "q_w": q_w,
                    "u_W_m2K": u,
                    "lmtd_K": lmtd,
                    "area_m2": area,
                    "suggested": suggested,
                })
            except Exception as e:
                hx_preview.append({"name": name, "category": "HeatExchanger", "error": f"failed to read: {str(e)}"})
    
    return pressure_preview, hx_preview


def preview_all_devices_auto(
    Application,
    block_info: Dict[str, str],
    current_unit_set: Optional[str],
):
    """단위 세트를 자동으로 설정하여 통합 프리뷰를 생성합니다."""
    power_unit = None
    pressure_unit = None
    flow_unit = None
    heat_unit = None
    temperature_unit = None
    if current_unit_set:
        from aspen_data_extractor import get_unit_type_value
        power_unit = get_unit_type_value(Application, current_unit_set, 'POWER')
        pressure_unit = get_unit_type_value(Application, current_unit_set, 'PRESSURE')
        flow_unit = get_unit_type_value(Application, current_unit_set, 'VOLUME-FLOW')
        heat_unit = get_unit_type_value(Application, current_unit_set, 'HEAT')
        temperature_unit = get_unit_type_value(Application, current_unit_set, 'TEMPERATURE')
    return preview_all_devices_from_aspen(Application, block_info, power_unit, pressure_unit, flow_unit, heat_unit, temperature_unit)


def print_preview_all_results(
    pressure_preview: List[Dict],
    hx_preview: List[Dict],
    Application,
    power_unit: Optional[str],
    pressure_unit: Optional[str],
):
    """압력 장치와 열교환기 프리뷰를 모두 출력합니다."""
    print_preview_results(pressure_preview, Application, power_unit, pressure_unit)
    print_preview_hx_results(hx_preview)
