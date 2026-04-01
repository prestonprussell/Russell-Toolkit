from app.entra_graph import _canonical_integricom_license_from_sku_part
from app.processing import _normalize_integricom_branch
from app.processing import (
    INTEGRICOM_LICENSE_BP,
    INTEGRICOM_LICENSE_F3,
    INTEGRICOM_LICENSE_P1,
    INTEGRICOM_LICENSE_P2,
    INTEGRICOM_LICENSE_TEAMS_ESSENTIALS,
)


def test_canonical_integricom_license_from_sku_part_mappings() -> None:
    assert _canonical_integricom_license_from_sku_part("SPB") == INTEGRICOM_LICENSE_BP
    assert _canonical_integricom_license_from_sku_part("exchangestandard") == INTEGRICOM_LICENSE_P1
    assert _canonical_integricom_license_from_sku_part("EXCHANGEENTERPRISE") == INTEGRICOM_LICENSE_P2
    assert _canonical_integricom_license_from_sku_part("SPE_F3") == INTEGRICOM_LICENSE_F3
    assert _canonical_integricom_license_from_sku_part("TEAMS_ESSENTIALS_AAD") == INTEGRICOM_LICENSE_TEAMS_ESSENTIALS


def test_canonical_integricom_license_from_sku_part_unknown_returns_none() -> None:
    assert _canonical_integricom_license_from_sku_part("UNKNOWN_SKU_PART") is None
    assert _canonical_integricom_license_from_sku_part("") is None


def test_normalize_integricom_branch_uses_construction_department_override() -> None:
    assert _normalize_integricom_branch("Doraville", "Construction") == "Construction"
    assert _normalize_integricom_branch("Corporate", "Construction Operations") == "Construction"
    assert _normalize_integricom_branch("Doraville", "Operations") == "Doraville"
