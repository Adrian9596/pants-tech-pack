from __future__ import annotations
from typing import List, Optional
from pydantic import BaseModel, Field
from datetime import datetime


class ColorwayInput(BaseModel):
    name: str = "CW1"
    primary_color: str = ""
    primary_pantone: str = ""
    secondary_color: str = ""
    secondary_pantone: str = ""
    accent_color: str = ""
    accent_pantone: str = ""


class BOMItem(BaseModel):
    item_type: str = "Fabric"  # Fabric / Trim / Label / Packaging
    description: str = ""
    supplier: str = "TBD"
    color: str = ""
    pantone: str = ""
    content: str = ""
    width_size: str = ""
    placement: str = ""
    qty: float = 1.0
    unit: str = "yds"
    unit_cost: Optional[float] = None
    lead_time: int = 0
    notes: str = ""


class CostingInput(BaseModel):
    cut_cost: Optional[float] = None
    make_cost: Optional[float] = None
    trim_finishing: Optional[float] = None
    washing_treatment: Optional[float] = None
    embroidery_print: Optional[float] = None
    hang_tag_ticketing: Optional[float] = None
    freight_factory: Optional[float] = None
    testing_compliance: Optional[float] = None
    material_waste_pct: float = 0.05
    overhead_pct: float = 0.10
    margin_pct: float = 0.12
    export_docs: Optional[float] = None
    local_transport: Optional[float] = None
    target_wholesale: Optional[float] = None
    target_retail: Optional[float] = None


class TechPackRequest(BaseModel):
    brand_name: str = "[Brand Name]"
    style_name: str = "[Style Name]"
    style_number: str = ""
    season: str = ""
    designer: str = "TBD"
    garment_type: str = "knit_top"
    category: str = ""
    fabric_content: str = ""
    size_range: str = "XS-XL"
    target_fob: str = "TBD"
    factory_name: str = "TBD"

    # which sheets to include; "all" means all 7
    sheets: List[str] = Field(default_factory=lambda: ["all"])

    # size chart / grading block key, e.g. "male_1", "female_2", "block_1"
    size_chart: str = ""

    colorways: List[ColorwayInput] = Field(default_factory=list)
    bom_items: List[BOMItem] = Field(default_factory=list)
    costing: CostingInput = Field(default_factory=CostingInput)

    def resolved_style_number(self) -> str:
        if self.style_number:
            return self.style_number
        year = datetime.now().year
        return f"STY-{year}-001"

    def should_generate(self, sheet: str) -> bool:
        if "all" in self.sheets:
            return True
        return sheet in self.sheets
