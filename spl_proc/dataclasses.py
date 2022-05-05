
from dataclasses import dataclass, field
from decimal import Decimal


@dataclass
class SupplierPricelistItem:
    supp_code: str
    supp_item_code: str
    supp_price: Decimal
    supp_uom: str
    supp_sell_uom: str
    supp_eoq: str
    supp_conv_factor: Decimal
