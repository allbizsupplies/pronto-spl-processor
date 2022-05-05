

import csv
from typing import List

from spl_proc.dataclasses import SupplierPricelistItem


def export_supplier_pricelist(filepath: str, spl_items: List[SupplierPricelistItem]):
    def spl_item_to_row(spl_item: SupplierPricelistItem):
        return {
            "supplier_code": spl_item.supp_code,
            "supp_item_code": spl_item.supp_item_code,
            "supp_uom": spl_item.supp_uom,
            "supp_eoq": spl_item.supp_eoq,
            "supp_conv_factor": spl_item.supp_conv_factor,
            "supp_price_1": spl_item.supp_price,
        }

    with open(filepath, "w", encoding="iso8859-14") as file:
        writer = csv.DictWriter(file, get_spl_fieldnames(), dialect="excel")
        writer.writerows(
            [spl_item_to_row(spl_item) for spl_item in spl_items])


def get_spl_fieldnames():
    return [
        "supplier_code",
        "catalogue_part_no",
        "supp_item_code",
        "desc_line_1",
        "desc_line_2",
        "supp_uom",
        "supp_sell_uom",
        "supp_eoq",
        "supp_conv_factor",
        "supp_price_1",
        "supp_price_2",
        "supp_price_3",
        "supp_price_4",
        "gst",
        "barcode",
        "carton_size",
        "flc_page_no",
        "rrp",
        "major_category",
        "minor_category",
        "was_manufacturer_code",
        "item_code",
        "office_choice_code",
        "quantity_1_pronto_0",
        "quantity_2_pronto_1",
        "quantity_3_pronto_2",
        "quantity_4_pronto_3",
        "price_1_pronto_0",
        "price_2_pronto_1",
        "price_3_pronto_2",
        "price_4_pronto_3",
        "supp_priority",
        "supp_inner_uom",
        "supp_inner_barcode",
        "supp_inner_conversion_factor",
        "supp_outer_uom",
        "supp_outer_barcode",
        "supp_outer_conversion_factor",
        "unit_measurements",
        "unit_weight",
        "cartons_per_pallet",
        "eoq",
        "sell_uom",
        "is_consumable",
        "is_branded",
        "is_green",
        "created_on",
        "status",
        "product_class",
        "product_group",
        "legacy_item_code",
    ]
