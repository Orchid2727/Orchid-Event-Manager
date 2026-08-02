from __future__ import annotations

import re
from pathlib import Path
from typing import Mapping

import pandas as pd

from modules.blank_garment_rules import BLANK_DECORATION_LABEL, apply_blank_garment_defaults

TEMP_TRUSPEC_STYLE = "24/7"
TRUSPEC_VENDOR = "Tru-Spec"
VF_STYLE_OVERRIDES = {"2574", "SP3A"}
TRUSPEC_STYLE_OVERRIDES = {"4266", "1122", "24/7"}
KNOWN_TRUSPEC_BLANK_STYLES = {"1104", "4266", "1122", "24/7"}


# Preferred distributor routing supplied by Orchid Uniforms.
# Red Kap is a hard exception and always routes to VF. When a brand is sold by
# both SanMar and S&S Activewear, SanMar is the preferred purchase vendor.
SANMAR_BRAND_ALIASES = {
    "A4": ("A4",),
    "Allmade": ("Allmade",),
    "BELLA+CANVAS": ("BELLA+CANVAS", "Bella + Canvas", "Bella and Canvas", "Bella Canvas"),
    "Brooks Brothers": ("Brooks Brothers",),
    "Bulwark": ("Bulwark",),
    "Carhartt": ("Carhartt",),
    "Champion": ("Champion",),
    "Comfort Colors": ("Comfort Colors",),
    "CornerStone": ("CornerStone", "Corner Stone"),
    "Cotopaxi": ("Cotopaxi",),
    "District": ("District",),
    "Eddie Bauer": ("Eddie Bauer",),
    "Flexfit": ("Flexfit", "Flex Fit"),
    "Gildan": ("Gildan",),
    "Jerzees": ("Jerzees",),
    "Mercer+Mettle": ("Mercer+Mettle", "Mercer + Mettle", "Mercer and Mettle", "Mercer Mettle"),
    "MiiR": ("MiiR",),
    "New Era": ("New Era",),
    "Next Level Apparel": ("Next Level Apparel", "Next Level"),
    "Nike": ("Nike",),
    "OGIO": ("OGIO",),
    "Outdoor Research": ("Outdoor Research",),
    "Port & Co": ("Port & Co", "Port and Company", "Port & Company", "Port Company"),
    "Port Authority": ("Port Authority",),
    "Rabbit Skins": ("Rabbit Skins",),
    "Richardson": ("Richardson",),
    "Russell Outdoors": ("Russell Outdoors",),
    "Spacecraft": ("Spacecraft",),
    "Sport-Tek": ("Sport-Tek", "Sport Tek", "SportTek"),
    "Stanley/Stella": ("Stanley/Stella", "Stanley Stella", "Stanley and Stella"),
    "tentree": ("tentree", "ten tree"),
    "The North Face": ("The North Face", "North Face"),
    "Tommy Bahama": ("Tommy Bahama",),
    "TravisMathew": ("TravisMathew", "Travis Mathew"),
    "Volunteer Knitwear": ("Volunteer Knitwear",),
    "Wink": ("Wink",),
}

# Brands shown on the S&S Activewear catalog that are not in Orchid's SanMar
# preference list. These route to S&S only after the VF and SanMar rules.
SS_BRAND_ALIASES = {
    "47 Brand": ("47 Brand", "'47 Brand"),
    "Adams Headwear": ("Adams Headwear", "Adams"),
    "adidas": ("adidas",),
    "Alleson Athletic": ("Alleson Athletic", "Alleson"),
    "AllPro": ("AllPro", "All Pro"),
    "Alpine Fleece": ("Alpine Fleece",),
    "American Apparel": ("American Apparel",),
    "ANETIK": ("ANETIK",),
    "Artisan Collection by Reprime": ("Artisan Collection by Reprime", "Reprime"),
    "Atlantis Headwear": ("Atlantis Headwear", "Atlantis"),
    "Augusta Sportswear": ("Augusta Sportswear", "Augusta"),
    "Authentic Pigment": ("Authentic Pigment",),
    "Badger": ("Badger",),
    "BAGedge": ("BAGedge", "Bag Edge"),
    "Bayside": ("Bayside",),
    "Berne Apparel": ("Berne Apparel", "Berne"),
    "Big Accessories": ("Big Accessories",),
    "Boxercraft": ("Boxercraft",),
    "C2 Sport": ("C2 Sport",),
    "Carmel Towel Company": ("Carmel Towel Company", "Carmel Towel"),
    "Chef Designs": ("Chef Designs",),
    "Classic Caps": ("Classic Caps",),
    "Code Five": ("Code Five",),
    "ColorTone": ("ColorTone", "Color Tone"),
    "Columbia": ("Columbia",),
    "ComfortWash": ("ComfortWash", "Comfort Wash"),
    "CORE365": ("CORE365", "Core 365"),
    "Devon & Jones": ("Devon & Jones", "Devon and Jones", "Devon Jones"),
    "Dickies": ("Dickies",),
    "Doggie Skins": ("Doggie Skins",),
    "DRI DUCK": ("DRI DUCK", "Dri Duck"),
    "econscious": ("econscious",),
    "Fairway & Greene": ("Fairway & Greene", "Fairway and Greene"),
    "Hanes": ("Hanes",),
    "Harriton": ("Harriton",),
    "Holloway": ("Holloway",),
    "HUK": ("HUK",),
    "Imperial": ("Imperial",),
    "Independent Trading Co.": ("Independent Trading Co", "Independent Trading Company"),
    "Infinity Her": ("Infinity Her",),
    "J. America": ("J. America", "J America"),
    "JAANUU": ("JAANUU",),
    "Kastlfel": ("Kastlfel",),
    "Kati": ("Kati",),
    "Kishigo": ("Kishigo",),
    "Lane Seven": ("Lane Seven",),
    "LAT": ("LAT", "LAT Apparel"),
    "Legacy": ("Legacy",),
    "Liberty Bags": ("Liberty Bags",),
    "LOCALE": ("LOCALE",),
    "Los Angeles Apparel": ("Los Angeles Apparel",),
    "M&O": ("M&O", "M and O"),
    "Marmot": ("Marmot",),
    "MERET": ("MERET",),
    "MV Sport": ("MV Sport",),
    "Nautica": ("Nautica",),
    "Ninja Transfers": ("Ninja Transfers",),
    "Nomadix": ("Nomadix",),
    "North End": ("North End",),
    "OAD": ("OAD",),
    "Oakley": ("Oakley",),
    "Onna by Premier": ("Onna by Premier", "Onna Premier"),
    "Paragon": ("Paragon",),
    "Pukka": ("Pukka",),
    "Puma Golf": ("Puma Golf", "Puma"),
    "Q-Tees": ("Q-Tees", "Q Tees"),
    "RealTree": ("RealTree", "Real Tree"),
    "Recover": ("Recover",),
    "Russell Athletic": ("Russell Athletic",),
    "Shaka Wear": ("Shaka Wear",),
    "SoftShirts": ("SoftShirts", "Soft Shirts"),
    "Sportsman": ("Sportsman",),
    "Spyder": ("Spyder",),
    "Sublivie": ("Sublivie",),
    "Swannies": ("Swannies",),
    "TASC Performance": ("TASC Performance", "TASC"),
    "Team 365": ("Team 365",),
    "The Game": ("The Game",),
    "Threadfast Apparel": ("Threadfast Apparel", "Threadfast"),
    "TriDri": ("TriDri", "Tri Dri"),
    "Tultex": ("Tultex",),
    "UltraClub": ("UltraClub", "Ultra Club"),
    "Under Armour": ("Under Armour",),
    "Unionwear": ("Unionwear", "Union Wear"),
    "US Blanks": ("US Blanks",),
    "Valucap": ("Valucap", "Value Cap"),
    "Vineyard Vines": ("Vineyard Vines",),
    "Weatherproof": ("Weatherproof",),
    "YP Classics": ("YP Classics",),
    "Zero Restriction": ("Zero Restriction",),
}

RED_KAP_ALIASES = ("Red Kap", "RedKap")

# Brands Orchid normally purchases directly from the named vendor instead of
# routing through a broadline distributor. These rules fill blank Product
# Master vendors while keeping every dropdown editable.
DIRECT_VENDOR_BRAND_ALIASES = {
    "Wrangler": ("Wrangler",),
    "Tru-Spec": ("Tru-Spec", "Tru Spec"),
    "Burnside": ("Burnside",),
    "Cutter & Buck": ("Cutter & Buck", "Cutter and Buck", "Cutter Buck"),
    "Outdoor Cap": ("Outdoor Cap",),
}


def _brand_normalize(value: object) -> str:
    text = _clean(value).casefold()
    text = text.replace("&", " and ").replace("+", " and ").replace("/", " ")
    text = re.sub(r"[®™'’`.]", "", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _contains_brand(text: str, aliases: tuple[str, ...]) -> bool:
    normalized = f" {_brand_normalize(text)} "
    for alias in sorted(aliases, key=lambda item: len(_brand_normalize(item)), reverse=True):
        term = _brand_normalize(alias)
        if not term:
            continue
        if term == "a4":
            # Avoid interpreting an arbitrary embedded style code as the A4 brand.
            stripped = normalized.strip()
            if stripped == "a4" or stripped.startswith("a4 ") or " a4 apparel " in normalized:
                return True
            continue
        if f" {term} " in normalized:
            return True
    return False


def preferred_vendor_for_brand_text(value: object) -> tuple[str, str]:
    """Return (brand, preferred vendor) from a Shopify/product description.

    Priority is VF Red Kap exception, direct-brand vendors, SanMar, then
    S&S Activewear.
    """
    text = _clean(value)
    if not text:
        return "", ""
    if _contains_brand(text, RED_KAP_ALIASES):
        return "Red Kap", "VF"
    for brand, aliases in DIRECT_VENDOR_BRAND_ALIASES.items():
        if _contains_brand(text, aliases):
            return brand, brand
    for brand, aliases in SANMAR_BRAND_ALIASES.items():
        if _contains_brand(text, aliases):
            return brand, "SanMar"
    for brand, aliases in SS_BRAND_ALIASES.items():
        if _contains_brand(text, aliases):
            return brand, "S&S Activewear"
    return "", ""


def _row_brand_routing(row: Mapping[str, object]) -> tuple[str, str]:
    fields = (
        "Brand", "Product Brand", "Product Name", "Description",
        "Original Line Item", "Original Shopify Line", "Lineitem name",
    )
    combined = " | ".join(_clean(row.get(field, "")) for field in fields if _clean(row.get(field, "")))
    return preferred_vendor_for_brand_text(combined)


def _clean(value) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def _style(value) -> str:
    return re.sub(r"\s+", "", _clean(value)).upper()


def is_temporary_truspec_item(
    style: object = "",
    product: object = "",
    original_line: object = "",
    parser_source: object = "",
) -> bool:
    """Identify the manually entered Tru-Spec pant records needing the 24/7 workaround.

    Style 2574 is explicitly excluded because it is a VF product.
    """
    style_key = _style(style)
    if style_key in VF_STYLE_OVERRIDES:
        return False
    # A free-text custom item can carry the correct color-specific Tru-Spec
    # style number. Preserve that explicit number instead of forcing 24/7.
    if _clean(parser_source).casefold() == "smart custom item" and style_key and style_key != TEMP_TRUSPEC_STYLE:
        return False

    text = " ".join([_clean(product), _clean(original_line)]).casefold()
    phrases = (
        "tru-spec",
        "24/7 original tactical pant",
        "original tactical pant rip-stop",
        "original tactical pant rip stop",
        "pro flex pant",
        "pro-flex pant",
        "agility pant",
    )
    return any(phrase in text for phrase in phrases)


def apply_parsed_product_overrides(record: Mapping[str, object]) -> dict:
    """Return a parsed Shopify record with temporary business overrides applied."""
    updated = dict(record)
    style = _clean(updated.get("Style Number", ""))
    product = _clean(updated.get("Product Name", ""))
    original = _clean(updated.get("Original Line Item", ""))
    parser_source = _clean(updated.get("Parser Source", ""))

    if is_temporary_truspec_item(style, product, original, parser_source):
        updated["Style Number"] = TEMP_TRUSPEC_STYLE
        # Preserve the complete pant name and color so different pants never merge.
        reason = _clean(updated.get("Needs Review", ""))
        reason_parts = [
            part.strip()
            for part in reason.split(";")
            if part.strip() and "style number could not be determined" not in part.casefold()
        ]
        updated["Needs Review"] = "; ".join(reason_parts)

    return updated


def apply_routing_overrides(frame: pd.DataFrame) -> pd.DataFrame:
    """Apply business routing rules without overwriting reliable custom-item data."""
    if frame.empty:
        return frame

    result = apply_blank_garment_defaults(frame.copy())
    for column in [
        "Style Number", "Product Name", "Original Line Item", "Vendor",
        "Decoration Type", "Decoration Color", "Parser Source", "Detected Vendor",
        "Product Category", "Requires Size", "Requires Color", "Requires Decoration",
        "Master Match",
    ]:
        if column not in result.columns:
            result[column] = ""

    # Smart custom items may contain the supplier/vendor even when Shopify's
    # catalog fields are blank. Use that hint only when Product Master did not
    # already provide a vendor.
    missing_vendor = result["Vendor"].map(_clean).eq("")
    detected_vendor = result["Detected Vendor"].map(_clean)
    use_detected = missing_vendor & detected_vendor.ne("")
    if use_detected.any():
        result.loc[use_detected, "Vendor"] = detected_vendor[use_detected]

    # Apply Orchid's preferred distributor rules only when Product Master did
    # not already provide an authoritative vendor. An exact Product Master
    # style/product match with a nonblank saved vendor must survive every later
    # brand inference step. This prevents products such as HJ51 and FRP07SWD
    # from being changed from Berne Apparel to S&S Activewear during review
    # creation merely because their descriptions contain a broadline brand.
    brand_routes = result.apply(_row_brand_routing, axis=1)
    if len(brand_routes):
        result["Detected Brand"] = [item[0] for item in brand_routes]
        preferred_vendors = pd.Series([item[1] for item in brand_routes], index=result.index)
        reliable_master_match = result["Master Match"].map(_clean).str.casefold().isin({
            "style",
            "style - exact style/color alias",
            "style - single-row style inference",
            "product alias",
            "contained product alias",
        })
        saved_master_vendor = reliable_master_match & result["Vendor"].map(_clean).ne("")
        brand_mask = preferred_vendors.map(_clean).ne("") & ~saved_master_vendor
        if brand_mask.any():
            result.loc[brand_mask, "Vendor"] = preferred_vendors[brand_mask]
            result.loc[brand_mask, "Vendor Routing Source"] = "Preferred brand vendor"
        if saved_master_vendor.any():
            result.loc[saved_master_vendor, "Vendor Routing Source"] = "Product Master vendor"

    # Style 1104 (Ascent) is a known Tru-Spec pant. Older builds repeatedly
    # treated it as a new product because the title does not contain the brand.
    known_truspec_blank = result["Style Number"].map(_style).isin(KNOWN_TRUSPEC_BLANK_STYLES)
    if known_truspec_blank.any():
        result.loc[known_truspec_blank, "Vendor"] = TRUSPEC_VENDOR
        result.loc[known_truspec_blank, "Product Category"] = "Pants / Jeans / Shorts"
        result.loc[known_truspec_blank, "Requires Size"] = "Yes"
        result.loc[known_truspec_blank, "Requires Color"] = "Yes"
        result.loc[known_truspec_blank, "Requires Decoration"] = "No"
        result.loc[known_truspec_blank, "Decoration Type"] = BLANK_DECORATION_LABEL
        result.loc[known_truspec_blank, "Decoration Color"] = ""
        result.loc[known_truspec_blank, "Vendor Routing Source"] = "Known Orchid style rule"

    detected_temp_mask = result.apply(
        lambda row: is_temporary_truspec_item(
            row.get("Style Number", ""),
            row.get("Product Name", ""),
            row.get("Original Line Item", ""),
            row.get("Parser Source", ""),
        ),
        axis=1,
    )
    explicit_custom_style = (
        result["Parser Source"].map(_clean).str.casefold().eq("smart custom item")
        & result["Style Number"].map(_style).ne("")
        & result["Style Number"].map(_style).ne(TEMP_TRUSPEC_STYLE)
    )
    # A vendor value by itself must never rewrite a genuine style to the
    # temporary 24/7 placeholder. Only descriptions that explicitly match the
    # temporary Tru-Spec pant rule are converted.
    reliable_master_match = result["Master Match"].map(_clean).str.casefold().isin({
        "style", "style - exact style/color alias", "style - single-row style inference",
        "product alias", "contained product alias",
    })
    temp_mask = detected_temp_mask & ~explicit_custom_style & ~known_truspec_blank & ~reliable_master_match
    if temp_mask.any():
        result.loc[temp_mask, "Style Number"] = TEMP_TRUSPEC_STYLE
        result.loc[temp_mask, "Vendor"] = TRUSPEC_VENDOR
        result.loc[temp_mask, "Decoration Type"] = BLANK_DECORATION_LABEL
        result.loc[temp_mask, "Decoration Color"] = ""

    # Explicit Tru-Spec custom pants keep their entered style but still route as
    # blank garments unless the review workbook is deliberately changed.
    explicit_truspec = explicit_custom_style & result["Vendor"].map(_clean).str.casefold().eq(TRUSPEC_VENDOR.casefold())
    if explicit_truspec.any():
        result.loc[explicit_truspec, "Decoration Type"] = result.loc[explicit_truspec, "Decoration Type"].map(_clean).replace("", BLANK_DECORATION_LABEL)
        result.loc[explicit_truspec, "Decoration Color"] = ""

    truspec_mask = result["Style Number"].map(_style).isin(TRUSPEC_STYLE_OVERRIDES)
    if truspec_mask.any():
        result.loc[truspec_mask, "Vendor"] = TRUSPEC_VENDOR

    vf_mask = result["Style Number"].map(_style).isin(VF_STYLE_OVERRIDES)
    if vf_mask.any():
        result.loc[vf_mask, "Vendor"] = "VF"

    # DRI DUCK is a brand purchased through S&S Activewear, never a standalone vendor.
    brand_text = (
        result.get("Product Name", "").map(_clean) + " " +
        result.get("Original Line Item", "").map(_clean) + " " +
        result.get("Detected Vendor", "").map(_clean) + " " +
        result.get("Vendor", "").map(_clean)
    ).str.casefold()
    dri_duck_mask = brand_text.str.contains(r"\bdri[ -]?duck\b", regex=True, na=False)
    if dri_duck_mask.any():
        result.loc[dri_duck_mask, "Vendor"] = "S&S Activewear"

    return apply_blank_garment_defaults(result)



def enforce_permanent_vendor_overrides(product_master_path) -> int:
    """Persist Orchid's hard vendor rules into the active Product Master.

    This prevents re-imports or regenerated reviews from reviving an older vendor
    assignment for known Tru-Spec families or known VF/Red Kap styles. Only exact
    known styles or clearly identified brand families are changed.
    """
    path = Path(product_master_path).expanduser()
    if not path.exists():
        return 0
    try:
        frame = pd.read_csv(path, dtype=str).fillna("")
    except Exception:
        return 0
    if frame.empty or "Vendor" not in frame.columns:
        return 0
    for column in ("Style Number", "Product Name", "Product Aliases"):
        if column not in frame.columns:
            frame[column] = ""
    style = frame["Style Number"].map(_style)
    text = (frame["Product Name"].map(_clean) + " " + frame["Product Aliases"].map(_clean)).str.casefold()
    vendor_key = frame["Vendor"].map(_clean).str.casefold()
    tru_mask = style.isin(TRUSPEC_STYLE_OVERRIDES | KNOWN_TRUSPEC_BLANK_STYLES) | text.str.contains(
        r"(?:tru[ -]?spec|24/7 original tactical pant|original tactical pant rip[ -]?stop)",
        regex=True, na=False,
    )
    vf_mask = style.isin(VF_STYLE_OVERRIDES) | text.str.contains(
        r"\bred[ -]?kap\b", regex=True, na=False,
    )
    changed_tru = tru_mask & vendor_key.ne(TRUSPEC_VENDOR.casefold())
    changed_vf = vf_mask & vendor_key.ne("vf")
    changed = changed_tru | changed_vf
    count = int(changed.sum())
    if count:
        frame.loc[changed_tru, "Vendor"] = TRUSPEC_VENDOR
        frame.loc[changed_vf, "Vendor"] = "VF"
        frame.to_csv(path, index=False)
    return count
