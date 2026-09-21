"""
Generate a legacy-format report.json for Power BI Desktop.
Unlike PBIR (preview), this format is natively supported by all PBI Desktop versions.
All visual configs are embedded as stringified JSON inside a single report.json.
"""

import json, uuid, os

# ── Helpers ─────────────────────────────────────────────────────────────

def uid():
    return str(uuid.uuid4()).replace("-", "")[:16]

def lit(val):
    return {"expr": {"Literal": {"Value": str(val)}}}

def lit_str(s):
    return {"expr": {"Literal": {"Value": f"'{s}'"}}}

def solid_color(hex_c):
    return {"solid": {"color": lit(f"'{hex_c}'")}}

# ── Source aliases for prototypeQuery ───────────────────────────────────

SOURCES = {
    "Fact_Trips":   "f",
    "Dim_Date":     "d",
    "Dim_Driver":   "dr",
    "Dim_Rider":    "r",
    "Dim_Location": "l",
    "Dim_Promo":    "p",
    "_Measures":    "m",
}

def from_entry(table):
    return {"Name": SOURCES[table], "Entity": table, "Type": 0}

def select_measure(table, prop):
    s = SOURCES[table]
    return {
        "Measure": {
            "Expression": {"SourceRef": {"Source": s}},
            "Property": prop
        },
        "Name": f"{s}.{prop}"
    }

def select_column(table, prop):
    s = SOURCES[table]
    return {
        "Column": {
            "Expression": {"SourceRef": {"Source": s}},
            "Property": prop
        },
        "Name": f"{s}.{prop}"
    }

# ── Colors ──────────────────────────────────────────────────────────────

C_PRIMARY  = "#1A73E8"
C_GREEN    = "#0D9D58"
C_RED      = "#D93025"
C_AMBER    = "#F9AB00"
C_PURPLE   = "#7B61FF"
C_TEAL     = "#00897B"
C_DARK     = "#1B1B2F"
C_WHITE    = "#FFFFFF"
C_SUBTLE   = "#5F6368"
C_BORDER   = "#E8EAED"
C_BG       = "#F8F9FA"

# ── Visual builders ────────────────────────────────────────────────────

def make_title_props(title_text, color=C_DARK, size="12D"):
    return {
        "title": [{
            "properties": {
                "show": lit("true"),
                "titleWrap": lit("true"),
                "text": lit_str(title_text),
                "fontSize": lit(size),
                "fontFamily": lit("'Segoe UI Semibold'"),
                "fontColor": solid_color(color)
            }
        }],
        "background": [{
            "properties": {
                "show": lit("true"),
                "color": solid_color(C_WHITE),
                "transparency": lit("0D")
            }
        }],
        "visualHeader": [{
            "properties": {
                "show": lit("false")
            }
        }],
        "border": [{
            "properties": {
                "show": lit("true"),
                "color": solid_color(C_BORDER),
                "radius": lit("8D")
            }
        }],
        "dropShadow": [{
            "properties": {
                "show": lit("true"),
                "color": solid_color("#00000010"),
                "position": lit("'Outer'"),
                "preset": lit("'BottomRight'")
            }
        }]
    }


def build_card(name, x, y, w, h, z, title, table, measure, title_color=C_PRIMARY):
    config = {
        "name": name,
        "layouts": [{"id": 0, "position": {"x": x, "y": y, "width": w, "height": h}}],
        "singleVisual": {
            "visualType": "card",
            "projections": {
                "Values": [{"queryRef": f"{SOURCES[table]}.{measure}"}]
            },
            "prototypeQuery": {
                "Version": 2,
                "From": [from_entry(table)],
                "Select": [select_measure(table, measure)]
            },
            "objects": {
                "labels": [{"properties": {
                    "fontSize": lit("20D"),
                    "color": solid_color(C_DARK),
                    "fontFamily": lit("'Segoe UI Semibold'"),
                    "labelPrecision": lit("1D")
                }}],
                "categoryLabels": [{"properties": {
                    "show": lit("true"),
                    "fontSize": lit("9D"),
                    "color": solid_color(C_SUBTLE),
                    "fontFamily": lit("'Segoe UI'")
                }}]
            },
            "vcObjects": make_title_props(title, title_color, "10D")
        }
    }
    return {"x": x, "y": y, "z": z, "width": w, "height": h,
            "config": json.dumps(config, ensure_ascii=False)}


def build_chart(name, x, y, w, h, z, title, vis_type,
                cat_table=None, cat_col=None,
                val_entries=None, legend_table=None, legend_col=None,
                extra_objects=None):
    """
    val_entries: list of (table, measure_name)
    """
    froms = set()
    selects = []
    projections = {}

    if cat_table and cat_col:
        froms.add(cat_table)
        selects.append(select_column(cat_table, cat_col))
        projections["Category"] = [{"queryRef": f"{SOURCES[cat_table]}.{cat_col}"}]

    if legend_table and legend_col:
        froms.add(legend_table)
        selects.append(select_column(legend_table, legend_col))
        projections["Series"] = [{"queryRef": f"{SOURCES[legend_table]}.{legend_col}"}]

    y_key = "Y"
    if vis_type in ("card", "multiRowCard", "tableEx", "slicer"):
        y_key = "Values"

    if val_entries:
        projs = []
        for tbl, meas in val_entries:
            froms.add(tbl)
            selects.append(select_measure(tbl, meas))
            projs.append({"queryRef": f"{SOURCES[tbl]}.{meas}"})
        projections[y_key] = projs

    config = {
        "name": name,
        "layouts": [{"id": 0, "position": {"x": x, "y": y, "width": w, "height": h}}],
        "singleVisual": {
            "visualType": vis_type,
            "projections": projections,
            "prototypeQuery": {
                "Version": 2,
                "From": [from_entry(t) for t in sorted(froms)],
                "Select": selects
            },
            "vcObjects": make_title_props(title)
        }
    }
    if extra_objects:
        config["singleVisual"]["objects"] = extra_objects

    return {"x": x, "y": y, "z": z, "width": w, "height": h,
            "config": json.dumps(config, ensure_ascii=False)}


def build_slicer(name, x, y, w, h, z, title, table, column):
    froms = [from_entry(table)]
    selects = [select_column(table, column)]

    config = {
        "name": name,
        "layouts": [{"id": 0, "position": {"x": x, "y": y, "width": w, "height": h}}],
        "singleVisual": {
            "visualType": "slicer",
            "projections": {
                "Values": [{"queryRef": f"{SOURCES[table]}.{column}"}]
            },
            "prototypeQuery": {
                "Version": 2,
                "From": froms,
                "Select": selects
            },
            "objects": {
                "data": [{"properties": {
                    "mode": lit("'Dropdown'")
                }}],
                "selection": [{"properties": {
                    "selectAllCheckboxEnabled": lit("true"),
                    "singleSelect": lit("false")
                }}],
                "header": [{"properties": {
                    "show": lit("true"),
                    "fontColor": solid_color(C_DARK),
                    "fontSize": lit("10D"),
                    "fontFamily": lit("'Segoe UI Semibold'"),
                    "background": solid_color(C_WHITE)
                }}],
                "items": [{"properties": {
                    "fontColor": solid_color(C_DARK),
                    "fontSize": lit("9D"),
                    "fontFamily": lit("'Segoe UI'"),
                    "background": solid_color(C_WHITE)
                }}]
            },
            "vcObjects": make_title_props(title, C_DARK, "10D")
        }
    }
    return {"x": x, "y": y, "z": z, "width": w, "height": h,
            "config": json.dumps(config, ensure_ascii=False)}


def build_table(name, x, y, w, h, z, title, entries):
    """entries: list of (table, name, is_measure)"""
    froms = set()
    selects = []
    projs = []

    for tbl, prop, is_meas in entries:
        froms.add(tbl)
        if is_meas:
            selects.append(select_measure(tbl, prop))
        else:
            selects.append(select_column(tbl, prop))
        projs.append({"queryRef": f"{SOURCES[tbl]}.{prop}"})

    config = {
        "name": name,
        "layouts": [{"id": 0, "position": {"x": x, "y": y, "width": w, "height": h}}],
        "singleVisual": {
            "visualType": "tableEx",
            "projections": {"Values": projs},
            "prototypeQuery": {
                "Version": 2,
                "From": [from_entry(t) for t in sorted(froms)],
                "Select": selects
            },
            "objects": {
                "columnHeaders": [{"properties": {
                    "fontColor": solid_color(C_WHITE),
                    "backColor": solid_color(C_DARK),
                    "fontSize": lit("9D"),
                    "fontFamily": lit("'Segoe UI Semibold'")
                }}],
                "values": [{"properties": {
                    "fontSize": lit("9D"),
                    "fontFamily": lit("'Segoe UI'"),
                    "backColor": solid_color(C_WHITE)
                }}],
                "total": [{"properties": {
                    "fontSize": lit("9D"),
                    "fontFamily": lit("'Segoe UI Semibold'"),
                    "totals": lit("true")
                }}],
                "grid": [{"properties": {
                    "gridVertical": lit("true"),
                    "gridVerticalColor": solid_color("#E0E0E0"),
                    "rowPadding": lit("3D")
                }}]
            },
            "vcObjects": make_title_props(title)
        }
    }
    return {"x": x, "y": y, "z": z, "width": w, "height": h,
            "config": json.dumps(config, ensure_ascii=False)}


def build_header_shape(name, x, y, w, h, z, bg_color=C_DARK):
    """A colored rectangle used as a page header background."""
    config = {
        "name": name,
        "layouts": [{"id": 0, "position": {"x": x, "y": y, "width": w, "height": h}}],
        "singleVisual": {
            "visualType": "shape",
            "objects": {
                "general": [{"properties": {
                    "keepLayerOrder": lit("true")
                }}],
                "line": [{"properties": {
                    "show": lit("false")
                }}],
                "fill": [{"properties": {
                    "fillColor": solid_color(bg_color),
                    "transparency": lit("0D")
                }}],
                "rotation": [{"properties": {
                    "angle": lit("0D")
                }}]
            },
            "vcObjects": {
                "background": [{"properties": {"show": lit("false")}}],
                "border": [{"properties": {"show": lit("false")}}],
                "visualHeader": [{"properties": {"show": lit("false")}}]
            }
        }
    }
    return {"x": x, "y": y, "z": z, "width": w, "height": h,
            "config": json.dumps(config, ensure_ascii=False)}


def build_text_label(name, x, y, w, h, z, text, font_size="20", color=C_WHITE):
    """Text label for page headers."""
    para = json.dumps([{
        "textRuns": [{
            "value": text,
            "textStyle": {
                "fontFamily": "Segoe UI Semibold",
                "fontSize": f"{font_size}pt",
                "color": color
            }
        }]
    }])
    config = {
        "name": name,
        "layouts": [{"id": 0, "position": {"x": x, "y": y, "width": w, "height": h}}],
        "singleVisual": {
            "visualType": "textbox",
            "objects": {
                "general": [{"properties": {
                    "paragraphs": lit(para)
                }}]
            },
            "vcObjects": {
                "background": [{"properties": {"show": lit("false")}}],
                "border": [{"properties": {"show": lit("false")}}],
                "visualHeader": [{"properties": {"show": lit("false")}}]
            }
        }
    }
    return {"x": x, "y": y, "z": z, "width": w, "height": h,
            "config": json.dumps(config, ensure_ascii=False)}


# ── Build section (page) ───────────────────────────────────────────────

def make_section(name, display_name, ordinal, visuals, bg_color=C_BG):
    section_config = {
        "layouts": [{"id": 0, "position": {"x": 0, "y": 0, "width": 1280, "height": 720}}],
        "visibility": 0
    }
    return {
        "name": name,
        "displayName": display_name,
        "ordinal": ordinal,
        "displayOption": 1,
        "width": 1280,
        "height": 720,
        "config": json.dumps(section_config, ensure_ascii=False),
        "filters": "[]",
        "visualContainers": visuals
    }


# ═════════════════════════════════════════════════════════════════════════
# PAGE 1: Executive Overview
# ═════════════════════════════════════════════════════════════════════════

print("Building Page 1: Executive Overview...")

p1 = []
p1.append(build_header_shape("p1bg", 0, 0, 1280, 52, 0, C_DARK))
p1.append(build_text_label("p1title", 15, 6, 600, 40, 1, "Ride-Share Analytics  |  Executive Overview", "14", C_WHITE))

# KPI row
p1.append(build_card("c1trips",   20,  62, 192, 95, 10, "Total Trips",       "_Measures", "Total Trips",       C_PRIMARY))
p1.append(build_card("c1rev",    222,  62, 192, 95, 11, "Total Revenue",     "_Measures", "Total Revenue",     C_GREEN))
p1.append(build_card("c1cancel", 424,  62, 192, 95, 12, "Cancellation Rate", "_Measures", "Cancellation Rate %", C_RED))
p1.append(build_card("c1fare",   626,  62, 192, 95, 13, "Avg Fare",          "_Measures", "Avg Fare",          C_AMBER))
p1.append(build_card("c1yoy",    828,  62, 192, 95, 14, "YoY Growth",        "_Measures", "YoY Growth %",      C_PURPLE))

# City slicer
p1.append(build_slicer("s1city", 1035, 62, 225, 95, 15, "City", "Dim_Location", "City"))

# Revenue Trend (line)
p1.append(build_chart("ch1trend", 20, 168, 615, 255, 20,
    "Monthly Revenue Trend", "lineChart",
    cat_table="Dim_Date", cat_col="Date",
    val_entries=[("_Measures", "Total Revenue")]))

# Trips by City (bar)
p1.append(build_chart("ch1city", 650, 168, 610, 255, 21,
    "Trips by City", "clusteredBarChart",
    cat_table="Dim_Location", cat_col="City",
    val_entries=[("_Measures", "Total Trips")]))

# Status donut
p1.append(build_chart("ch1status", 20, 434, 295, 268, 22,
    "Completed vs Cancelled", "donutChart",
    cat_table="Fact_Trips", cat_col="Status",
    val_entries=[("_Measures", "Total Trips")]))

# Revenue by Vehicle
p1.append(build_chart("ch1veh", 330, 434, 295, 268, 23,
    "Revenue by Vehicle Type", "clusteredColumnChart",
    cat_table="Dim_Driver", cat_col="VehicleType",
    val_entries=[("_Measures", "Total Revenue")]))

# Payment donut
p1.append(build_chart("ch1pay", 640, 434, 295, 268, 24,
    "Payment Methods", "donutChart",
    cat_table="Fact_Trips", cat_col="PaymentMethod",
    val_entries=[("_Measures", "Total Revenue")]))

# Revenue by Quarter
p1.append(build_chart("ch1qtr", 950, 434, 310, 268, 25,
    "Revenue by Quarter", "clusteredColumnChart",
    cat_table="Dim_Date", cat_col="Quarter",
    val_entries=[("_Measures", "Total Revenue")]))

section1 = make_section("ExecOverview", "Executive Overview", 0, p1)


# ═════════════════════════════════════════════════════════════════════════
# PAGE 2: Demand & Cancellation Deep-Dive
# ═════════════════════════════════════════════════════════════════════════

print("Building Page 2: Demand & Cancellation...")

p2 = []
p2.append(build_header_shape("p2bg", 0, 0, 1280, 52, 0, "#B71C1C"))
p2.append(build_text_label("p2title", 15, 6, 600, 40, 1, "Demand & Cancellation Deep-Dive", "14", C_WHITE))

# Slicers
p2.append(build_slicer("s2city", 20, 62, 190, 90, 5, "City", "Dim_Location", "City"))
p2.append(build_slicer("s2year", 220, 62, 140, 90, 6, "Year", "Dim_Date", "Year"))
p2.append(build_slicer("s2qtr",  370, 62, 140, 90, 7, "Quarter", "Dim_Date", "Quarter"))

# Cards
p2.append(build_card("c2rate",   525, 62, 175, 90, 8,  "Cancellation Rate", "_Measures", "Cancellation Rate %", C_RED))
p2.append(build_card("c2cnt",    710, 62, 175, 90, 9,  "Cancelled Trips",   "_Measures", "Cancelled Trips",     C_RED))
p2.append(build_card("c2surge",  895, 62, 175, 90, 10, "Avg Surge at Cancel", "_Measures", "Avg Surge at Cancellation", C_AMBER))
p2.append(build_card("c2wait",  1080, 62, 180, 90, 11, "Avg Wait Time",     "_Measures", "Avg Wait Time",       C_TEAL))

# Trip Volume by Hour
p2.append(build_chart("ch2hour", 20, 162, 400, 255, 20,
    "Trip Volume by Hour of Day", "clusteredColumnChart",
    cat_table="Fact_Trips", cat_col="HourOfDay",
    val_entries=[("_Measures", "Total Trips")]))

# Cancel Rate by City
p2.append(build_chart("ch2city", 435, 162, 400, 255, 21,
    "Cancellation Rate by City", "clusteredBarChart",
    cat_table="Dim_Location", cat_col="City",
    val_entries=[("_Measures", "Cancellation Rate %")]))

# Cancel Rate by Surge
p2.append(build_chart("ch2surge", 850, 162, 410, 255, 22,
    "Cancel Rate vs Surge Multiplier", "lineChart",
    cat_table="Fact_Trips", cat_col="SurgeMultiplier",
    val_entries=[("_Measures", "Cancellation Rate %")]))

# Cancel Reasons
p2.append(build_chart("ch2reason", 20, 427, 400, 275, 23,
    "Cancellation Reasons", "clusteredBarChart",
    cat_table="Fact_Trips", cat_col="CancellationReason",
    val_entries=[("_Measures", "Cancelled Trips")]))

# Cancel by DayOfWeek
p2.append(build_chart("ch2dow", 435, 427, 400, 275, 24,
    "Cancellation Rate by Day of Week", "clusteredColumnChart",
    cat_table="Dim_Date", cat_col="DayOfWeek",
    val_entries=[("_Measures", "Cancellation Rate %")]))

# Cancel Trend Monthly
p2.append(build_chart("ch2trend", 850, 427, 410, 275, 25,
    "Monthly Cancellation Trend", "lineChart",
    cat_table="Dim_Date", cat_col="Date",
    val_entries=[("_Measures", "Cancellation Rate %"), ("_Measures", "Cancelled Trips")]))

section2 = make_section("DemandCancel", "Demand & Cancellation", 1, p2)


# ═════════════════════════════════════════════════════════════════════════
# PAGE 3: Driver Performance
# ═════════════════════════════════════════════════════════════════════════

print("Building Page 3: Driver Performance...")

p3 = []
p3.append(build_header_shape("p3bg", 0, 0, 1280, 52, 0, "#1565C0"))
p3.append(build_text_label("p3title", 15, 6, 500, 40, 1, "Driver Performance", "14", C_WHITE))

# Slicers
p3.append(build_slicer("s3city", 20, 62, 195, 90, 5, "City", "Dim_Driver", "HomeCity"))
p3.append(build_slicer("s3veh",  225, 62, 195, 90, 6, "Vehicle Type", "Dim_Driver", "VehicleType"))

# Cards
p3.append(build_card("c3rpd",  435, 62, 195, 90, 7, "Revenue / Driver", "_Measures", "Revenue per Driver", C_PRIMARY))
p3.append(build_card("c3fare", 640, 62, 195, 90, 8, "Avg Fare",         "_Measures", "Avg Fare",           C_GREEN))
p3.append(build_card("c3rpk",  845, 62, 195, 90, 9, "Revenue / KM",     "_Measures", "Revenue per KM",     C_AMBER))
p3.append(build_card("c3tips",1050, 62, 210, 90,10, "Total Tips",        "_Measures", "Total Tips",         C_TEAL))

# Top Drivers table
p3.append(build_table("tbl3top", 20, 162, 620, 290, 20,
    "Top Drivers by Revenue", [
        ("Dim_Driver", "DriverName",  False),
        ("Dim_Driver", "HomeCity",    False),
        ("Dim_Driver", "VehicleType", False),
        ("Dim_Driver", "Rating",      False),
        ("_Measures",  "Total Revenue", True),
        ("_Measures",  "Total Trips",   True),
        ("_Measures",  "Driver Revenue Rank", True),
    ]))

# Rating vs Revenue scatter
p3.append(build_chart("ch3scatter", 655, 162, 605, 290, 21,
    "Driver Rating vs Revenue", "scatterChart",
    cat_table="Dim_Driver", cat_col="Rating",
    val_entries=[("_Measures", "Total Revenue")]))

# Revenue/Driver by City
p3.append(build_chart("ch3city", 20, 462, 400, 240, 22,
    "Revenue per Driver by City", "clusteredBarChart",
    cat_table="Dim_Driver", cat_col="HomeCity",
    val_entries=[("_Measures", "Revenue per Driver")]))

# Revenue by Vehicle
p3.append(build_chart("ch3veh", 435, 462, 400, 240, 23,
    "Revenue & Trips by Vehicle Type", "clusteredColumnChart",
    cat_table="Dim_Driver", cat_col="VehicleType",
    val_entries=[("_Measures", "Total Revenue"), ("_Measures", "Total Trips")]))

# Tips by Vehicle
p3.append(build_chart("ch3tips", 850, 462, 410, 240, 24,
    "Tips by Vehicle Type", "clusteredColumnChart",
    cat_table="Dim_Driver", cat_col="VehicleType",
    val_entries=[("_Measures", "Total Tips")]))

section3 = make_section("DriverPerf", "Driver Performance", 2, p3)


# ═════════════════════════════════════════════════════════════════════════
# PAGE 4: Promotions Impact
# ═════════════════════════════════════════════════════════════════════════

print("Building Page 4: Promotions Impact...")

p4 = []
p4.append(build_header_shape("p4bg", 0, 0, 1280, 52, 0, "#E65100"))
p4.append(build_text_label("p4title", 15, 6, 500, 40, 1, "Promotions Impact", "14", C_WHITE))

# Slicers
p4.append(build_slicer("s4camp", 20, 62, 245, 90, 5, "Campaign", "Dim_Promo", "CampaignName"))
p4.append(build_slicer("s4tier", 275, 62, 195, 90, 6, "Loyalty Tier", "Dim_Rider", "LoyaltyTier"))

# Cards
p4.append(build_card("c4ptrips", 485, 62, 185, 90, 7, "Promo Trips",   "_Measures", "Promo Trip Count", C_PRIMARY))
p4.append(build_card("c4prev",   680, 62, 185, 90, 8, "Promo Revenue", "_Measures", "Promo Revenue",    C_GREEN))
p4.append(build_card("c4trev",   875, 62, 185, 90, 9, "Total Revenue", "_Measures", "Total Revenue",    C_PURPLE))
p4.append(build_card("c4tips",  1070, 62, 190, 90,10, "Total Tips",    "_Measures", "Total Tips",       C_TEAL))

# Revenue by Campaign
p4.append(build_chart("ch4camp", 20, 162, 615, 260, 20,
    "Revenue by Campaign", "clusteredBarChart",
    cat_table="Dim_Promo", cat_col="CampaignName",
    val_entries=[("_Measures", "Total Revenue")]))

# Trips by Campaign
p4.append(build_chart("ch4trips", 650, 162, 610, 260, 21,
    "Trips by Campaign", "clusteredColumnChart",
    cat_table="Dim_Promo", cat_col="CampaignName",
    val_entries=[("_Measures", "Total Trips"), ("_Measures", "Promo Trip Count")]))

# Revenue by Loyalty Tier donut
p4.append(build_chart("ch4tier", 20, 432, 295, 270, 22,
    "Revenue by Loyalty Tier", "donutChart",
    cat_table="Dim_Rider", cat_col="LoyaltyTier",
    val_entries=[("_Measures", "Total Revenue")]))

# Trips by Loyalty Tier
p4.append(build_chart("ch4tiertrips", 330, 432, 295, 270, 23,
    "Trips by Loyalty Tier", "clusteredColumnChart",
    cat_table="Dim_Rider", cat_col="LoyaltyTier",
    val_entries=[("_Measures", "Total Trips")]))

# Promo detail table
p4.append(build_table("tbl4promo", 640, 432, 620, 270, 24,
    "Promotion Details", [
        ("Dim_Promo", "PromoCode",    False),
        ("Dim_Promo", "CampaignName", False),
        ("Dim_Promo", "DiscountPct",  False),
        ("_Measures", "Total Trips",  True),
        ("_Measures", "Total Revenue", True),
        ("_Measures", "Avg Fare",     True),
    ]))

section4 = make_section("PromoImpact", "Promotions Impact", 3, p4)


# ═════════════════════════════════════════════════════════════════════════
# ASSEMBLE report.json
# ═════════════════════════════════════════════════════════════════════════

report_config = {
    "version": "5.56",
    "themeCollection": {
        "baseTheme": {
            "name": "CY24SU06",
            "reportVersionAtImport": "5.56",
            "type": "SharedResources"
        }
    },
    "activeSectionIndex": 0,
    "defaultDrillFilterOtherVisuals": True,
    "slowDataSourceSettings": {
        "isCrossHighlightingDisabled": False,
        "isSlicerSelectionsButtonEnabled": False,
        "isFilterSelectionsButtonEnabled": False,
        "isFieldWellButtonEnabled": False,
        "isApplyAllButtonEnabled": False
    },
    "linguisticSchemaSyncVersion": 2,
    "settings": {
        "useStylableVisualContainerHeader": True,
        "exportDataMode": 1,
        "useNewFilterPaneExperience": True,
        "allowChangeFilterTypes": True,
        "useEnhancedTooltips": True,
        "isPersistentUserStateDisabled": True
    }
}

report = {
    "id": uid(),
    "resourcePackages": [],
    "sections": [section1, section2, section3, section4],
    "config": json.dumps(report_config, ensure_ascii=False),
    "layoutOptimization": 0
}

output = os.path.join(
    "/Users/bsinga1/Desktop/Powerbi/RideShareAnalytics",
    "RideShareAnalytics.Report",
    "report.json"
)
with open(output, "w", encoding="utf-8") as f:
    json.dump(report, f, indent=2, ensure_ascii=False)

total_visuals = sum(len(s["visualContainers"]) for s in report["sections"])
print(f"\n{'='*60}")
print(f"✅  report.json written → {output}")
print(f"    Pages:   {len(report['sections'])}")
print(f"    Visuals: {total_visuals}")
print(f"    Size:    {os.path.getsize(output) / 1024:.0f} KB")
print(f"{'='*60}")
