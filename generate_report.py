"""
Generate the PBIR report definition for 4-page production-quality dashboard.

Pages:
  1. Executive Overview
  2. Demand & Cancellation Deep-Dive
  3. Driver Performance
  4. Promotions Impact

Color palette (professional dark/blue theme):
  Primary:   #1A73E8  (blue)
  Success:   #0D9D58  (green)
  Danger:    #D93025  (red)
  Warning:   #F9AB00  (amber)
  Dark BG:   #1B1B2F  (header / title bar)
  Light BG:  #F8F9FA  (page background)
"""

import json
import os

BASE = "/Users/bsinga1/Desktop/Powerbi/RideShareAnalytics/RideShareAnalytics.Report/definition"

# ── Color constants ─────────────────────────────────────────────────────

C_PRIMARY    = "#1A73E8"
C_SUCCESS    = "#0D9D58"
C_DANGER     = "#D93025"
C_WARNING    = "#F9AB00"
C_PURPLE     = "#7B61FF"
C_TEAL       = "#00897B"
C_DARK       = "#1B1B2F"
C_LIGHT_BG   = "#F8F9FA"
C_WHITE      = "#FFFFFF"
C_TEXT       = "#202124"
C_SUBTLE     = "#5F6368"

# Page dimensions (16:9)
PAGE_W = 1280
PAGE_H = 720

# ── Helpers ─────────────────────────────────────────────────────────────

def lit(val):
    """Wrap a value as a Power BI literal expression."""
    return {"expr": {"Literal": {"Value": str(val)}}}

def lit_str(val):
    return {"expr": {"Literal": {"Value": f"'{val}'"}}}

def color_expr(hex_color):
    return {"solid": {"color": lit(f"'{hex_color}'")}}

def measure_field(table, prop):
    return {
        "Measure": {
            "Expression": {"SourceRef": {"Entity": table}},
            "Property": prop
        }
    }

def column_field(table, prop):
    return {
        "Column": {
            "Expression": {"SourceRef": {"Entity": table}},
            "Property": prop
        }
    }

def make_projection(field, query_ref, active=True):
    return {"field": field, "queryRef": query_ref, "active": active}

def card_visual(name, x, y, w, h, z, tab, title, table, measure, title_color=C_PRIMARY):
    return {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visual/1.0.0/schema.json",
        "name": name,
        "position": {"x": x, "y": y, "z": z, "width": w, "height": h, "tabOrder": tab},
        "visual": {
            "visualType": "card",
            "query": {
                "queryState": {
                    "Values": {
                        "projections": [
                            make_projection(measure_field(table, measure), f"{table}.{measure}")
                        ]
                    }
                }
            },
            "objects": {
                "labels": [{"properties": {
                    "fontSize": lit("18D"),
                    "color": color_expr(C_DARK),
                    "fontFamily": lit("'Segoe UI Semibold'")
                }}],
                "categoryLabels": [{"properties": {
                    "show": lit("true"),
                    "fontSize": lit("10D"),
                    "color": color_expr(C_SUBTLE)
                }}]
            },
            "visualContainerObjects": {
                "title": [{"properties": {
                    "show": lit("true"),
                    "text": lit_str(title),
                    "fontSize": lit("10D"),
                    "fontFamily": lit("'Segoe UI Semibold'"),
                    "fontColor": color_expr(title_color)
                }}],
                "background": [{"properties": {
                    "show": lit("true"),
                    "color": color_expr(C_WHITE),
                    "transparency": lit("0D")
                }}],
                "border": [{"properties": {
                    "show": lit("true"),
                    "color": color_expr("#E0E0E0"),
                    "radius": lit("8D")
                }}],
                "padding": [{"properties": {
                    "top": lit("8D"),
                    "bottom": lit("8D"),
                    "left": lit("12D"),
                    "right": lit("12D")
                }}]
            }
        }
    }


def chart_visual(name, x, y, w, h, z, tab, title, vis_type, category_entries, value_entries, objects=None):
    """Generic chart visual builder."""
    query_state = {}
    if category_entries:
        query_state["Category"] = {
            "projections": [make_projection(e["field"], e["ref"]) for e in category_entries]
        }
    if value_entries:
        query_state["Y" if vis_type != "donutChart" else "Y"] = {
            "projections": [make_projection(e["field"], e["ref"]) for e in value_entries]
        }

    visual = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visual/1.0.0/schema.json",
        "name": name,
        "position": {"x": x, "y": y, "z": z, "width": w, "height": h, "tabOrder": tab},
        "visual": {
            "visualType": vis_type,
            "query": {"queryState": query_state},
            "visualContainerObjects": {
                "title": [{"properties": {
                    "show": lit("true"),
                    "text": lit_str(title),
                    "fontSize": lit("12D"),
                    "fontFamily": lit("'Segoe UI Semibold'"),
                    "fontColor": color_expr(C_DARK)
                }}],
                "background": [{"properties": {
                    "show": lit("true"),
                    "color": color_expr(C_WHITE),
                    "transparency": lit("0D")
                }}],
                "border": [{"properties": {
                    "show": lit("true"),
                    "color": color_expr("#E0E0E0"),
                    "radius": lit("8D")
                }}],
                "padding": [{"properties": {
                    "top": lit("6D"),
                    "bottom": lit("6D"),
                    "left": lit("10D"),
                    "right": lit("10D")
                }}]
            }
        }
    }
    if objects:
        visual["visual"]["objects"] = objects
    return visual


def slicer_visual(name, x, y, w, h, z, tab, title, table, column):
    return {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visual/1.0.0/schema.json",
        "name": name,
        "position": {"x": x, "y": y, "z": z, "width": w, "height": h, "tabOrder": tab},
        "visual": {
            "visualType": "slicer",
            "query": {
                "queryState": {
                    "Values": {
                        "projections": [
                            make_projection(column_field(table, column), f"{table}.{column}")
                        ]
                    }
                }
            },
            "objects": {
                "data": [{"properties": {
                    "mode": lit("'Dropdown'")
                }}]
            },
            "visualContainerObjects": {
                "title": [{"properties": {
                    "show": lit("true"),
                    "text": lit_str(title),
                    "fontSize": lit("10D"),
                    "fontFamily": lit("'Segoe UI Semibold'"),
                    "fontColor": color_expr(C_DARK)
                }}],
                "background": [{"properties": {
                    "show": lit("true"),
                    "color": color_expr(C_WHITE),
                    "transparency": lit("0D")
                }}],
                "border": [{"properties": {
                    "show": lit("true"),
                    "color": color_expr("#E0E0E0"),
                    "radius": lit("8D")
                }}]
            }
        }
    }


def table_visual(name, x, y, w, h, z, tab, title, value_entries):
    return {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visual/1.0.0/schema.json",
        "name": name,
        "position": {"x": x, "y": y, "z": z, "width": w, "height": h, "tabOrder": tab},
        "visual": {
            "visualType": "tableEx",
            "query": {
                "queryState": {
                    "Values": {
                        "projections": [make_projection(e["field"], e["ref"]) for e in value_entries]
                    }
                }
            },
            "objects": {
                "columnHeaders": [{"properties": {
                    "fontColor": color_expr(C_WHITE),
                    "backColor": color_expr(C_DARK),
                    "fontSize": lit("9D"),
                    "fontFamily": lit("'Segoe UI Semibold'")
                }}],
                "values": [{"properties": {
                    "fontSize": lit("9D"),
                    "fontFamily": lit("'Segoe UI'")
                }}],
                "grid": [{"properties": {
                    "gridVertical": lit("true"),
                    "gridVerticalColor": color_expr("#E0E0E0"),
                    "rowPadding": lit("4D")
                }}]
            },
            "visualContainerObjects": {
                "title": [{"properties": {
                    "show": lit("true"),
                    "text": lit_str(title),
                    "fontSize": lit("12D"),
                    "fontFamily": lit("'Segoe UI Semibold'"),
                    "fontColor": color_expr(C_DARK)
                }}],
                "background": [{"properties": {
                    "show": lit("true"),
                    "color": color_expr(C_WHITE),
                    "transparency": lit("0D")
                }}],
                "border": [{"properties": {
                    "show": lit("true"),
                    "color": color_expr("#E0E0E0"),
                    "radius": lit("8D")
                }}]
            }
        }
    }


def text_box_visual(name, x, y, w, h, z, tab, text, font_size=20, color=C_WHITE, bg_color=C_DARK):
    """Title banner / text label."""
    return {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visual/1.0.0/schema.json",
        "name": name,
        "position": {"x": x, "y": y, "z": z, "width": w, "height": h, "tabOrder": tab},
        "visual": {
            "visualType": "textbox",
            "objects": {
                "general": [{"properties": {
                    "paragraphs": lit(json.dumps([{
                        "textRuns": [{"value": text, "textStyle": {"fontFamily": "Segoe UI Semibold", "fontSize": f"{font_size}px", "color": color}}]
                    }]))
                }}]
            },
            "visualContainerObjects": {
                "background": [{"properties": {
                    "show": lit("true"),
                    "color": color_expr(bg_color),
                    "transparency": lit("0D")
                }}],
                "border": [{"properties": {
                    "show": lit("false")
                }}]
            }
        }
    }


def write_visual(page_name, visual_name, visual_data):
    folder = os.path.join(BASE, "pages", page_name, "visuals", visual_name)
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, "visual.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(visual_data, f, indent=2, ensure_ascii=False)


def write_page(page_name, display_name, visuals_dict):
    page_dir = os.path.join(BASE, "pages", page_name)
    os.makedirs(page_dir, exist_ok=True)

    page_json = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/1.0.0/schema.json",
        "name": page_name,
        "displayName": display_name,
        "displayOption": 1,
        "height": PAGE_H,
        "width": PAGE_W,
        "visualContainers": []
    }
    with open(os.path.join(page_dir, "page.json"), "w", encoding="utf-8") as f:
        json.dump(page_json, f, indent=2, ensure_ascii=False)

    for vis_name, vis_data in visuals_dict.items():
        write_visual(page_name, vis_name, vis_data)

    print(f"  ✅ {display_name}: {len(visuals_dict)} visuals")


# ═════════════════════════════════════════════════════════════════════════
# PAGE 1: Executive Overview
# ═════════════════════════════════════════════════════════════════════════

print("Building report pages...")

p1_visuals = {}

# Title banner
p1_visuals["TitleBanner"] = text_box_visual(
    "TitleBanner", 0, 0, PAGE_W, 50, 0, 0,
    "🚗  Ride-Share Analytics  |  Executive Overview", 18, C_WHITE, C_DARK
)

# KPI Cards row
p1_visuals["CardTotalTrips"] = card_visual(
    "CardTotalTrips", 20, 60, 195, 100, 1, 1,
    "Total Trips", "_Measures", "Total Trips", C_PRIMARY
)
p1_visuals["CardRevenue"] = card_visual(
    "CardRevenue", 225, 60, 195, 100, 2, 2,
    "Total Revenue", "_Measures", "Total Revenue", C_SUCCESS
)
p1_visuals["CardCancelRate"] = card_visual(
    "CardCancelRate", 430, 60, 195, 100, 3, 3,
    "Cancellation Rate", "_Measures", "Cancellation Rate %", C_DANGER
)
p1_visuals["CardAvgFare"] = card_visual(
    "CardAvgFare", 635, 60, 195, 100, 4, 4,
    "Avg Fare", "_Measures", "Avg Fare", C_WARNING
)
p1_visuals["CardYoY"] = card_visual(
    "CardYoY", 840, 60, 195, 100, 5, 5,
    "YoY Growth", "_Measures", "YoY Growth %", C_PURPLE
)

# Slicers
p1_visuals["SlicerCity"] = slicer_visual(
    "SlicerCity", 1050, 60, 210, 100, 6, 6,
    "City", "Dim_Location", "City"
)

# Revenue trend line chart
p1_visuals["RevenueTrend"] = chart_visual(
    "RevenueTrend", 20, 170, 620, 260, 10, 10,
    "Revenue Trend (Monthly)", "lineChart",
    [{"field": column_field("Dim_Date", "Date"), "ref": "Dim_Date.Date"}],
    [
        {"field": measure_field("_Measures", "Total Revenue"), "ref": "_Measures.Total Revenue"},
    ]
)

# Trips by City bar chart
p1_visuals["TripsByCity"] = chart_visual(
    "TripsByCity", 660, 170, 600, 260, 11, 11,
    "Trips by City", "clusteredBarChart",
    [{"field": column_field("Dim_Location", "City"), "ref": "Dim_Location.City"}],
    [{"field": measure_field("_Measures", "Total Trips"), "ref": "_Measures.Total Trips"}]
)

# Trip Status donut
p1_visuals["TripStatus"] = chart_visual(
    "TripStatus", 20, 440, 300, 260, 12, 12,
    "Trip Status Breakdown", "donutChart",
    [{"field": column_field("Fact_Trips", "Status"), "ref": "Fact_Trips.Status"}],
    [{"field": measure_field("_Measures", "Total Trips"), "ref": "_Measures.Total Trips"}]
)

# Revenue by Vehicle Type
p1_visuals["RevenueByVehicle"] = chart_visual(
    "RevenueByVehicle", 340, 440, 300, 260, 13, 13,
    "Revenue by Vehicle Type", "clusteredColumnChart",
    [{"field": column_field("Dim_Driver", "VehicleType"), "ref": "Dim_Driver.VehicleType"}],
    [{"field": measure_field("_Measures", "Total Revenue"), "ref": "_Measures.Total Revenue"}]
)

# Payment Method donut
p1_visuals["PaymentMethod"] = chart_visual(
    "PaymentMethod", 660, 440, 300, 260, 14, 14,
    "Payment Methods", "donutChart",
    [{"field": column_field("Fact_Trips", "PaymentMethod"), "ref": "Fact_Trips.PaymentMethod"}],
    [{"field": measure_field("_Measures", "Total Revenue"), "ref": "_Measures.Total Revenue"}]
)

# Revenue by Quarter
p1_visuals["RevenueByQuarter"] = chart_visual(
    "RevenueByQuarter", 980, 440, 280, 260, 15, 15,
    "Revenue by Quarter", "clusteredColumnChart",
    [{"field": column_field("Dim_Date", "Quarter"), "ref": "Dim_Date.Quarter"}],
    [{"field": measure_field("_Measures", "Total Revenue"), "ref": "_Measures.Total Revenue"}]
)

write_page("ExecutiveOverview", "Executive Overview", p1_visuals)


# ═════════════════════════════════════════════════════════════════════════
# PAGE 2: Demand & Cancellation Deep-Dive
# ═════════════════════════════════════════════════════════════════════════

p2_visuals = {}

p2_visuals["TitleBanner2"] = text_box_visual(
    "TitleBanner2", 0, 0, PAGE_W, 50, 0, 0,
    "📊  Demand & Cancellation Deep-Dive", 18, C_WHITE, "#B71C1C"
)

# Slicer row
p2_visuals["SlicerCity2"] = slicer_visual(
    "SlicerCity2", 20, 60, 200, 90, 1, 1,
    "City", "Dim_Location", "City"
)
p2_visuals["SlicerYear"] = slicer_visual(
    "SlicerYear", 230, 60, 150, 90, 2, 2,
    "Year", "Dim_Date", "Year"
)
p2_visuals["SlicerQuarter"] = slicer_visual(
    "SlicerQuarter", 390, 60, 150, 90, 3, 3,
    "Quarter", "Dim_Date", "Quarter"
)

# Cards
p2_visuals["CardCancelRate2"] = card_visual(
    "CardCancelRate2", 560, 60, 170, 90, 4, 4,
    "Cancellation Rate", "_Measures", "Cancellation Rate %", C_DANGER
)
p2_visuals["CardCancelledTrips"] = card_visual(
    "CardCancelledTrips", 740, 60, 170, 90, 5, 5,
    "Cancelled Trips", "_Measures", "Cancelled Trips", C_DANGER
)
p2_visuals["CardAvgSurge"] = card_visual(
    "CardAvgSurge", 920, 60, 170, 90, 6, 6,
    "Avg Surge at Cancel", "_Measures", "Avg Surge at Cancellation", C_WARNING
)
p2_visuals["CardAvgWait"] = card_visual(
    "CardAvgWait", 1100, 60, 160, 90, 7, 7,
    "Avg Wait Time", "_Measures", "Avg Wait Time", C_TEAL
)

# Trip volume by Hour (bar chart)
p2_visuals["TripsByHour"] = chart_visual(
    "TripsByHour", 20, 160, 400, 260, 10, 10,
    "Trip Volume by Hour of Day", "clusteredColumnChart",
    [{"field": column_field("Fact_Trips", "HourOfDay"), "ref": "Fact_Trips.HourOfDay"}],
    [{"field": measure_field("_Measures", "Total Trips"), "ref": "_Measures.Total Trips"}]
)

# Cancellation Rate by City (bar chart)
p2_visuals["CancelByCity"] = chart_visual(
    "CancelByCity", 440, 160, 400, 260, 11, 11,
    "Cancellation Rate by City", "clusteredBarChart",
    [{"field": column_field("Dim_Location", "City"), "ref": "Dim_Location.City"}],
    [{"field": measure_field("_Measures", "Cancellation Rate %"), "ref": "_Measures.Cancellation Rate %"}]
)

# Cancellation Rate by Surge bucket (line)
p2_visuals["CancelBySurge"] = chart_visual(
    "CancelBySurge", 860, 160, 400, 260, 12, 12,
    "Cancellation Rate by Surge Level", "lineChart",
    [{"field": column_field("Fact_Trips", "SurgeMultiplier"), "ref": "Fact_Trips.SurgeMultiplier"}],
    [{"field": measure_field("_Measures", "Cancellation Rate %"), "ref": "_Measures.Cancellation Rate %"}]
)

# Cancellation Reasons breakdown (bar)
p2_visuals["CancelReasons"] = chart_visual(
    "CancelReasons", 20, 430, 400, 270, 13, 13,
    "Cancellation Reasons", "clusteredBarChart",
    [{"field": column_field("Fact_Trips", "CancellationReason"), "ref": "Fact_Trips.CancellationReason"}],
    [{"field": measure_field("_Measures", "Cancelled Trips"), "ref": "_Measures.Cancelled Trips"}]
)

# Cancellation by Day of Week
p2_visuals["CancelByDay"] = chart_visual(
    "CancelByDay", 440, 430, 400, 270, 14, 14,
    "Cancellation Rate by Day of Week", "clusteredColumnChart",
    [{"field": column_field("Dim_Date", "DayOfWeek"), "ref": "Dim_Date.DayOfWeek"}],
    [{"field": measure_field("_Measures", "Cancellation Rate %"), "ref": "_Measures.Cancellation Rate %"}]
)

# Trips and Cancellation Rate over time (combo)
p2_visuals["TrendCancel"] = chart_visual(
    "TrendCancel", 860, 430, 400, 270, 15, 15,
    "Cancellation Trend (Monthly)", "lineChart",
    [{"field": column_field("Dim_Date", "Date"), "ref": "Dim_Date.Date"}],
    [
        {"field": measure_field("_Measures", "Cancellation Rate %"), "ref": "_Measures.Cancellation Rate %"},
        {"field": measure_field("_Measures", "Cancelled Trips"), "ref": "_Measures.Cancelled Trips"},
    ]
)

write_page("DemandCancellation", "Demand & Cancellation", p2_visuals)


# ═════════════════════════════════════════════════════════════════════════
# PAGE 3: Driver Performance
# ═════════════════════════════════════════════════════════════════════════

p3_visuals = {}

p3_visuals["TitleBanner3"] = text_box_visual(
    "TitleBanner3", 0, 0, PAGE_W, 50, 0, 0,
    "🏆  Driver Performance", 18, C_WHITE, "#1565C0"
)

# Slicer
p3_visuals["SlicerCity3"] = slicer_visual(
    "SlicerCity3", 20, 60, 200, 90, 1, 1,
    "City", "Dim_Driver", "HomeCity"
)
p3_visuals["SlicerVehicle"] = slicer_visual(
    "SlicerVehicle", 230, 60, 200, 90, 2, 2,
    "Vehicle Type", "Dim_Driver", "VehicleType"
)

# KPI Cards
p3_visuals["CardRevPerDriver"] = card_visual(
    "CardRevPerDriver", 450, 60, 200, 90, 3, 3,
    "Revenue per Driver", "_Measures", "Revenue per Driver", C_PRIMARY
)
p3_visuals["CardTotalDrivers"] = card_visual(
    "CardTotalDrivers", 660, 60, 200, 90, 4, 4,
    "Active Drivers", "_Measures", "Total Trips", C_TEAL
)
p3_visuals["CardAvgFare3"] = card_visual(
    "CardAvgFare3", 870, 60, 200, 90, 5, 5,
    "Avg Fare", "_Measures", "Avg Fare", C_SUCCESS
)
p3_visuals["CardRevPerKM"] = card_visual(
    "CardRevPerKM", 1080, 60, 180, 90, 6, 6,
    "Revenue per KM", "_Measures", "Revenue per KM", C_WARNING
)

# Top Drivers table
p3_visuals["TopDrivers"] = table_visual(
    "TopDrivers", 20, 160, 620, 300, 10, 10,
    "Top Drivers by Revenue",
    [
        {"field": measure_field("_Measures", "Driver Revenue Rank"), "ref": "_Measures.Driver Revenue Rank"},
        {"field": column_field("Dim_Driver", "DriverName"), "ref": "Dim_Driver.DriverName"},
        {"field": column_field("Dim_Driver", "HomeCity"), "ref": "Dim_Driver.HomeCity"},
        {"field": column_field("Dim_Driver", "VehicleType"), "ref": "Dim_Driver.VehicleType"},
        {"field": column_field("Dim_Driver", "Rating"), "ref": "Dim_Driver.Rating"},
        {"field": measure_field("_Measures", "Total Revenue"), "ref": "_Measures.Total Revenue"},
        {"field": measure_field("_Measures", "Total Trips"), "ref": "_Measures.Total Trips"},
    ]
)

# Rating vs Revenue scatter
p3_visuals["RatingVsRevenue"] = chart_visual(
    "RatingVsRevenue", 660, 160, 600, 300, 11, 11,
    "Driver Rating vs Revenue", "scatterChart",
    [{"field": column_field("Dim_Driver", "Rating"), "ref": "Dim_Driver.Rating"}],
    [{"field": measure_field("_Measures", "Total Revenue"), "ref": "_Measures.Total Revenue"}]
)

# Revenue by City bar
p3_visuals["RevByCity"] = chart_visual(
    "RevByCity", 20, 470, 400, 230, 12, 12,
    "Revenue per Driver by City", "clusteredBarChart",
    [{"field": column_field("Dim_Driver", "HomeCity"), "ref": "Dim_Driver.HomeCity"}],
    [{"field": measure_field("_Measures", "Revenue per Driver"), "ref": "_Measures.Revenue per Driver"}]
)

# Revenue by Vehicle Type
p3_visuals["RevByVehicle3"] = chart_visual(
    "RevByVehicle3", 440, 470, 400, 230, 13, 13,
    "Revenue by Vehicle Type", "clusteredColumnChart",
    [{"field": column_field("Dim_Driver", "VehicleType"), "ref": "Dim_Driver.VehicleType"}],
    [
        {"field": measure_field("_Measures", "Total Revenue"), "ref": "_Measures.Total Revenue"},
        {"field": measure_field("_Measures", "Total Trips"), "ref": "_Measures.Total Trips"},
    ]
)

# Trips per Driver distribution (by rating bucket)
p3_visuals["TripsPerDriver"] = chart_visual(
    "TripsPerDriver", 860, 470, 400, 230, 14, 14,
    "Trips by Driver Rating", "clusteredColumnChart",
    [{"field": column_field("Dim_Driver", "Rating"), "ref": "Dim_Driver.Rating"}],
    [{"field": measure_field("_Measures", "Total Trips"), "ref": "_Measures.Total Trips"}]
)

write_page("DriverPerformance", "Driver Performance", p3_visuals)


# ═════════════════════════════════════════════════════════════════════════
# PAGE 4: Promotions Impact
# ═════════════════════════════════════════════════════════════════════════

p4_visuals = {}

p4_visuals["TitleBanner4"] = text_box_visual(
    "TitleBanner4", 0, 0, PAGE_W, 50, 0, 0,
    "🎯  Promotions Impact", 18, C_WHITE, "#E65100"
)

# Slicer
p4_visuals["SlicerCampaign"] = slicer_visual(
    "SlicerCampaign", 20, 60, 250, 90, 1, 1,
    "Campaign", "Dim_Promo", "CampaignName"
)
p4_visuals["SlicerTier"] = slicer_visual(
    "SlicerTier", 280, 60, 200, 90, 2, 2,
    "Loyalty Tier", "Dim_Rider", "LoyaltyTier"
)

# Cards
p4_visuals["CardPromoTrips"] = card_visual(
    "CardPromoTrips", 500, 60, 185, 90, 3, 3,
    "Promo Trips", "_Measures", "Promo Trip Count", C_PRIMARY
)
p4_visuals["CardPromoRevenue"] = card_visual(
    "CardPromoRevenue", 695, 60, 185, 90, 4, 4,
    "Promo Revenue", "_Measures", "Promo Revenue", C_SUCCESS
)
p4_visuals["CardTotalRevenue4"] = card_visual(
    "CardTotalRevenue4", 890, 60, 185, 90, 5, 5,
    "Total Revenue", "_Measures", "Total Revenue", C_PURPLE
)
p4_visuals["CardTotalTips4"] = card_visual(
    "CardTotalTips4", 1085, 60, 175, 90, 6, 6,
    "Total Tips", "_Measures", "Total Tips", C_TEAL
)

# Revenue by Campaign bar
p4_visuals["RevByCampaign"] = chart_visual(
    "RevByCampaign", 20, 160, 620, 260, 10, 10,
    "Revenue by Campaign", "clusteredBarChart",
    [{"field": column_field("Dim_Promo", "CampaignName"), "ref": "Dim_Promo.CampaignName"}],
    [{"field": measure_field("_Measures", "Total Revenue"), "ref": "_Measures.Total Revenue"}]
)

# Trips by Campaign
p4_visuals["TripsByCampaign"] = chart_visual(
    "TripsByCampaign", 660, 160, 600, 260, 11, 11,
    "Trips by Campaign", "clusteredColumnChart",
    [{"field": column_field("Dim_Promo", "CampaignName"), "ref": "Dim_Promo.CampaignName"}],
    [
        {"field": measure_field("_Measures", "Total Trips"), "ref": "_Measures.Total Trips"},
        {"field": measure_field("_Measures", "Promo Trip Count"), "ref": "_Measures.Promo Trip Count"},
    ]
)

# Revenue by Loyalty Tier donut
p4_visuals["RevByTier"] = chart_visual(
    "RevByTier", 20, 430, 300, 270, 12, 12,
    "Revenue by Loyalty Tier", "donutChart",
    [{"field": column_field("Dim_Rider", "LoyaltyTier"), "ref": "Dim_Rider.LoyaltyTier"}],
    [{"field": measure_field("_Measures", "Total Revenue"), "ref": "_Measures.Total Revenue"}]
)

# Trips by Loyalty Tier
p4_visuals["TripsByTier"] = chart_visual(
    "TripsByTier", 340, 430, 300, 270, 13, 13,
    "Trips by Loyalty Tier", "clusteredColumnChart",
    [{"field": column_field("Dim_Rider", "LoyaltyTier"), "ref": "Dim_Rider.LoyaltyTier"}],
    [{"field": measure_field("_Measures", "Total Trips"), "ref": "_Measures.Total Trips"}]
)

# Promo detail table
p4_visuals["PromoTable"] = table_visual(
    "PromoTable", 660, 430, 600, 270, 14, 14,
    "Promotion Details",
    [
        {"field": column_field("Dim_Promo", "PromoCode"), "ref": "Dim_Promo.PromoCode"},
        {"field": column_field("Dim_Promo", "CampaignName"), "ref": "Dim_Promo.CampaignName"},
        {"field": column_field("Dim_Promo", "DiscountPct"), "ref": "Dim_Promo.DiscountPct"},
        {"field": measure_field("_Measures", "Total Trips"), "ref": "_Measures.Total Trips"},
        {"field": measure_field("_Measures", "Total Revenue"), "ref": "_Measures.Total Revenue"},
        {"field": measure_field("_Measures", "Avg Fare"), "ref": "_Measures.Avg Fare"},
    ]
)

write_page("PromotionsImpact", "Promotions Impact", p4_visuals)


# ── pages.json ──────────────────────────────────────────────────────────

pages_json = {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pages/1.0.0/schema.json",
    "pageOrder": [
        "ExecutiveOverview",
        "DemandCancellation",
        "DriverPerformance",
        "PromotionsImpact"
    ]
}
with open(os.path.join(BASE, "pages", "pages.json"), "w", encoding="utf-8") as f:
    json.dump(pages_json, f, indent=2)

# ── report.json ─────────────────────────────────────────────────────────

report_json = {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/1.0.0/schema.json",
    "themeCollection": {
        "baseTheme": {
            "name": "CY24SU06",
            "reportVersionAtImport": "5.56",
            "type": "SharedResources"
        }
    },
    "layoutOptimization": 0,
    "resourcePackages": []
}
with open(os.path.join(BASE, "report.json"), "w", encoding="utf-8") as f:
    json.dump(report_json, f, indent=2)

print(f"\n{'='*60}")
print("✅  All report pages generated successfully!")
print(f"{'='*60}")
