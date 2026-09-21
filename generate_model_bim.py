"""
Generate the model.bim (TMSL) file for the RideShareAnalytics PBIP project.
Includes: 6 data tables + 1 measures table + 1 parameter, 5 relationships, 14 DAX measures.
"""

import json
import os

MODEL = {
    "compatibilityLevel": 1567,
    "model": {
        "culture": "en-US",
        "dataAccessOptions": {
            "legacyRedirects": True,
            "returnErrorValuesAsNull": True
        },
        "defaultPowerBIDataSourceVersion": "powerBI_V3",
        "sourceQueryCulture": "en-US",
        "tables": [],
        "relationships": [],
        "annotations": [
            {"name": "PBI_QueryOrder", "value": "[\"ExcelFilePath\",\"Fact_Trips\",\"Dim_Driver\",\"Dim_Rider\",\"Dim_Location\",\"Dim_Date\",\"Dim_Promo\"]"},
            {"name": "PBIDesktopVersion", "value": "2.138.1004.0"},
            {"name": "__PBI_TimeIntelligenceEnabled", "value": "1"}
        ]
    }
}

# ── Helper to generate lineage tags ────────────────────────────────────

tag_counter = [0]
def tag():
    tag_counter[0] += 1
    n = tag_counter[0]
    return f"{n:08x}-{n:04x}-{n:04x}-{n:04x}-{n:012x}"


# ── ExcelFilePath parameter ─────────────────────────────────────────────

MODEL["model"]["tables"].append({
    "name": "ExcelFilePath",
    "lineageTag": tag(),
    "isHidden": True,
    "columns": [
        {"name": "ExcelFilePath", "dataType": "string", "sourceColumn": "ExcelFilePath",
         "lineageTag": tag(), "summarizeBy": "none", "isHidden": True}
    ],
    "partitions": [{
        "name": "ExcelFilePath",
        "mode": "import",
        "source": {
            "type": "m",
            "expression": [
                "let",
                "    Source = \"C:\\\\RideShareAnalytics\\\\RideShareAnalytics_Data.xlsx\" meta [IsParameterQuery=true, Type=\"Text\", IsParameterQueryRequired=true]",
                "in",
                "    Source"
            ]
        }
    }]
})

# ── Fact_Trips ──────────────────────────────────────────────────────────

fact_trips_tag = tag()
MODEL["model"]["tables"].append({
    "name": "Fact_Trips",
    "lineageTag": fact_trips_tag,
    "columns": [
        {"name": "TripKey",            "dataType": "int64",    "sourceColumn": "TripKey",            "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "DateKey",            "dataType": "int64",    "sourceColumn": "DateKey",            "lineageTag": tag(), "summarizeBy": "none", "isHidden": True},
        {"name": "DriverKey",          "dataType": "int64",    "sourceColumn": "DriverKey",          "lineageTag": tag(), "summarizeBy": "none", "isHidden": True},
        {"name": "RiderKey",           "dataType": "int64",    "sourceColumn": "RiderKey",           "lineageTag": tag(), "summarizeBy": "none", "isHidden": True},
        {"name": "LocationKey",        "dataType": "int64",    "sourceColumn": "LocationKey",        "lineageTag": tag(), "summarizeBy": "none", "isHidden": True},
        {"name": "PromoKey",           "dataType": "int64",    "sourceColumn": "PromoKey",           "lineageTag": tag(), "summarizeBy": "none", "isHidden": True},
        {"name": "HourOfDay",          "dataType": "int64",    "sourceColumn": "HourOfDay",          "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "Fare",               "dataType": "decimal",  "sourceColumn": "Fare",               "lineageTag": tag(), "summarizeBy": "sum",     "formatString": "₹#,##0.00"},
        {"name": "DistanceKM",         "dataType": "double",   "sourceColumn": "DistanceKM",         "lineageTag": tag(), "summarizeBy": "sum",     "formatString": "#,##0.0"},
        {"name": "DurationMin",        "dataType": "int64",    "sourceColumn": "DurationMin",        "lineageTag": tag(), "summarizeBy": "sum",     "formatString": "#,##0"},
        {"name": "WaitTimeMin",        "dataType": "double",   "sourceColumn": "WaitTimeMin",        "lineageTag": tag(), "summarizeBy": "average", "formatString": "#,##0.0"},
        {"name": "Tip",                "dataType": "decimal",  "sourceColumn": "Tip",                "lineageTag": tag(), "summarizeBy": "sum",     "formatString": "₹#,##0.00"},
        {"name": "SurgeMultiplier",    "dataType": "double",   "sourceColumn": "SurgeMultiplier",    "lineageTag": tag(), "summarizeBy": "average", "formatString": "#,##0.0x"},
        {"name": "Status",             "dataType": "string",   "sourceColumn": "Status",             "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "CancellationReason", "dataType": "string",   "sourceColumn": "CancellationReason", "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "PaymentMethod",      "dataType": "string",   "sourceColumn": "PaymentMethod",      "lineageTag": tag(), "summarizeBy": "none"},
    ],
    "partitions": [{
        "name": "Fact_Trips",
        "mode": "import",
        "source": {
            "type": "m",
            "expression": [
                "let",
                "    Source = Excel.Workbook(File.Contents(ExcelFilePath), null, true),",
                "    Sheet = Source{[Item=\"Fact_Trips\",Kind=\"Sheet\"]}[Data],",
                "    Headers = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),",
                "",
                "    // Data cleaning: standardize PaymentMethod casing (~6% had inconsistent case)",
                "    CleanPayment = Table.TransformColumns(Headers, {{\"PaymentMethod\", each ",
                "        if Text.Lower(_) = \"upi\" then \"UPI\"",
                "        else if Text.Lower(_) = \"cash\" then \"Cash\"",
                "        else if Text.Lower(_) = \"card\" then \"Card\"",
                "        else if Text.Lower(_) = \"wallet\" then \"Wallet\"",
                "        else _, type text}}),",
                "",
                "    // Data cleaning: replace null tips with 0 for completed trips (~8% had missing tips)",
                "    CleanTips = Table.ReplaceValue(CleanPayment, null, 0, Replacer.ReplaceValue, {\"Tip\"}),",
                "",
                "    Typed = Table.TransformColumnTypes(CleanTips, {",
                "        {\"TripKey\", Int64.Type},",
                "        {\"DateKey\", Int64.Type},",
                "        {\"DriverKey\", Int64.Type},",
                "        {\"RiderKey\", Int64.Type},",
                "        {\"LocationKey\", Int64.Type},",
                "        {\"PromoKey\", Int64.Type},",
                "        {\"HourOfDay\", Int64.Type},",
                "        {\"Fare\", Currency.Type},",
                "        {\"DistanceKM\", type number},",
                "        {\"DurationMin\", Int64.Type},",
                "        {\"WaitTimeMin\", type number},",
                "        {\"Tip\", Currency.Type},",
                "        {\"SurgeMultiplier\", type number},",
                "        {\"Status\", type text},",
                "        {\"CancellationReason\", type text},",
                "        {\"PaymentMethod\", type text}",
                "    })",
                "in",
                "    Typed"
            ]
        }
    }]
})

# ── Dim_Driver ──────────────────────────────────────────────────────────

MODEL["model"]["tables"].append({
    "name": "Dim_Driver",
    "lineageTag": tag(),
    "columns": [
        {"name": "DriverKey",   "dataType": "int64",    "sourceColumn": "DriverKey",   "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "DriverName",  "dataType": "string",   "sourceColumn": "DriverName",  "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "VehicleType", "dataType": "string",   "sourceColumn": "VehicleType", "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "Rating",      "dataType": "double",   "sourceColumn": "Rating",      "lineageTag": tag(), "summarizeBy": "average", "formatString": "#,##0.0"},
        {"name": "JoinDate",    "dataType": "dateTime", "sourceColumn": "JoinDate",    "lineageTag": tag(), "summarizeBy": "none", "formatString": "Short Date"},
        {"name": "HomeCity",    "dataType": "string",   "sourceColumn": "HomeCity",    "lineageTag": tag(), "summarizeBy": "none"},
    ],
    "partitions": [{
        "name": "Dim_Driver",
        "mode": "import",
        "source": {
            "type": "m",
            "expression": [
                "let",
                "    Source = Excel.Workbook(File.Contents(ExcelFilePath), null, true),",
                "    Sheet = Source{[Item=\"Dim_Driver\",Kind=\"Sheet\"]}[Data],",
                "    Headers = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),",
                "    Typed = Table.TransformColumnTypes(Headers, {",
                "        {\"DriverKey\", Int64.Type},",
                "        {\"DriverName\", type text},",
                "        {\"VehicleType\", type text},",
                "        {\"Rating\", type number},",
                "        {\"JoinDate\", type date},",
                "        {\"HomeCity\", type text}",
                "    })",
                "in",
                "    Typed"
            ]
        }
    }]
})

# ── Dim_Rider ───────────────────────────────────────────────────────────

MODEL["model"]["tables"].append({
    "name": "Dim_Rider",
    "lineageTag": tag(),
    "columns": [
        {"name": "RiderKey",    "dataType": "int64",    "sourceColumn": "RiderKey",    "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "RiderName",   "dataType": "string",   "sourceColumn": "RiderName",   "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "SignupDate",  "dataType": "dateTime", "sourceColumn": "SignupDate",  "lineageTag": tag(), "summarizeBy": "none", "formatString": "Short Date"},
        {"name": "LoyaltyTier", "dataType": "string",   "sourceColumn": "LoyaltyTier", "lineageTag": tag(), "summarizeBy": "none"},
    ],
    "partitions": [{
        "name": "Dim_Rider",
        "mode": "import",
        "source": {
            "type": "m",
            "expression": [
                "let",
                "    Source = Excel.Workbook(File.Contents(ExcelFilePath), null, true),",
                "    Sheet = Source{[Item=\"Dim_Rider\",Kind=\"Sheet\"]}[Data],",
                "    Headers = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),",
                "    Typed = Table.TransformColumnTypes(Headers, {",
                "        {\"RiderKey\", Int64.Type},",
                "        {\"RiderName\", type text},",
                "        {\"SignupDate\", type date},",
                "        {\"LoyaltyTier\", type text}",
                "    })",
                "in",
                "    Typed"
            ]
        }
    }]
})

# ── Dim_Location ────────────────────────────────────────────────────────

MODEL["model"]["tables"].append({
    "name": "Dim_Location",
    "lineageTag": tag(),
    "columns": [
        {"name": "LocationKey", "dataType": "int64",  "sourceColumn": "LocationKey", "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "City",        "dataType": "string", "sourceColumn": "City",        "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "Zone",        "dataType": "string", "sourceColumn": "Zone",        "lineageTag": tag(), "summarizeBy": "none"},
    ],
    "partitions": [{
        "name": "Dim_Location",
        "mode": "import",
        "source": {
            "type": "m",
            "expression": [
                "let",
                "    Source = Excel.Workbook(File.Contents(ExcelFilePath), null, true),",
                "    Sheet = Source{[Item=\"Dim_Location\",Kind=\"Sheet\"]}[Data],",
                "    Headers = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),",
                "    Typed = Table.TransformColumnTypes(Headers, {",
                "        {\"LocationKey\", Int64.Type},",
                "        {\"City\", type text},",
                "        {\"Zone\", type text}",
                "    })",
                "in",
                "    Typed"
            ]
        }
    }]
})

# ── Dim_Date (marked as Date table) ─────────────────────────────────────

dim_date_tag = tag()
MODEL["model"]["tables"].append({
    "name": "Dim_Date",
    "lineageTag": dim_date_tag,
    "dataCategory": "Time",
    "isHidden": False,
    "columns": [
        {"name": "DateKey",   "dataType": "int64",    "sourceColumn": "DateKey",   "lineageTag": tag(), "summarizeBy": "none", "isHidden": True},
        {"name": "Date",      "dataType": "dateTime", "sourceColumn": "Date",      "lineageTag": tag(), "summarizeBy": "none", "isKey": True, "formatString": "Short Date"},
        {"name": "Year",      "dataType": "int64",    "sourceColumn": "Year",      "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "Quarter",   "dataType": "string",   "sourceColumn": "Quarter",   "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "MonthName", "dataType": "string",   "sourceColumn": "MonthName", "lineageTag": tag(), "summarizeBy": "none", "sortByColumn": "MonthNum"},
        {"name": "MonthNum",  "dataType": "int64",    "sourceColumn": "MonthNum",  "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "Day",       "dataType": "int64",    "sourceColumn": "Day",       "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "DayOfWeek", "dataType": "string",   "sourceColumn": "DayOfWeek", "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "WeekNum",   "dataType": "int64",    "sourceColumn": "WeekNum",   "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "IsWeekend", "dataType": "boolean",  "sourceColumn": "IsWeekend", "lineageTag": tag(), "summarizeBy": "none"},
    ],
    "partitions": [{
        "name": "Dim_Date",
        "mode": "import",
        "source": {
            "type": "m",
            "expression": [
                "let",
                "    Source = Excel.Workbook(File.Contents(ExcelFilePath), null, true),",
                "    Sheet = Source{[Item=\"Dim_Date\",Kind=\"Sheet\"]}[Data],",
                "    Headers = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),",
                "    Typed = Table.TransformColumnTypes(Headers, {",
                "        {\"DateKey\", Int64.Type},",
                "        {\"Date\", type date},",
                "        {\"Year\", Int64.Type},",
                "        {\"Quarter\", type text},",
                "        {\"MonthName\", type text},",
                "        {\"MonthNum\", Int64.Type},",
                "        {\"Day\", Int64.Type},",
                "        {\"DayOfWeek\", type text},",
                "        {\"WeekNum\", Int64.Type},",
                "        {\"IsWeekend\", type logical}",
                "    })",
                "in",
                "    Typed"
            ]
        }
    }]
})

# ── Dim_Promo ───────────────────────────────────────────────────────────

MODEL["model"]["tables"].append({
    "name": "Dim_Promo",
    "lineageTag": tag(),
    "columns": [
        {"name": "PromoKey",     "dataType": "int64",  "sourceColumn": "PromoKey",     "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "PromoCode",    "dataType": "string", "sourceColumn": "PromoCode",    "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "CampaignName", "dataType": "string", "sourceColumn": "CampaignName", "lineageTag": tag(), "summarizeBy": "none"},
        {"name": "DiscountPct",  "dataType": "int64",  "sourceColumn": "DiscountPct",  "lineageTag": tag(), "summarizeBy": "average", "formatString": "#,##0'%'"},
    ],
    "partitions": [{
        "name": "Dim_Promo",
        "mode": "import",
        "source": {
            "type": "m",
            "expression": [
                "let",
                "    Source = Excel.Workbook(File.Contents(ExcelFilePath), null, true),",
                "    Sheet = Source{[Item=\"Dim_Promo\",Kind=\"Sheet\"]}[Data],",
                "    Headers = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),",
                "    Typed = Table.TransformColumnTypes(Headers, {",
                "        {\"PromoKey\", Int64.Type},",
                "        {\"PromoCode\", type text},",
                "        {\"CampaignName\", type text},",
                "        {\"DiscountPct\", Int64.Type}",
                "    })",
                "in",
                "    Typed"
            ]
        }
    }]
})

# ── _Measures table (dedicated display table) ──────────────────────────

MODEL["model"]["tables"].append({
    "name": "_Measures",
    "lineageTag": tag(),
    "isHidden": False,
    "columns": [
        {"name": "MeasureHelper", "dataType": "int64", "sourceColumn": "MeasureHelper",
         "lineageTag": tag(), "summarizeBy": "none", "isHidden": True}
    ],
    "measures": [
        # ── Core Metrics ────────────────────────────────────────────
        {
            "name": "Total Trips",
            "expression": "COUNTROWS(Fact_Trips)",
            "formatString": "#,##0",
            "displayFolder": "Core Metrics",
            "lineageTag": tag()
        },
        {
            "name": "Total Revenue",
            "expression": [
                "CALCULATE(",
                "    SUM(Fact_Trips[Fare]),",
                "    Fact_Trips[Status] = \"Completed\"",
                ")"
            ],
            "formatString": "₹#,##0",
            "displayFolder": "Core Metrics",
            "lineageTag": tag()
        },
        {
            "name": "Total Tips",
            "expression": [
                "CALCULATE(",
                "    SUM(Fact_Trips[Tip]),",
                "    Fact_Trips[Status] = \"Completed\"",
                ")"
            ],
            "formatString": "₹#,##0",
            "displayFolder": "Core Metrics",
            "lineageTag": tag()
        },
        {
            "name": "Avg Fare",
            "expression": [
                "CALCULATE(",
                "    AVERAGE(Fact_Trips[Fare]),",
                "    Fact_Trips[Status] = \"Completed\"",
                ")"
            ],
            "formatString": "₹#,##0.00",
            "displayFolder": "Core Metrics",
            "lineageTag": tag()
        },
        {
            "name": "Cancellation Rate %",
            "expression": [
                "VAR TotalTrips = COUNTROWS(Fact_Trips)",
                "VAR CancelledTrips =",
                "    CALCULATE(",
                "        COUNTROWS(Fact_Trips),",
                "        Fact_Trips[Status] = \"Cancelled\"",
                "    )",
                "RETURN",
                "    DIVIDE(CancelledTrips, TotalTrips, 0)"
            ],
            "formatString": "0.0%",
            "displayFolder": "Core Metrics",
            "lineageTag": tag()
        },
        {
            "name": "Completed Trips",
            "expression": [
                "CALCULATE(",
                "    COUNTROWS(Fact_Trips),",
                "    Fact_Trips[Status] = \"Completed\"",
                ")"
            ],
            "formatString": "#,##0",
            "displayFolder": "Core Metrics",
            "lineageTag": tag()
        },
        {
            "name": "Cancelled Trips",
            "expression": [
                "CALCULATE(",
                "    COUNTROWS(Fact_Trips),",
                "    Fact_Trips[Status] = \"Cancelled\"",
                ")"
            ],
            "formatString": "#,##0",
            "displayFolder": "Core Metrics",
            "lineageTag": tag()
        },

        # ── Efficiency Ratios ───────────────────────────────────────
        {
            "name": "Revenue per KM",
            "expression": [
                "VAR TotalRev = [Total Revenue]",
                "VAR TotalDist =",
                "    CALCULATE(",
                "        SUM(Fact_Trips[DistanceKM]),",
                "        Fact_Trips[Status] = \"Completed\"",
                "    )",
                "RETURN",
                "    DIVIDE(TotalRev, TotalDist, 0)"
            ],
            "formatString": "₹#,##0.00",
            "displayFolder": "Efficiency",
            "lineageTag": tag()
        },
        {
            "name": "Avg Wait Time",
            "expression": "AVERAGE(Fact_Trips[WaitTimeMin])",
            "formatString": "#,##0.0 \" min\"",
            "displayFolder": "Efficiency",
            "lineageTag": tag()
        },
        {
            "name": "Revenue per Driver",
            "expression": [
                "DIVIDE(",
                "    [Total Revenue],",
                "    DISTINCTCOUNT(Fact_Trips[DriverKey]),",
                "    0",
                ")"
            ],
            "formatString": "₹#,##0",
            "displayFolder": "Efficiency",
            "lineageTag": tag()
        },

        # ── Time Intelligence ───────────────────────────────────────
        {
            "name": "MTD Revenue",
            "expression": [
                "CALCULATE(",
                "    [Total Revenue],",
                "    DATESMTD(Dim_Date[Date])",
                ")"
            ],
            "formatString": "₹#,##0",
            "displayFolder": "Time Intelligence",
            "lineageTag": tag()
        },
        {
            "name": "YoY Growth %",
            "expression": [
                "VAR CurrentRev = [Total Revenue]",
                "VAR PriorYearRev =",
                "    CALCULATE(",
                "        [Total Revenue],",
                "        SAMEPERIODLASTYEAR(Dim_Date[Date])",
                "    )",
                "RETURN",
                "    DIVIDE(",
                "        CurrentRev - PriorYearRev,",
                "        PriorYearRev,",
                "        BLANK()",
                "    )"
            ],
            "formatString": "0.0%",
            "displayFolder": "Time Intelligence",
            "lineageTag": tag()
        },
        {
            "name": "Rolling 7-Day Avg Trips",
            "expression": [
                "CALCULATE(",
                "    DIVIDE(",
                "        COUNTROWS(Fact_Trips),",
                "        7,",
                "        0",
                "    ),",
                "    DATESINPERIOD(",
                "        Dim_Date[Date],",
                "        MAX(Dim_Date[Date]),",
                "        -7,",
                "        DAY",
                "    )",
                ")"
            ],
            "formatString": "#,##0",
            "displayFolder": "Time Intelligence",
            "lineageTag": tag()
        },

        # ── Ranking ─────────────────────────────────────────────────
        {
            "name": "Driver Revenue Rank",
            "expression": [
                "IF(",
                "    HASONEVALUE(Dim_Driver[DriverName]),",
                "    RANKX(",
                "        ALL(Dim_Driver[DriverName]),",
                "        [Total Revenue],",
                "        ,",
                "        DESC,",
                "        Dense",
                "    )",
                ")"
            ],
            "formatString": "#,##0",
            "displayFolder": "Ranking",
            "lineageTag": tag()
        },
        {
            "name": "Driver Rank in City",
            "expression": [
                "IF(",
                "    HASONEVALUE(Dim_Driver[DriverName]),",
                "    RANKX(",
                "        ALLEXCEPT(Dim_Driver, Dim_Driver[HomeCity]),",
                "        [Total Revenue],",
                "        ,",
                "        DESC,",
                "        Dense",
                "    )",
                ")"
            ],
            "formatString": "#,##0",
            "displayFolder": "Ranking",
            "lineageTag": tag()
        },
        {
            "name": "Avg Surge at Cancellation",
            "expression": [
                "CALCULATE(",
                "    AVERAGE(Fact_Trips[SurgeMultiplier]),",
                "    Fact_Trips[Status] = \"Cancelled\"",
                ")"
            ],
            "formatString": "#,##0.00x",
            "displayFolder": "Core Metrics",
            "lineageTag": tag()
        },
        {
            "name": "Promo Trip Count",
            "expression": [
                "CALCULATE(",
                "    COUNTROWS(Fact_Trips),",
                "    Fact_Trips[PromoKey] <> 0",
                ")"
            ],
            "formatString": "#,##0",
            "displayFolder": "Promotions",
            "lineageTag": tag()
        },
        {
            "name": "Promo Revenue",
            "expression": [
                "CALCULATE(",
                "    [Total Revenue],",
                "    Fact_Trips[PromoKey] <> 0",
                ")"
            ],
            "formatString": "₹#,##0",
            "displayFolder": "Promotions",
            "lineageTag": tag()
        },
    ],
    "partitions": [{
        "name": "_Measures",
        "mode": "import",
        "source": {
            "type": "calculated",
            "expression": "ROW(\"MeasureHelper\", 1)"
        }
    }]
})

# ── Relationships ───────────────────────────────────────────────────────

MODEL["model"]["relationships"] = [
    {
        "name": "Fact_Trips_to_Dim_Date",
        "fromTable": "Fact_Trips",
        "fromColumn": "DateKey",
        "toTable": "Dim_Date",
        "toColumn": "DateKey",
        "crossFilteringBehavior": "oneDirection",
        "isActive": True
    },
    {
        "name": "Fact_Trips_to_Dim_Driver",
        "fromTable": "Fact_Trips",
        "fromColumn": "DriverKey",
        "toTable": "Dim_Driver",
        "toColumn": "DriverKey",
        "crossFilteringBehavior": "oneDirection",
        "isActive": True
    },
    {
        "name": "Fact_Trips_to_Dim_Rider",
        "fromTable": "Fact_Trips",
        "fromColumn": "RiderKey",
        "toTable": "Dim_Rider",
        "toColumn": "RiderKey",
        "crossFilteringBehavior": "oneDirection",
        "isActive": True
    },
    {
        "name": "Fact_Trips_to_Dim_Location",
        "fromTable": "Fact_Trips",
        "fromColumn": "LocationKey",
        "toTable": "Dim_Location",
        "toColumn": "LocationKey",
        "crossFilteringBehavior": "oneDirection",
        "isActive": True
    },
    {
        "name": "Fact_Trips_to_Dim_Promo",
        "fromTable": "Fact_Trips",
        "fromColumn": "PromoKey",
        "toTable": "Dim_Promo",
        "toColumn": "PromoKey",
        "crossFilteringBehavior": "oneDirection",
        "isActive": True
    },
]

# ── Write model.bim ────────────────────────────────────────────────────

output_path = "/Users/bsinga1/Desktop/Powerbi/RideShareAnalytics/RideShareAnalytics.SemanticModel/model.bim"
with open(output_path, "w", encoding="utf-8") as f:
    json.dump(MODEL, f, indent=2, ensure_ascii=False, default=str)

print(f"✅  model.bim written → {output_path}")
print(f"    Tables: {len(MODEL['model']['tables'])}")
print(f"    Relationships: {len(MODEL['model']['relationships'])}")

measures_table = next(t for t in MODEL["model"]["tables"] if t["name"] == "_Measures")
print(f"    DAX Measures: {len(measures_table['measures'])}")
