from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import pandas as pd
import os
from pathlib import Path
from io import BytesIO
from datetime import datetime
from typing import Optional, Dict, Any, List

app = FastAPI(title="Accessories Sales Dashboard", version="1.0.0")

# -------------------- CORS --------------------
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# -------------------- Paths --------------------
BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"

# If you have static folder, mount it (optional)
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# -------------------- Request Model (IMPORTANT FIX) --------------------
class FilterRequest(BaseModel):
    quarter: Optional[str] = ""
    month: Optional[str] = ""
    location: Optional[str] = ""
    model: Optional[str] = ""

# -------------------- Excel Load + Column Normalize --------------------
df = pd.DataFrame()

CANON_COLS = {
    "Fiscal Quarter": "Fiscal Quarter",
    "Fiscal Month": "Fiscal Month",
    "Location": "Location",
    "Model Group": "Model Group",

    # these often come with extra spaces -> normalize
    "No of Billied Ros": "No of Billied Ros",
    "Acc Sale throughROs (GNDP)In Rs": "Acc Sale throughROs (GNDP)In Rs",
    "Acc Sale throughROs (MRP) In Rs": "Acc Sale throughROs (MRP) In Rs",
    "No of Counter ROs": "No of Counter ROs",
    "Acc Sale throughCounter (GNDP) In Rs": "Acc Sale throughCounter (GNDP) In Rs",
    "Acc Sale throughCounter (MRP) In Rs": "Acc Sale throughCounter (MRP) In Rs",
    "Acc Revenue (GNDP) / RO": "Acc Revenue (GNDP) / RO",
    "Acc Revenue (MRP) / RO": "Acc Revenue (MRP) / RO",
}

NUMERIC_COLS = [
    "No of Billied Ros",
    "Acc Sale throughROs (GNDP)In Rs",
    "Acc Sale throughROs (MRP) In Rs",
    "No of Counter ROs",
    "Acc Sale throughCounter (GNDP) In Rs",
    "Acc Sale throughCounter (MRP) In Rs",
    "Acc Revenue (GNDP) / RO",
    "Acc Revenue (MRP) / RO",
]

INDIAN_FINANCIAL_MONTHS = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']

quarters: List[str] = []
months: List[str] = []
locations: List[str] = []
models: List[str] = []

def _clean_col_name(c: str) -> str:
    # remove leading/trailing + collapse multiple spaces
    return " ".join(str(c).replace("\t", " ").strip().split())

def load_excel() -> pd.DataFrame:
    global quarters, months, locations, models

    possible_paths = [
        "/mnt/data/Accessories.xlsx",                 # Render Disk (most common)
        "/mnt/user-data/uploads/Accessories.xlsx",    # some setups
        str(BASE_DIR / "Accessories.xlsx"),
        "Accessories.xlsx",
        "./Accessories.xlsx",
        os.path.join(os.getcwd(), "Accessories.xlsx"),
    ]

    excel_path = None
    for p in possible_paths:
        if os.path.exists(p):
            excel_path = p
            break

    if not excel_path:
        print("Excel file not found. Tried:")
        for p in possible_paths:
            print(" -", p)
        return pd.DataFrame()

    print(f"Loading Excel from: {excel_path}")

    _df = pd.read_excel(excel_path, sheet_name="Sheet1")

    # Clean columns
    _df.columns = [_clean_col_name(c) for c in _df.columns]

    # Rename to canonical if possible
    rename_map = {}
    for c in _df.columns:
        # match by cleaned name
        if c in CANON_COLS:
            rename_map[c] = CANON_COLS[c]

    _df = _df.rename(columns=rename_map)

    # Ensure required columns exist
    required = ["Fiscal Quarter", "Fiscal Month", "Location", "Model Group"]
    for r in required:
        if r not in _df.columns:
            raise ValueError(f"Missing required column: {r}")

    # Convert to string for filter columns
    _df["Fiscal Quarter"] = _df["Fiscal Quarter"].astype(str).str.strip()
    _df["Fiscal Month"] = _df["Fiscal Month"].astype(str).str.strip()
    _df["Location"] = _df["Location"].astype(str).str.strip()
    _df["Model Group"] = _df["Model Group"].astype(str).str.strip()

    # Convert numeric columns
    for c in NUMERIC_COLS:
        if c in _df.columns:
            _df[c] = pd.to_numeric(_df[c], errors="coerce").fillna(0)

    # Dropdown options
    quarters = sorted([x for x in _df["Fiscal Quarter"].dropna().unique().tolist() if x != "nan"])
    all_months = [x for x in _df["Fiscal Month"].dropna().unique().tolist() if x != "nan"]
    months = sorted(all_months, key=lambda x: INDIAN_FINANCIAL_MONTHS.index(x) if x in INDIAN_FINANCIAL_MONTHS else 999)

    locations = sorted([x for x in _df["Location"].dropna().unique().tolist() if x != "nan"])
    models = sorted([x for x in _df["Model Group"].dropna().unique().tolist() if x != "nan"])

    print(f"Rows: {len(_df)} | Quarters: {quarters} | Months: {months} | Locations: {locations} | Models: {models}")

    return _df

try:
    df = load_excel()
except Exception as e:
    print("Excel load error:", e)
    df = pd.DataFrame()

def apply_filter(_df: pd.DataFrame, req: FilterRequest) -> pd.DataFrame:
    f = _df.copy()

    if req.quarter and req.quarter.strip():
        f = f[f["Fiscal Quarter"].astype(str) == req.quarter.strip()]

    if req.month and req.month.strip():
        f = f[f["Fiscal Month"].astype(str) == req.month.strip()]

    if req.location and req.location.strip():
        f = f[f["Location"].astype(str) == req.location.strip()]

    if req.model and req.model.strip():
        f = f[f["Model Group"].astype(str) == req.model.strip()]

    return f

def compute_totals(_df: pd.DataFrame) -> Dict[str, Any]:
    totals: Dict[str, Any] = {}

    for c in NUMERIC_COLS:
        totals[c] = float(_df[c].sum()) if c in _df.columns else 0.0

    # Make count columns integers
    totals["No of Billied Ros"] = int(_df["No of Billied Ros"].sum()) if "No of Billied Ros" in _df.columns else 0
    totals["No of Counter ROs"] = int(_df["No of Counter ROs"].sum()) if "No of Counter ROs" in _df.columns else 0

    return totals

# -------------------- Routes --------------------
@app.get("/")
def read_root():
    index_file = TEMPLATES_DIR / "index.html"
    if not index_file.exists():
        raise HTTPException(status_code=500, detail="templates/index.html not found")
    return FileResponse(str(index_file), media_type="text/html")

@app.get("/api/stats")
def get_stats():
    return {
        "total_records": int(len(df)) if not df.empty else 0,
        "total_quarters": len(quarters),
        "total_months": len(months),
        "total_locations": len(locations),
        "total_models": len(models),
        "data_last_updated": datetime.now().isoformat(),
    }

@app.get("/api/filter-options")
def get_filter_options():
    return {
        "quarters": quarters,
        "months": months,
        "locations": locations,
        "models": models,
    }

@app.post("/api/get-data")
def get_data(req: FilterRequest):
    if df.empty:
        return {"data": [], "totals": compute_totals(pd.DataFrame(columns=NUMERIC_COLS)), "count": 0}

    filtered_df = apply_filter(df, req)
    totals = compute_totals(filtered_df)

    # return records
    records = filtered_df.to_dict("records")

    return {
        "data": records,
        "totals": totals,
        "count": int(len(filtered_df)),
    }

@app.post("/api/export-excel")
def export_excel(req: FilterRequest):
    if df.empty:
        raise HTTPException(status_code=400, detail="No data loaded")

    filtered_df = apply_filter(df, req)

    output = BytesIO()

    try:
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            filtered_df.to_excel(writer, sheet_name="Sales Data", index=False)

            # Totals row
            totals = compute_totals(filtered_df)

            totals_row = {c: "" for c in filtered_df.columns}
            totals_row["Fiscal Quarter"] = "TOTAL"

            for c in NUMERIC_COLS:
                if c in filtered_df.columns:
                    totals_row[c] = totals[c]

            totals_df = pd.DataFrame([totals_row])
            start_row = len(filtered_df) + 2
            totals_df.to_excel(writer, sheet_name="Sales Data", startrow=start_row - 1, index=False, header=False)

            ws = writer.sheets["Sales Data"]

            thin = Side(style="thin", color="D1D5DB")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)

            # Header format
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="4F46E5", end_color="4F46E5", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border

            # Data format
            for row in ws.iter_rows(min_row=2, max_row=len(filtered_df) + 1):
                for cell in row:
                    cell.border = border
                    cell.alignment = Alignment(horizontal="right" if isinstance(cell.value, (int, float)) else "left")

            # Totals row format
            totals_excel_row = len(filtered_df) + 2
            for cell in ws[totals_excel_row]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="E0E7FF", end_color="E0E7FF", fill_type="solid")
                cell.border = border

            # Column widths
            for col_cells in ws.columns:
                col_letter = col_cells[0].column_letter
                max_len = 0
                for c in col_cells:
                    v = "" if c.value is None else str(c.value)
                    if len(v) > max_len:
                        max_len = len(v)
                ws.column_dimensions[col_letter].width = min(max_len + 2, 50)

        output.seek(0)

        filename = "accessories_sales_filtered.xlsx"
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"},
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Export failed: {str(e)}")

@app.get("/api/health")
def health_check():
    return {"status": "ok", "timestamp": datetime.now().isoformat()}
