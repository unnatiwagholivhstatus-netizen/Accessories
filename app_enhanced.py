from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import os
from typing import List, Optional
from io import BytesIO
import json
from datetime import datetime

app = FastAPI(title="Accessories Sales Dashboard", version="1.0.0")

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Load the Excel file
excel_path = None
possible_paths = [
    '/mnt/user-data/uploads/Accessories.xlsx',
    'Accessories.xlsx',
    './Accessories.xlsx',
    os.path.join(os.getcwd(), 'Accessories.xlsx'),
]

try:
    for path in possible_paths:
        if os.path.exists(path):
            excel_path = path
            break
    
    if excel_path:
        df = pd.read_excel(excel_path, sheet_name='Sheet1')
        print(f"Excel file loaded successfully from: {excel_path}")
    else:
        print(f"Excel file not found in any of these locations:")
        for path in possible_paths:
            print(f"   - {path}")
        df = pd.DataFrame()
except Exception as e:
    print(f"Error loading Excel file: {e}")
    df = pd.DataFrame()

# Clean column names
df.columns = df.columns.str.strip()

# Ensure data types
df['Fiscal Quarter'] = df['Fiscal Quarter'].astype(str)
df['Fiscal Month'] = df['Fiscal Month'].astype(str)
df['Location'] = df['Location'].astype(str)
df['Model Group'] = df['Model Group'].astype(str)

# Convert numeric columns
numeric_cols = [
    'No of Billied Ros  ',
    'Acc Sale throughROs (GNDP)In Rs',
    'Acc Sale throughROs (MRP) In Rs',
    'No of Counter ROs  ',
    'Acc Sale throughCounter (GNDP) In Rs',
    'Acc Sale throughCounter (MRP) In Rs',
    'Acc Revenue (GNDP) / RO  ',
    'Acc Revenue (MRP) / RO  '
]

for col in numeric_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

# Indian Financial Year Month Order
indian_financial_months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']

# Get unique values for dropdowns
quarters = sorted([str(x) for x in df['Fiscal Quarter'].unique().tolist() if pd.notna(x)])

# Sort months in Indian financial year order
all_months = [str(x) for x in df['Fiscal Month'].unique().tolist() if pd.notna(x)]
months = sorted(all_months, key=lambda x: indian_financial_months.index(x) if x in indian_financial_months else 999)

locations = sorted([str(x) for x in df['Location'].unique().tolist() if pd.notna(x)])
models = sorted([str(x) for x in df['Model Group'].unique().tolist() if pd.notna(x)])

print(f"Data loaded: {len(df)} rows")
print(f"Quarters: {quarters}")
print(f"Months (Indian Financial Order): {months}")
print(f"Divisions: {locations}")
print(f"Models: {models}")

@app.get("/")
def read_root():
    """Serve the main HTML page"""
    return FileResponse("templates/index.html", media_type="text/html")

@app.get("/api/stats")
def get_stats():
    """Get overall statistics"""
    return {
        "total_records": len(df),
        "total_quarters": len(quarters),
        "total_months": len(months),
        "total_locations": len(locations),
        "total_models": len(models),
        "data_last_updated": datetime.now().isoformat()
    }

@app.get("/api/filter-options")
def get_filter_options():
    """Get all available filter options"""
    return {
        "quarters": quarters,
        "months": months,
        "locations": locations,
        "models": models
    }

@app.post("/api/get-data")
def get_data(
    quarter: Optional[str] = None,
    month: Optional[str] = None,
    location: Optional[str] = None,
    model: Optional[str] = None
):
    """Get filtered data based on selected criteria"""
    filtered_df = df.copy()
    
    if quarter and quarter != "":
        filtered_df = filtered_df[filtered_df['Fiscal Quarter'].astype(str) == str(quarter)]
    if month and month != "":
        filtered_df = filtered_df[filtered_df['Fiscal Month'].astype(str) == str(month)]
    if location and location != "":
        filtered_df = filtered_df[filtered_df['Location'].astype(str) == str(location)]
    if model and model != "":
        filtered_df = filtered_df[filtered_df['Model Group'].astype(str) == str(model)]
    
    # Convert to records
    records = filtered_df.to_dict('records')
    
    # Calculate totals - handle missing columns gracefully
    totals = {}
    for col in numeric_cols:
        if col in filtered_df.columns:
            totals[col] = float(filtered_df[col].sum())
        else:
            totals[col] = 0.0
    
    # Ensure proper count values - these are the billied ros and counter ros counts
    if 'No of Billied Ros  ' in filtered_df.columns:
        totals['No of Billied Ros  '] = int(filtered_df['No of Billied Ros  '].sum())
    else:
        totals['No of Billied Ros  '] = 0
        
    if 'No of Counter ROs  ' in filtered_df.columns:
        totals['No of Counter ROs  '] = int(filtered_df['No of Counter ROs  '].sum())
    else:
        totals['No of Counter ROs  '] = 0
    
    return {
        "data": records,
        "totals": totals,
        "count": len(filtered_df)
    }

@app.post("/api/export-excel")
def export_excel(
    quarter: Optional[str] = None,
    month: Optional[str] = None,
    location: Optional[str] = None,
    model: Optional[str] = None
):
    """Export filtered data as Excel file"""
    filtered_df = df.copy()
    
    if quarter and quarter != "":
        filtered_df = filtered_df[filtered_df['Fiscal Quarter'].astype(str) == str(quarter)]
    if month and month != "":
        filtered_df = filtered_df[filtered_df['Fiscal Month'].astype(str) == str(month)]
    if location and location != "":
        filtered_df = filtered_df[filtered_df['Location'].astype(str) == str(location)]
    if model and model != "":
        filtered_df = filtered_df[filtered_df['Model Group'].astype(str) == str(model)]
    
    # Create Excel file in memory
    output = BytesIO()
    
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            filtered_df.to_excel(writer, sheet_name='Sales Data', index=False)
            
            # Add totals row
            totals_row = {}
            totals_row['Fiscal Quarter'] = 'TOTAL'
            for col in numeric_cols:
                if col in filtered_df.columns:
                    totals_row[col] = filtered_df[col].sum()
            
            totals_df = pd.DataFrame([totals_row])
            totals_df.to_excel(writer, sheet_name='Sales Data', startrow=len(filtered_df)+1, index=False, header=False)
            
            # Get the worksheet to format it
            worksheet = writer.sheets['Sales Data']
            
            # Apply formatting
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # Format header
            for cell in worksheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="4F46E5", end_color="4F46E5", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border
            
            # Format data cells
            for row in worksheet.iter_rows(min_row=2, max_row=len(filtered_df)+2):
                for cell in row:
                    cell.border = thin_border
                    if isinstance(cell.value, (int, float)):
                        cell.alignment = Alignment(horizontal="right")
            
            # Format totals row
            totals_start_row = len(filtered_df) + 2
            for cell in worksheet[totals_start_row]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="E0E7FF", end_color="E0E7FF", fill_type="solid")
                cell.border = thin_border
            
            # Adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        output.seek(0)
        return StreamingResponse(
            iter([output.getvalue()]),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=accessories_sales.xlsx"}
        )
    
    except Exception as e:
        print(f"Export error: {e}")
        raise HTTPException(status_code=500, detail=f"Export failed: {str(e)}")

@app.get("/api/health")
def health_check():
    """Health check endpoint"""
    return {
        "status": "ok",
        "message": "Accessories Sales Dashboard API is running",
        "timestamp": datetime.now().isoformat()
    }

if __name__ == "__main__":
    import uvicorn
    print("\n" + "="*60)
    print("Starting Accessories Sales Dashboard")
    print("="*60)
    print("Dashboard: http://localhost:8000")
    print("API Docs: http://localhost:8000/docs")
    print("="*60 + "\n")
    
    uvicorn.run(app, host="0.0.0.0", port=8000)
