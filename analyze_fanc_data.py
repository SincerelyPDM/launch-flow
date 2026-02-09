
import pandas as pd
import openpyxl
import os
import re

def load_data():
    base_path = "/Users/boyoungkim/Desktop/launch-flow"
    
    # Paths
    parts_path = os.path.join(base_path, "references/parts_260208_v2.csv")
    sales_path = os.path.join(base_path, "references/sales_performance_290101_20251231.csv")
    crawler_path = os.path.join(base_path, "flows/FANC/web-crawler_koreagift_products_20251119_201638.csv")
    quotation_path = os.path.join(base_path, "flows/FANC/FANC_Quotation_v1.xlsx")

    # Load Data with specific options if needed
    try:
        parts_df = pd.read_csv(parts_path)
    except Exception as e:
        print(f"Error loading parts: {e}")
        parts_df = pd.DataFrame()

    try:
        sales_df = pd.read_csv(sales_path)
    except Exception as e:
        print(f"Error loading sales: {e}")
        sales_df = pd.DataFrame()

    try:
        crawler_df = pd.read_csv(crawler_path)
    except Exception as e:
        print(f"Error loading crawler: {e}")
        crawler_df = pd.DataFrame()
        
    quotation_data = []
    try:
        wb = openpyxl.load_workbook(quotation_path, data_only=True)
        sheet = wb.active
        print("\n--- DEBUG: Raw Quotation Rows (First 10) ---")
        for i, row in enumerate(sheet.iter_rows(values_only=True)):
            row_data = [str(x) if x is not None else "" for x in row]
            quotation_data.append(row_data)
            if i < 10:
                print(row_data)
    except Exception as e:
        print(f"Error loading quotation: {e}")

    return parts_df, sales_df, crawler_df, quotation_data

def extract_product_specs(quotation_data):
    print("\n--- Product Specs Extraction ---")
    specs = {
        "Product Name": "N/A",  # Default values
        "Dimensions": "N/A",
        "Weight": "N/A",
        "Material": "N/A",
        "Price": "N/A",
        "Spec": "N/A"
    }
    
    # Based on debug output:
    # Header: ['NO.', '슈틸루스터 모델명', '현재 제품명', '제품명', ... '제품사이즈', '비고', '공급가', '주요 스팩', '']
    # Index:   0      1                2             3          5            7         8
    # Row 6 is header, Row 7 is first data (index 6 in 0-indexed list if we included header, but logic iterates)
    
    found_header = False
    for row in quotation_data:
        # Check if this is the header row
        if len(row) > 3 and "제품명" in row[3] and "공급가" in row[7]:
            found_header = True
            continue # Skip header
        
        if found_header:
            # First data row after header
            # clean data
            specs["Product Name"] = row[3]
            specs["Dimensions"] = row[5].replace("×", "x") # Normalize
            specs["Price"] = row[7]
            specs["Spec"] = row[8]
            # Dimensions usually L x W x H. 
            # 41x137x51 -> need to parse carefully in matching
            break 
            
    print("Extracted Specs:", specs)
    return specs

def analyze_sales_and_market(sales_df, crawler_df):
    print("\n--- Sales & Market Analysis ---")
    
    # Sales
    # Filter for synonyms of "Fan"
    fan_keywords = ['Fan', '선풍기', '팬', '써큘레이터']
    pattern = '|'.join(fan_keywords)
    
    # Columns from debug: '품목', '수주액 2023', '판매량 2023', etc.
    prod_col = '품목'
    
    sales_summary = "No Sales Data Found"

    if prod_col in sales_df.columns and not sales_df.empty:
        fan_sales = sales_df[sales_df[prod_col].astype(str).str.contains(pattern, case=False, na=False)]
        
        if not fan_sales.empty:
            count = len(fan_sales)
            
            # Calculate totals for 3 years (2022, 2023, 2024)
            # Columns like '수주액 2022', '수주액 2023', '수주액 2024'
            years = ['2022', '2023', '2024']
            total_qty = 0
            total_rev = 0
            yearly_breakdown = []
            
            for y in years:
                rev_col = f"수주액 {y}"
                qty_col = f"판매량 {y}"
                
                y_rev = 0
                y_qty = 0
                
                if rev_col in fan_sales.columns:
                    # Clean commas if string
                    if fan_sales[rev_col].dtype == object:
                         y_rev = pd.to_numeric(fan_sales[rev_col].astype(str).str.replace(',', ''), errors='coerce').sum()
                    else:
                         y_rev = fan_sales[rev_col].sum()
                
                if qty_col in fan_sales.columns:
                     if fan_sales[qty_col].dtype == object:
                         y_qty = pd.to_numeric(fan_sales[qty_col].astype(str).str.replace(',', ''), errors='coerce').sum()
                     else:
                         y_qty = fan_sales[qty_col].sum()
                
                total_rev += y_rev
                total_qty += y_qty
                yearly_breakdown.append(f"{y}: {y_qty:,.0f} units, {y_rev:,.0f} KRW")

            sales_summary = f"Total Sales (2022-2024): {count} Products Triggered.\nTotal Vol: {total_qty:,.0f}, Total Rev: {total_rev:,.0f} KRW"
            sales_summary += "\nYearly Breakdown:\n" + "\n".join(yearly_breakdown)
            
            # Also valid sales in 2025?
            y25_rev = 0
            if '수주액 2025' in fan_sales.columns:
                 y25_rev = pd.to_numeric(fan_sales['수주액 2025'].astype(str).str.replace(',',''), errors='coerce').sum()
                 sales_summary += f"\n2025 (YTD): {y25_rev:,.0f} KRW"
                 
        else:
            sales_summary = "No sales records found for 'Fan' keywords."
    
    print(sales_summary)

    # Market (Crawler)
    market_summary = "No Competitor Data Found"
    if not crawler_df.empty:
        price_col = next((c for c in crawler_df.columns if '가격' in c or 'Price' in c), None)
        if price_col:
            # Clean price data (remove commas, '원', etc)
            crawler_df['CleanPrice'] = crawler_df[price_col].astype(str).str.replace(r'[^\d]', '', regex=True)
            crawler_df['CleanPrice'] = pd.to_numeric(crawler_df['CleanPrice'], errors='coerce')
            
            stats = crawler_df['CleanPrice'].describe()
            market_summary = f"Competitor Pricing (n={len(crawler_df)}):\nAvg: {stats['mean']:,.0f}\nMin: {stats['min']:,.0f}\nMax: {stats['max']:,.0f}"
    
    print(market_summary)
    return sales_summary, market_summary

def parse_dimensions(dim_str):
    # Extracts L, W, H from string like "100x200x30" or "100*200*30"
    if not isinstance(dim_str, str): return None
    nums = re.findall(r'\d+', dim_str)
    if len(nums) >= 3:
        return [int(x) for x in nums[:3]] # [L, W, H]
    return None

def package_matching(parts_df, product_dims_str):
    print("\n--- Package Matching ---")
    
    product_dims = parse_dimensions(product_dims_str)
    if not product_dims:
        print(f"Could not parse product dimensions from: {product_dims_str}")
        return

    print(f"Product Dimensions (L x W x H): {product_dims}")
    
    # Filter for Package Parts
    # Look for '파츠 유형' == 'Package Parts' or similar
    type_col = next((c for c in parts_df.columns if '유형' in c or 'Type' in c), None)
    name_col = next((c for c in parts_df.columns if '파츠명' in c or 'Name' in c), None)
    # Correct column name is '파츠: 가로x세로x높이(장폭고)'
    size_col = next((c for c in parts_df.columns if '규격' in c or 'Size' in c or '가로' in c or '장폭고' in c), None)
    
    if not (type_col and size_col):
        print("Required columns (Type, Size) not found in Parts DB.")
        return

    packages = parts_df[parts_df[type_col].astype(str).str.contains('Package', case=False, na=False)].copy()
    
    matches = []
    
    for _, pkg in packages.iterrows():
        pkg_dims = parse_dimensions(str(pkg[size_col]))
        if pkg_dims:
            # Calculate Gap (Fit Score)
            # Assuming simple comparison: Package Inner > Product
            # Need to match orientation? Sort dims to be safe? 
            # Usually L>=L, W>=W, H>=H. Let's sort both for "best fit" check regardless of orientation
            p_sorted = sorted(product_dims)
            pkg_sorted = sorted(pkg_dims)
            
            diffs = [pkg_sorted[i] - p_sorted[i] for i in range(3)]
            
            # Check if it fits (all diffs >= 0)
            if all(d >= 0 for d in diffs):
                gap_score = sum(diffs) # Simple heuristic
                match_info = {
                    'Name': pkg[name_col],
                    'Size': pkg[size_col],
                    'Gap': diffs, # [L_gap, W_gap, H_gap]
                    'Total_Gap': gap_score,
                    'Verdict': 'Optimal' if 5 <= min(diffs) and max(diffs) <= 15 else 'Check' # Simple rule
                }
                matches.append(match_info)
    
    # Sort by Total Gap (tighter is usually better but not too tight)
    matches.sort(key=lambda x: x['Total_Gap'])
    
    print(f"Found {len(matches)} potential packages.")
    for m in matches[:5]:
        print(m)

def main():
    parts_df, sales_df, crawler_df, quotation_data = load_data()
    
    specs = extract_product_specs(quotation_data)
    analyze_sales_and_market(sales_df, crawler_df)
    
    # Run matching if we have dimensions
    if specs["Dimensions"] != "N/A":
        package_matching(parts_df, specs["Dimensions"])
    else:
        print("Skipping package matching due to missing dimensions.")
    
if __name__ == "__main__":
    main()
