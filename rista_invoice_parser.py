import requests
import pandas as pd
from bs4 import BeautifulSoup
import re
from pathlib import Path
from datetime import datetime

# -------------------------------------------------
# CONFIG
# -------------------------------------------------
LINKS_FILE = "invoice_links.txt"
SALES_FILE = "Channasandra - Multidate - Sales By Items.csv"
OUTPUT_FILE = "rista_inventory_consolidated.xlsx"
DUPLICATE_FILE = "duplicate_links.txt"

EXCLUDE_SALES_CATEGORIES = ["Cakes", "Eggless Cakes", "Menu Items"]

HEADERS = {"User-Agent": "Mozilla/5.0"}

# -------------------------------------------------
# UTILITIES
# -------------------------------------------------
def normalize(text):
    if pd.isna(text):
        return ""
    return str(text).strip().lower()

def find_column(df, keywords, required=True):
    for col in df.columns:
        col_l = col.lower()
        if any(k in col_l for k in keywords):
            return col
    if required:
        raise ValueError(f"Required column not found. Tried keywords: {keywords}")
    return None

# -------------------------------------------------
# READ LINKS + LOG DUPLICATES
# -------------------------------------------------
def read_unique_links(file_path):
    seen = {}
    unique_links = []
    duplicates = []

    with open(file_path, "r", encoding="utf-8") as f:
        for line_no, line in enumerate(f, start=1):
            link = line.split("#", 1)[0].strip()
            if not link:
                continue

            if link not in seen:
                seen[link] = line_no
                unique_links.append(link)
            else:
                duplicates.append((line_no, seen[link], link))

    if duplicates:
        with open(DUPLICATE_FILE, "w", encoding="utf-8") as f:
            f.write("Duplicate Invoice Links\n")
            f.write("-----------------------\n")
            for dup, first, link in duplicates:
                f.write(
                    f"Duplicate at line {dup}, first seen at line {first}: {link}\n"
                )

    return unique_links

# -------------------------------------------------
# FETCH HTML
# -------------------------------------------------
def fetch_html(url):
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    return r.text

# -------------------------------------------------
# PARSE INVOICE HTML
# -------------------------------------------------
def parse_invoice(html):
    soup = BeautifulSoup(html, "html.parser")
    items = []

    date_text = soup.find(string=re.compile("DATE & TIME"))
    invoice_date = None
    if date_text:
        m = re.search(r"\d{2}-\w{3}-\d{4}", date_text)
        if m:
            invoice_date = m.group(0)

    table = soup.find("table", attrs={"width": "100%"})
    rows = table.find_all("tr")

    current_category = None

    for row in rows:
        tds = row.find_all("td")

        # Category row
        if len(tds) == 1 and "bold" in tds[0].get("class", []):
            current_category = tds[0].get_text(strip=True)
            continue

        # Item row
        if len(tds) == 6 and re.match(r"\d+\.", tds[0].get_text(strip=True)):
            item = tds[0].get_text(strip=True).split(".", 1)[1].strip()
            qty = float(re.search(r"[\d.]+", tds[1].get_text()).group())
            amt = float(
                tds[5].get_text().replace("₹", "").replace(",", "").strip()
            )

            items.append({
                "Date": invoice_date,
                "Category": current_category,
                "Item": item,
                "Quantity": qty,
                "Amount": amt
            })

    return items

# -------------------------------------------------
# INVENTORY VS SALES COMPARISON
# -------------------------------------------------
def compare_inventory_vs_sales(consolidated_df):
    if not Path(SALES_FILE).exists():
        print("⚠️ Sales file not found, skipping comparison.")
        return None, None, None, None

    sales_df = (
        pd.read_csv(SALES_FILE)
        if SALES_FILE.lower().endswith(".csv")
        else pd.read_excel(SALES_FILE)
    )

    sales_df.columns = [c.strip() for c in sales_df.columns]

    # Detect columns
    item_col = find_column(sales_df, ["item"])
    qty_col = find_column(sales_df, ["qty", "quantity"])
    amount_col = find_column(sales_df, ["net", "amount", "sales"])
    category_col = find_column(sales_df, ["category", "group"], required=False)

    # Apply category exclusion ONLY if column exists
    if category_col:
        sales_df = sales_df[
            ~sales_df[category_col].isin(EXCLUDE_SALES_CATEGORIES)
        ]

    # Normalize item names
    sales_df["Item_N"] = sales_df[item_col].apply(normalize)
    consolidated_df["Item_N"] = consolidated_df["Item"].apply(normalize)

    # Aggregate sales
    sales_sum = (
        sales_df.groupby("Item_N", as_index=False)
        .agg({
            qty_col: "sum",
            amount_col: "sum"
        })
        .rename(columns={
            qty_col: "Sold Qty",
            amount_col: "Net Sales"
        })
    )

    # Aggregate inventory
    inv_sum = (
        consolidated_df.groupby("Item_N", as_index=False)
        .agg({
            "Item": "first",
            "Quantity": "sum",
            "Amount": "sum"
        })
        .rename(columns={
            "Quantity": "Inventory Qty",
            "Amount": "Investment"
        })
    )

    # Compared
    compared = pd.merge(inv_sum, sales_sum, on="Item_N", how="inner")
    compared["Balance Qty"] = compared["Inventory Qty"] - compared["Sold Qty"]
    compared["Status"] = compared["Balance Qty"].apply(
        lambda x: "In Stock" if x > 0 else ("Sold Out" if x == 0 else "Lost")
    )

    compared = compared[[
        "Item", "Inventory Qty", "Sold Qty",
        "Balance Qty", "Investment", "Net Sales", "Status"
    ]]

    # Inventory not sold
    inventory_only = inv_sum[
        ~inv_sum["Item_N"].isin(sales_sum["Item_N"])
    ][["Item", "Inventory Qty", "Investment"]]

    # Sold but not in inventory
    sales_only = sales_sum[
        ~sales_sum["Item_N"].isin(inv_sum["Item_N"])
    ][["Item_N", "Sold Qty", "Net Sales"]].rename(
        columns={"Item_N": "Item"}
    )

    # Summary
    summary = pd.DataFrame({
        "Metric": [
            "Total Investment",
            "Total Sales",
            "Items In Stock",
            "Items Sold Out",
            "Items Lost"
        ],
        "Value": [
            compared["Investment"].sum(),
            compared["Net Sales"].sum(),
            (compared["Status"] == "In Stock").sum(),
            (compared["Status"] == "Sold Out").sum(),
            (compared["Status"] == "Lost").sum()
        ]
    })

    return compared, inventory_only, sales_only, summary

# -------------------------------------------------
# MAIN
# -------------------------------------------------
def main():
    links = read_unique_links(LINKS_FILE)

    all_items = []
    for link in links:
        html = fetch_html(link)
        all_items.extend(parse_invoice(html))

    if not all_items:
        print("❌ No invoice data extracted.")
        return

    df = pd.DataFrame(all_items)

    consolidated = (
        df.groupby(["Category", "Item"], as_index=False)
        .agg({"Quantity": "sum", "Amount": "sum"})
    )

    daywise = (
        df.groupby(["Date", "Category", "Item"], as_index=False)
        .agg({"Quantity": "sum"})
    )

    # Pivot with proper date sorting (Latest → Oldest)
    pivot = (
        df.pivot_table(
            index=["Category", "Item"],
            columns="Date",
            values="Quantity",
            aggfunc="sum",
            fill_value=0
        )
        .reset_index()
    )

    date_cols = [
        c for c in pivot.columns
        if isinstance(c, str) and re.match(r"\d{2}-\w{3}-\d{4}", c)
    ]

    date_cols_sorted = sorted(
        date_cols,
        key=lambda x: datetime.strptime(x, "%d-%b-%Y"),
        reverse=True
    )

    pivot = pivot[["Category", "Item"] + date_cols_sorted]
    pivot["Total"] = pivot[date_cols_sorted].sum(axis=1)

    compared, inv_only, sales_only, summary = compare_inventory_vs_sales(consolidated)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        consolidated.to_excel(writer, index=False, sheet_name="Consolidated")
        daywise.to_excel(writer, index=False, sheet_name="Day Wise Quantity")
        pivot.to_excel(writer, index=False, sheet_name="Pivot Day Wise Qty")
        df.to_excel(writer, index=False, sheet_name="Raw Data")

        if compared is not None:
            compared.to_excel(writer, index=False, sheet_name="Compared Items")
            inv_only.to_excel(writer, index=False, sheet_name="Inventory Not Sold")
            sales_only.to_excel(writer, index=False, sheet_name="Sold Not In Inventory")
            summary.to_excel(writer, index=False, sheet_name="Summary")

    print("✅ All reports generated successfully")

# -------------------------------------------------
if __name__ == "__main__":
    main()