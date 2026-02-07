from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

import pandas as pd
import re
import time

# -------------------------------------------------
# CONFIG
# -------------------------------------------------
INVOICE_LINKS = [
    "https://apps.ristaapps.com/be/inventory?_o=y&hideBom=false&hideRates=false&id=1589651c-7e10-4f0e-bc8b-7e0fd00feeaa&uuid=922a2bb1-841c-4ef1-9a76-e3460f329898&_sign=ddtHruD1DR9MmMEp2rTeeAEwuztFqMKpUdARKuIIyKc%3D"
]

OUTPUT_FILE = "rista_inventory_consolidated.xlsx"

# -------------------------------------------------
# SETUP CHROME
# -------------------------------------------------
def get_driver():
    options = Options()
    options.add_argument("--start-maximized")
    # options.add_argument("--headless")  # enable later if needed

    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)

# -------------------------------------------------
# PARSE INVOICE PAGE
# -------------------------------------------------
def parse_invoice(driver, url):
    driver.get(url)

    # ✅ Wait for page JS to finish loading
    WebDriverWait(driver, 30).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

    # Extra wait for React rendering
    time.sleep(5)

    body_text = driver.find_element(By.TAG_NAME, "body").text
    lines = [l.strip() for l in body_text.split("\n") if l.strip()]

    items = []
    current_category = None

    for line in lines:

        # ---------------------------
        # CATEGORY HEADER
        # ---------------------------
        if (
            not line[0].isdigit()
            and "₹" not in line
            and not line.startswith("SKU:")
            and len(line.split()) <= 4
        ):
            current_category = line
            continue

        # ---------------------------
        # ITEM ROW
        # ---------------------------
        if re.match(r"^\d+\.", line):
            try:
                parts = re.split(r"\s{2,}", line)

                item_name = parts[0].split(".", 1)[1].strip()
                qty_raw = parts[1]
                amount_raw = parts[-1]

                qty_match = re.search(r"[\d.]+", qty_raw)
                quantity = float(qty_match.group()) if qty_match else 0

                amount = float(
                    amount_raw.replace("₹", "")
                    .replace(",", "")
                    .strip()
                )

                items.append({
                    "Category": current_category or "Unknown",
                    "Item": item_name,
                    "Quantity": quantity,
                    "Amount": amount
                })

            except Exception:
                continue

    return items

# -------------------------------------------------
# MAIN
# -------------------------------------------------
def main():
    driver = get_driver()
    all_items = []

    try:
        for link in INVOICE_LINKS:
            print(f"Processing: {link}")
            items = parse_invoice(driver, link)
            all_items.extend(items)

    finally:
        driver.quit()

    if not all_items:
        print("❌ No items extracted even via Selenium.")
        return

    df = pd.DataFrame(all_items)

    consolidated = (
        df.groupby(["Category", "Item"], as_index=False)
          .agg({"Quantity": "sum", "Amount": "sum"})
    )

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        consolidated.to_excel(writer, index=False, sheet_name="Consolidated")
        df.to_excel(writer, index=False, sheet_name="Raw Data")

    print(f"✅ Excel generated successfully: {OUTPUT_FILE}")

# -------------------------------------------------
# RUN
# -------------------------------------------------
if __name__ == "__main__":
    main()