"""
Headless ETL Pipeline for Thailand Oil Price Structure (EPPO) & Daily Retail Monitor.
Supports both:
  1. Full historical backfill across all EPPO pagination pages (with checkpointing).
  2. Fast incremental daily sync (stops early once caught up with inventory cache).
"""

from __future__ import annotations

import datetime as dt
import io
import json
import os
import re
import time
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any

import openpyxl
import requests
from bs4 import BeautifulSoup

ROOT = Path(__file__).resolve().parents[2]
DATA_DIR = ROOT / "data"
CACHE_DIR = DATA_DIR / "cache" / "eppo"
INVENTORY_PATH = DATA_DIR / "eppo_inventory_cache.json"
PUBLIC_DATA_DIR = ROOT / "frontend" / "public" / "data"
PUBLIC_EPPO_DIR = PUBLIC_DATA_DIR / "eppo"

EPPO_BASE_URL = (
    "https://www.eppo.go.th/data-energy-statistic/energy-price-th/"
    "%e0%b9%82%e0%b8%84%e0%b8%a3%e0%b8%87%e0%b8%aa%e0%b8%a3%e0%b9%89%e0%b8%b2%e0%b8%87"
    "%e0%b8%a3%e0%b8%b2%e0%b8%84%e0%b8%b2%e0%b8%99%e0%b9%89%e0%b8%b3%e0%b8%a1%e0%b8%b1%e0%b8%99-lpg/"
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
    ),
    "Referer": "https://www.eppo.go.th/",
    "Accept-Language": "th-TH,th;q=0.9,en;q=0.8",
}

FUEL_CODE_MAP = {
    "GASOHOL95 E20": "E20",
    "GASOHOL E20": "E20",
    "GASOHOL95 E85": "E85",
    "GASOHOL E85": "E85",
    "GASOHOL95 E10": "G95",
    "GASOHOL 95": "G95",
    "GASOHOL95": "G95",
    "GASOHOL91": "G91",
    "GASOHOL 91": "G91",
    "ULG95": "ULG95",
    "ULG 95": "ULG95",
    "ULG": "ULG95",
    "H-DIESEL B20": "B20",
    "H-DIESEL 20": "B20",
    "H-DIESEL B10": "B10",
    "H-DIESEL B7": "B7",
    "H-DIESEL": "HDIESEL",
    "DIESEL": "HDIESEL",
    "FO 600": "FO600",
    "FO 1500": "FO1500",
    "LPG": "LPG",
}


def map_fuel_id(raw_name: str) -> str:
    upper = " ".join(str(raw_name).upper().split())
    for key, code in FUEL_CODE_MAP.items():
        if upper.startswith(key):
            return code
    return re.sub(r"[^A-Z0-9]+", "_", upper).strip("_")


def to_float(val: Any) -> float:
    if val is None or val == "":
        return 0.0
    try:
        return round(float(val), 4)
    except (TypeError, ValueError):
        return 0.0


def extract_links_from_html(html_text: str) -> list[dict[str, Any]]:
    soup = BeautifulSoup(html_text, "html.parser")
    items: list[dict[str, Any]] = []
    seen_urls: set[str] = set()

    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        m = re.search(r"pt-price-st-(\d{4})-(\d{1,2})-(\d{1,2})(?:-(\d+))?\.xlsx?$", href, re.I)
        if not m:
            continue
        if href in seen_urls:
            continue
        seen_urls.add(href)

        date_str = f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
        version = int(m.group(4)) if m.group(4) else 0
        title = " ".join(a.get_text(" ", strip=True).split())
        parent_text = " ".join(a.parent.parent.get_text(" ", strip=True).split()) if a.parent and a.parent.parent else title
        note = "กบน. ปรับอัตราเงินกองทุนน้ำมัน" if "กบน" in parent_text else ""

        items.append({
            "date": date_str,
            "version": version,
            "url": href,
            "title": title,
            "note": note,
        })
    return items


def _iter_sheet_rows(file_bytes: bytes, url: str) -> list[list[Any]]:
    """Read top 45 rows x 20 cols of first sheet, supporting both .xlsx (openpyxl) and legacy .xls (xlrd)."""
    if not url.lower().endswith(".xls"):
        try:
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
            ws = wb[wb.sheetnames[0]]
            rows: list[list[Any]] = []
            for r_idx, row in enumerate(ws.iter_rows(max_row=45, max_col=20, values_only=True)):
                rows.append(list(row))
            wb.close()
            return rows
        except Exception:
            pass

    import xlrd
    book = xlrd.open_workbook(file_contents=file_bytes)
    sheet = book.sheet_by_index(0)
    rows = []
    for r in range(min(sheet.nrows, 45)):
        row_vals: list[Any] = []
        for c in range(min(sheet.ncols, 20)):
            cell = sheet.cell(r, c)
            if cell.ctype == xlrd.XL_CELL_DATE:
                row_vals.append(xlrd.xldate_as_datetime(cell.value, book.datemode))
            else:
                row_vals.append(cell.value if cell.value != "" else None)
        rows.append(row_vals)
    return rows


def parse_eppo_xlsx(xlsx_bytes: bytes, fallback_date: str, url: str = "") -> dict[str, Any]:
    """Parse Sheet 1 ('Oil Price Structure') of an EPPO .xlsx or .xls file."""
    grid = _iter_sheet_rows(xlsx_bytes, url)

    header_row = None
    col_map: dict[str, int] = {}
    doc_date = fallback_date
    exchange_rate = None

    for r_idx, row_vals in enumerate(grid[:30]):
        for val in row_vals:
            if isinstance(val, dt.datetime) and val.year >= 2010:
                doc_date = val.strftime("%Y-%m-%d")
            s = str(val or "").upper().strip()
            if "EX-REFIN" in s:
                header_row = r_idx
        if header_row == r_idx:
            for c_idx, val in enumerate(row_vals):
                s = " ".join(str(val or "").upper().split())
                if "EX-REFIN" in s:
                    col_map["ex_refinery"] = c_idx
                elif "DISCOUNT" in s:
                    col_map["discount"] = c_idx
                elif "EXCISE" in s:
                    col_map["excise_tax"] = c_idx
                elif "M. TAX" in s or "M.TAX" in s:
                    col_map["m_tax"] = c_idx
                elif "OIL FUND" in s:
                    col_map["oil_fund"] = c_idx
                elif "CONSV" in s:
                    col_map["cons_fund"] = c_idx
                elif "WHOLESALE" in s:
                    col_map["wholesale"] = c_idx
                elif "VAT (WS)" in s:
                    col_map["vat_ws"] = c_idx
                elif "WS&VAT" in s or "WS & VAT" in s:
                    col_map["ws_vat"] = c_idx
                elif "MARKETING" in s:
                    col_map["mkt_margin"] = c_idx
                elif "VAT (MM)" in s:
                    col_map["vat_mm"] = c_idx
                elif "RETAIL" in s:
                    col_map["retail"] = c_idx
            break

    if header_row is None:
        raise ValueError("Could not locate EX-REFIN header row in EPPO workbook")

    name_col = max(0, col_map.get("ex_refinery", 1) - 1)
    products: list[dict[str, Any]] = []
    for row_vals in grid[header_row + 1 :]:
        raw_name = row_vals[name_col] if name_col < len(row_vals) else None
        if raw_name is None:
            continue
        name_str = " ".join(str(raw_name).split())
        if not name_str:
            continue
        if "EXCHANGE RATE" in name_str.upper():
            for v in row_vals[name_col + 1 : name_col + 8]:
                if isinstance(v, (int, float)):
                    exchange_rate = round(float(v), 4)
                    break
            break
        if name_str.upper().startswith(("UNIT:", "หมายเหตุ", "REMARK", "*")):
            break

        def get_col(key: str) -> float:
            idx = col_map.get(key)
            if idx is None or idx >= len(row_vals):
                return 0.0
            return to_float(row_vals[idx])

        ex_ref = get_col("ex_refinery")
        discount = get_col("discount")
        excise = get_col("excise_tax")
        m_tax = get_col("m_tax")
        oil_fund = get_col("oil_fund")
        cons_fund = get_col("cons_fund")
        ws_price = get_col("wholesale")
        vat_ws = get_col("vat_ws")
        ws_vat = get_col("ws_vat")
        mm = get_col("mkt_margin")
        vat_mm = get_col("vat_mm")
        retail = get_col("retail")

        effective_retail = retail if retail > 0 else ws_vat

        products.append({
            "id": map_fuel_id(name_str),
            "name": name_str,
            "ex_refinery": round(ex_ref, 2),
            "discount": round(discount, 2),
            "excise_tax": round(excise, 3),
            "m_tax": round(m_tax, 3),
            "oil_fund": round(oil_fund, 2),
            "cons_fund": round(cons_fund, 2),
            "wholesale": round(ws_price, 2),
            "vat_ws": round(vat_ws, 2),
            "ws_vat": round(ws_vat, 2),
            "mkt_margin": round(mm, 2),
            "vat_mm": round(vat_mm, 2),
            "vat": round(vat_ws + vat_mm, 2),
            "retail": round(effective_retail, 2),
        })

    return {
        "date": doc_date,
        "exchange_rate": exchange_rate,
        "products": products,
    }


def load_json(path: Path, default: Any) -> Any:
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return default


def save_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def fetch_pttor_retail() -> tuple[str, str, dict[str, float]]:
    """Fetch official daily retail prices from PTTOR OilPrice.asmx (works 7 days/week)."""
    soap = """<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body><CurrentOilPrice xmlns="http://www.pttor.com"><Language>TH</Language></CurrentOilPrice></soap:Body>
</soap:Envelope>"""
    r = requests.post(
        "https://orapiweb.pttor.com/oilservice/OilPrice.asmx",
        data=soap.encode("utf-8"),
        headers={
            "Content-Type": "text/xml; charset=utf-8",
            "SOAPAction": '"http://www.pttor.com/CurrentOilPrice"',
            "User-Agent": HEADERS["User-Agent"],
        },
        timeout=15,
    )
    r.raise_for_status()
    root = ET.fromstring(r.text)
    res_node = root.find(".//{http://www.pttor.com}CurrentOilPriceResult")
    if res_node is None or not res_node.text:
        return "", "05:00", {}
    ds = ET.fromstring(res_node.text)

    pttor_map = {
        "เบนซินแก๊สโซฮอล์ 95": "G95",
        "เบนซินแก๊สโซฮอล์ 91": "G91",
        "เบนซินแก๊สโซฮอล์ E20": "E20",
        "เบนซินแก๊สโซฮอล์ E85": "E85",
        "เบนซิน": "ULG95",
        "ดีเซล": "HDIESEL",
        "ดีเซล B20": "B20",
        "Super Power GSH95": "PG95",
        "Super Power Diesel": "PDIESEL",
        "Super Power X99": "PX99",
    }
    prices: dict[str, float] = {}
    eff_date = ""
    eff_time = "05:00"
    for f in ds.findall("FUEL"):
        prod = (f.findtext("PRODUCT") or "").strip()
        price_str = f.findtext("PRICE")
        p_dt = (f.findtext("PRICE_DATE") or "").strip()
        if p_dt and not eff_date:
            eff_date = p_dt[:10]
            if "T" in p_dt:
                eff_time = p_dt.split("T")[1][:5]
        code = pttor_map.get(prod)
        if code and price_str:
            try:
                prices[code] = round(float(price_str), 2)
            except ValueError:
                pass
    return eff_date, eff_time, prices


def fetch_bangchak_retail() -> tuple[str, str, dict[str, float]]:
    """Fetch official daily retail prices from Bangchak's ApiOilPrice2 JSON widget."""
    r = requests.get(
        "https://oil-price.bangchak.co.th/ApiOilPrice2/en",
        headers={"User-Agent": HEADERS["User-Agent"]},
        timeout=15,
    )
    r.raise_for_status()
    payload = r.json()
    if not payload:
        return "", "05:00", {}
    record = payload[0]
    eff_date = ""
    raw_date = (record.get("OilPriceDate") or "").strip()
    if re.match(r"^\d{2}/\d{2}/\d{4}$", raw_date):
        dd, mm, yyyy = raw_date.split("/")
        eff_date = f"{yyyy}-{mm}-{dd}"
    eff_time = (record.get("OilPriceTime") or "05:00").strip()[:5]

    bcp_map = {
        "Gasohol 95 S EVO": "G95",
        "Gasohol 91 S EVO": "G91",
        "Gasohol E20 S EVO": "E20",
        "Gasohol E85 S EVO": "E85",
        "Hi Diesel S": "HDIESEL",
        "DIESEL B20": "B20",
        "Hi Premium 98 Plus": "P98",
        "Hi Premium Diesel Plus": "PDIESEL",
    }
    oils = json.loads(record.get("OilList") or "[]")
    prices: dict[str, float] = {}
    for item in oils:
        code = bcp_map.get((item.get("OilName") or "").strip())
        if code and item.get("PriceToday") is not None:
            try:
                prices[code] = round(float(item["PriceToday"]), 2)
            except (TypeError, ValueError):
                pass
    return eff_date, eff_time, prices


def fetch_caltex_retail() -> dict[str, float]:
    """Fetch official daily retail prices from Caltex Thailand's fuel-prices page."""
    r = requests.get(
        "https://www.caltex.com/th/motorists/products-and-services/fuel-prices.html",
        headers={"User-Agent": HEADERS["User-Agent"]},
        timeout=15,
    )
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    caltex_map = {
        "โกลด์ 95 เทครอน®": "ULG95",
        "แก๊สโซฮอล์ 95 เทครอน®": "G95",
        "แก๊สโซฮอล์ 91 เทครอน®": "G91",
        "แก๊สโซฮอล์ E20": "E20",
        "ดีเซล เทครอน® ดี": "HDIESEL",
        "ดีเซล B20": "B20",
        "พาวเวอร์ ดีเซล เทครอน® ดี": "PDIESEL",
    }
    prices: dict[str, float] = {}
    for card in soup.select(".price-item"):
        name_el = card.select_one(".first-col p")
        price_els = card.select(".col-xs-4.col-md-3 p")
        if not name_el or not price_els:
            continue
        code = caltex_map.get(name_el.get_text(strip=True))
        if not code:
            continue
        m = re.search(r"(\d+(?:\.\d+)?)", price_els[0].get_text(strip=True))
        if m:
            prices[code] = round(float(m.group(1)), 2)
    return prices


def flush_outputs(
    year_records: dict[str, dict[str, Any]],
    inventory: dict[str, dict[str, Any]],
    updated_count: int,
) -> None:
    """Persist per-year JSONs, manifest, years_index, and retail seed to data/ and frontend/public/data/."""
    all_dates: list[str] = []
    years_list: list[str] = sorted(year_records.keys())
    for yr in years_list:
        sorted_day_map = dict(sorted(year_records[yr].items()))
        all_dates.extend(sorted_day_map.keys())
        save_json(CACHE_DIR / f"{yr}.json", sorted_day_map)
        save_json(PUBLIC_EPPO_DIR / f"{yr}.json", sorted_day_map)

    all_dates.sort()
    if all_dates:
        latest_date = all_dates[-1]
        latest_payload = year_records[latest_date[:4]][latest_date]
        save_json(DATA_DIR / "eppo_latest.json", latest_payload)
        save_json(PUBLIC_DATA_DIR / "eppo_latest.json", latest_payload)

    save_json(INVENTORY_PATH, inventory)
    idx_payload = {
        "years": years_list,
        "min_date": all_dates[0] if all_dates else "",
        "max_date": all_dates[-1] if all_dates else "",
    }
    save_json(DATA_DIR / "years_index.json", idx_payload)
    save_json(PUBLIC_DATA_DIR / "years_index.json", idx_payload)

    manifest = {
        "last_sync_utc": dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "updated_records": updated_count,
        "total_days": len(all_dates),
        "min_date": all_dates[0] if all_dates else "",
        "max_date": all_dates[-1] if all_dates else "",
    }
    save_json(DATA_DIR / "sync_manifest.json", manifest)
    save_json(PUBLIC_DATA_DIR / "sync_manifest.json", manifest)

    retail_public = PUBLIC_DATA_DIR / "retail_latest.json"
    if all_dates:
        latest_products = year_records[all_dates[-1][:4]][all_dates[-1]]["products"]
        ptt_date, ptt_time, ptt_prices = "", "05:00", {}
        try:
            ptt_date, ptt_time, ptt_prices = fetch_pttor_retail()
        except Exception as exc:
            print(f"[WARN] PTTOR retail fetch failed: {exc}")

        bcp_date, bcp_time, bcp_prices = "", "05:00", {}
        try:
            bcp_date, bcp_time, bcp_prices = fetch_bangchak_retail()
        except Exception as exc:
            print(f"[WARN] Bangchak retail fetch failed: {exc}")

        caltex_prices: dict[str, float] = {}
        try:
            caltex_prices = fetch_caltex_retail()
        except Exception as exc:
            print(f"[WARN] Caltex retail fetch failed: {exc}")

        items = []
        for p in latest_products:
            pid = p["id"]
            is_station_fuel = not pid.startswith("FO") and pid != "LPG"
            ptt_val = ptt_prices.get(pid, p["retail"] if is_station_fuel else None) if ptt_prices or is_station_fuel else None
            bcp_val = (
                bcp_prices.get(pid)
                if bcp_prices
                else (p["retail"] if is_station_fuel and pid != "ULG95" else None)
            )
            caltex_val = (
                caltex_prices.get(pid)
                if caltex_prices
                else (p["retail"] if is_station_fuel and pid not in ("E85", "B20") else None)
            )
            delta_val = round(ptt_val - p["retail"], 2) if (ptt_val is not None and p.get("retail") is not None) else 0.0
            items.append({
                "id": pid,
                "name": p["name"],
                "ptt": ptt_val,
                "bcp": bcp_val,
                "caltex": caltex_val,
                "eppo": p["retail"],
                "delta": delta_val,
            })

        for prem_id, prem_name in [
            ("PG95", "Super Power GSH95 (Premium)"),
            ("P98", "Hi Premium 98 Plus (Bangchak)"),
            ("PDIESEL", "Premium Diesel (Super Power / Hi Premium / Techron)"),
            ("PX99", "Super Power X99 (Premium)"),
        ]:
            ptt_p = ptt_prices.get(prem_id)
            bcp_p = bcp_prices.get(prem_id)
            caltex_p = caltex_prices.get(prem_id)
            if ptt_p is not None or bcp_p is not None or caltex_p is not None:
                items.append({
                    "id": prem_id,
                    "name": prem_name,
                    "ptt": ptt_p,
                    "bcp": bcp_p,
                    "caltex": caltex_p,
                    "eppo": None,
                    "delta": 0.0,
                })

        live_sources = []
        if ptt_prices:
            live_sources.append("PTTOR (orapiweb)")
        if bcp_prices:
            live_sources.append("Bangchak (ApiOilPrice2)")
        if caltex_prices:
            live_sources.append("Caltex (caltex.com)")
        fallback_brands = [b for b, ok in [("PTTOR", ptt_prices), ("Bangchak", bcp_prices), ("Caltex", caltex_prices)] if not ok]
        note_parts = []
        if live_sources:
            note_parts.append("ราคาจริงรายวัน: " + " • ".join(live_sources))
        if fallback_brands:
            note_parts.append(f"อ้างอิง EPPO: {'/'.join(fallback_brands)}")

        retail_seed = {
            "effective_date": ptt_date or bcp_date or all_dates[-1],
            "effective_time": ptt_time if ptt_date else (bcp_time if bcp_date else "05:00"),
            "note": " | ".join(note_parts) if note_parts else "อ้างอิงจากโครงสร้างราคา EPPO",
            "brands": ["PTTOR", "Bangchak", "Caltex"],
            "items": items,
        }
        save_json(DATA_DIR / "retail_latest.json", retail_seed)
        save_json(retail_public, retail_seed)


def main() -> None:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    PUBLIC_EPPO_DIR.mkdir(parents=True, exist_ok=True)

    force_refresh = os.getenv("FORCE_REFRESH", "false").lower() == "true"
    backfill_all = os.getenv("BACKFILL_ALL", "false").lower() == "true"
    max_pages = 0 if backfill_all else int(os.getenv("MAX_PAGES", "2"))
    min_year = int(os.getenv("MIN_YEAR", "2017"))

    inventory: dict[str, dict[str, Any]] = load_json(INVENTORY_PATH, {})
    year_records: dict[str, dict[str, Any]] = {}
    for p in CACHE_DIR.glob("*.json"):
        year_records[p.stem] = load_json(p, {})

    session = requests.Session()
    updated_count = 0
    error_count = 0
    page = 1
    below_min_streak = 0

    print(f"[INFO] Starting EPPO sync (backfill_all={backfill_all}, min_year={min_year}, max_pages={max_pages or 'ALL'}, force_refresh={force_refresh})")

    def download_and_parse(item: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
        f_resp = requests.get(item["url"], headers=HEADERS, timeout=25)
        f_resp.raise_for_status()
        parsed = parse_eppo_xlsx(f_resp.content, fallback_date=item["date"], url=item["url"])
        parsed["note"] = item["note"]
        parsed["source_url"] = item["url"]
        parsed["version"] = item["version"]
        return item, parsed

    while True:
        if max_pages > 0 and page > max_pages:
            break

        page_url = EPPO_BASE_URL if page == 1 else f"{EPPO_BASE_URL}page/{page}/"
        try:
            resp = session.get(page_url, headers=HEADERS, timeout=25)
        except requests.RequestException as exc:
            print(f"[WARN] Network error on page {page}: {exc}")
            break

        if resp.status_code != 200:
            print(f"[INFO] Reached end of pagination at page {page} (HTTP {resp.status_code})")
            break

        links = extract_links_from_html(resp.text)
        if not links:
            print(f"[INFO] No .xlsx links found on page {page}. Stopping.")
            break

        in_range_links = [it for it in links if it["date"] and int(it["date"][:4]) >= min_year]
        if not in_range_links:
            below_min_streak += 1
            if below_min_streak >= 5:
                print(f"[INFO] Reached {below_min_streak} consecutive pages below min_year={min_year}. Stopping backfill.")
                break
            page += 1
            continue
        below_min_streak = 0

        to_fetch = [
            it for it in in_range_links
            if force_refresh or not (inventory.get(it["date"]) and inventory[it["date"]].get("url") == it["url"])
        ]

        page_new = 0
        if to_fetch:
            with ThreadPoolExecutor(max_workers=8) as pool:
                futures = {pool.submit(download_and_parse, it): it for it in to_fetch}
                for fut in as_completed(futures):
                    it = futures[fut]
                    date_str = it["date"]
                    year = date_str[:4]
                    try:
                        _, parsed = fut.result()
                        year_records.setdefault(year, {})[date_str] = parsed
                        inventory[date_str] = {"url": it["url"], "version": it["version"]}
                        updated_count += 1
                        page_new += 1
                    except Exception as exc:
                        error_count += 1
                        print(f"[WARN] Failed {date_str} ({it['url']}): {exc}")
            print(f"[PAGE {page}] Synced {page_new} files (total_days={len(inventory)})")

        # Early-stop in incremental mode once a full page has no new/updated files
        if not backfill_all and not force_refresh and page_new == 0:
            print(f"[INFO] Page {page} already up to date. Stopping incremental scan early.")
            break

        # Checkpoint every 10 pages during large backfills
        if page % 10 == 0 and updated_count > 0:
            flush_outputs(year_records, inventory, updated_count)
            print(f"[CHECKPOINT] Saved progress at page {page} ({len(inventory)} total days)")

        page += 1

    flush_outputs(year_records, inventory, updated_count)
    print(f"[DONE] Synced {updated_count} new/updated days (errors={error_count}). Total days in DB: {len(inventory)}")


if __name__ == "__main__":
    main()
