export interface EppoProduct {
  id: string;
  name: string;
  ex_refinery: number;
  discount: number;
  excise_tax: number;
  m_tax: number;
  oil_fund: number;
  cons_fund: number;
  wholesale: number;
  vat_ws: number;
  ws_vat: number;
  mkt_margin: number;
  vat_mm: number;
  vat: number;
  retail: number;
}

export interface EppoDayRecord {
  date: string;
  exchange_rate: number | null;
  note: string;
  source_url: string;
  version: number;
  products: EppoProduct[];
}

export interface RetailItem {
  id: string;
  name: string;
  ptt: number | null;
  bcp: number | null;
  caltex: number | null;
  eppo: number | null;
  delta: number;
}

export interface RetailLatest {
  effective_date: string;
  effective_time: string;
  note: string;
  brands: string[];
  items: RetailItem[];
}

export interface YearsIndex {
  years: string[];
  min_date: string;
  max_date: string;
}

export interface SyncManifest {
  last_sync_utc: string;
  updated_records: number;
  total_days: number;
  min_date: string;
  max_date: string;
}
