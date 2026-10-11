"""
Data Parser for IIT / NIT / IIIT Excel Files.
Specifically designed for JOSAA / CSAB IIT, NIT, and IIIT datasets.
Handles sheet detection, category normalization, rank consolidation,
and NIT quota preservation.
"""

import os
import re
import pandas as pd
from metadata_db import get_institute_metadata

# Allowed Categories
ALLOWED_SEAT_TYPES = {
    "OPEN": "OPEN",
    "GEN": "OPEN",
    "GENERAL": "OPEN",
    "EWS": "EWS",
    "GEN-EWS": "EWS",
    "OBC": "OBC-NCL",
    "OBC-NCL": "OBC-NCL",
    "OBC - NCL": "OBC-NCL",
    "SC": "SC",
    "ST": "ST"
}

# Image 1 Category Sequence: OPEN-N, OPEN-W, OBC-N, OBC-W, EWS-N, EWS-W, SC-N, SC-W, ST-N, ST-W
CATEGORY_ORDER_IMAGE1 = [
    "OPEN-N", "OPEN-W",
    "OBC-N", "OBC-W",
    "EWS-N", "EWS-W",
    "SC-N", "SC-W",
    "ST-N", "ST-W"
]

# Section 6 Category Sequence (EWS before OBC)
CATEGORY_ORDER_SECTION6 = [
    "OPEN-N", "OPEN-W",
    "EWS-N", "EWS-W",
    "OBC-N", "OBC-W",
    "SC-N", "SC-W",
    "ST-N", "ST-W"
]

CATEGORY_LABELS = {
    "OPEN-N": ("OPEN", "Gender-Neutral"),
    "OPEN-W": ("OPEN", "Female-Only"),
    "EWS-N": ("EWS", "Gender-Neutral"),
    "EWS-W": ("EWS", "Female-Only"),
    "OBC-N": ("OBC-NCL", "Gender-Neutral"),
    "OBC-W": ("OBC-NCL", "Female-Only"),
    "SC-N": ("SC", "Gender-Neutral"),
    "SC-W": ("SC", "Female-Only"),
    "ST-N": ("ST", "Gender-Neutral"),
    "ST-W": ("ST", "Female-Only"),
}

def normalize_gender(gender_val):
    if not gender_val or pd.isna(gender_val):
        return None
    s = str(gender_val).strip().lower()
    if "female" in s or "-w" in s or s == "f" or s == "w":
        return "W"
    if "neutral" in s or "-n" in s or s == "n" or s == "m":
        return "N"
    return None

def normalize_category_and_gender(seat_type_val, gender_val=None):
    """
    Normalizes (Seat Type, Gender) into standard token:
    'OPEN-N', 'OPEN-W', 'OBC-N', 'OBC-W', 'EWS-N', 'EWS-W', 'SC-N', 'SC-W', 'ST-N', 'ST-W'.
    Returns (token, is_allowed, raw_name).
    """
    if seat_type_val is None or pd.isna(seat_type_val):
        return None, False, ""
        
    raw_st = str(seat_type_val).strip()
    
    # Check if combined in one string e.g. "OPEN-N" or "OPEN (Female)"
    match_comb = re.match(r"^([A-Z\-]+)[\s\-_]+([NWnw]|Female|Gender\-Neutral|Female\-only)$", raw_st, re.I)
    if match_comb:
        base_cat = match_comb.group(1).strip().upper()
        gdr_part = match_comb.group(2).strip()
        gdr_code = normalize_gender(gdr_part)
    else:
        base_cat = raw_st.upper()
        gdr_code = normalize_gender(gender_val)
        
    # Check if it has PwD or disallowed keywords
    if "PWD" in base_cat or "DEF" in base_cat or "FOREIGN" in base_cat or "OCI" in base_cat:
        return None, False, raw_st
        
    canonical_cat = ALLOWED_SEAT_TYPES.get(base_cat)
    if not canonical_cat:
        # Check without punctuation
        clean_cat = re.sub(r"[^A-Z]", "", base_cat)
        canonical_cat = ALLOWED_SEAT_TYPES.get(clean_cat)
        
    if not canonical_cat:
        return None, False, raw_st
        
    if not gdr_code:
        gdr_code = "N"  # Default if unspecified
        
    # Standardize name for token
    short_base = "OBC" if canonical_cat == "OBC-NCL" else canonical_cat
    token = f"{short_base}-{gdr_code}"
    return token, True, raw_st

def parse_rank_value(val):
    """Parses rank into integer or None."""
    if val is None or pd.isna(val):
        return None
    s = str(val).strip()
    if not s or s.lower() in ("-", "--", "---", "----", "-----", "nan", "null", "none"):
        return None
    # match numeric value
    m = re.search(r"(\d+)", s)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            return None
    return None

class JosaaExcelParser:
    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.xl = None
        self.sheet_map = {}
        self.detected_categories = set()
        self.ignored_categories = set()
        self.validation_summary = {}
        self.processed_data = {}
        
    def inspect(self):
        """Analyzes workbook sheets and structure before full processing."""
        if not os.path.exists(self.excel_path):
            raise FileNotFoundError(f"Excel file not found at: {self.excel_path}")
            
        self.xl = pd.ExcelFile(self.excel_path)
        all_sheets = self.xl.sheet_names
        
        # Map sheets
        for s in all_sheets:
            sl = s.lower().replace(" ", "").replace("_", "")
            if "iit" in sl and "iiit" not in sl:
                self.sheet_map["IIT"] = s
            elif "nit" in sl:
                self.sheet_map["NIT"] = s
            elif "iiit" in sl:
                self.sheet_map["IIIT"] = s
                
        summary = {
            "sheets_found": {
                "IIT": "IIT" in self.sheet_map,
                "NIT": "NIT" in self.sheet_map,
                "IIIT": "IIIT" in self.sheet_map,
            },
            "sheet_names": self.sheet_map,
            "records": {},
            "institutes": {},
            "programmes": {},
            "detected_categories": set(),
            "ignored_categories": set()
        }
        
        for kind in ["IIT", "NIT", "IIIT"]:
            if kind in self.sheet_map:
                s_name = self.sheet_map[kind]
                df = self.xl.parse(s_name)
                # drop completely blank rows and columns
                df = df.dropna(how="all", axis=0).dropna(how="all", axis=1)
                
                # identify columns
                inst_col = next((c for c in df.columns if "inst" in str(c).lower()), None)
                prog_col = next((c for c in df.columns if "program" in str(c).lower() or "branch" in str(c).lower()), None)
                st_col = next((c for c in df.columns if "seat" in str(c).lower() or "category" in str(c).lower()), None)
                gen_col = next((c for c in df.columns if "gender" in str(c).lower()), None)
                
                summary["records"][kind] = len(df)
                if inst_col:
                    summary["institutes"][kind] = df[inst_col].nunique()
                if prog_col:
                    summary["programmes"][kind] = df[prog_col].nunique()
                    
                # Scan seat types
                if st_col:
                    for _, r in df.iterrows():
                        st_val = r[st_col]
                        g_val = r[gen_col] if gen_col else None
                        token, is_allowed, raw = normalize_category_and_gender(st_val, g_val)
                        if is_allowed and token:
                            summary["detected_categories"].add(token)
                        elif raw:
                            summary["ignored_categories"].add(raw)
                            
        self.validation_summary = summary
        self.detected_categories = summary["detected_categories"]
        self.ignored_categories = summary["ignored_categories"]
        return summary

    def process_all(self, category_order_mode="IMAGE1"):
        """
        Parses all sheets and aggregates ranks.
        category_order_mode: 'IMAGE1' (OBC before EWS) or 'SECTION6' (EWS before OBC).
        """
        if not self.validation_summary:
            self.inspect()
            
        cat_order = CATEGORY_ORDER_IMAGE1 if category_order_mode == "IMAGE1" else CATEGORY_ORDER_SECTION6
        results = {}
        
        for kind in ["IIT", "NIT", "IIIT"]:
            if kind not in self.sheet_map:
                continue
                
            s_name = self.sheet_map[kind]
            df = self.xl.parse(s_name)
            df = df.dropna(how="all", axis=0).dropna(how="all", axis=1)
            
            # Map column names
            col_inst = next((c for c in df.columns if "inst" in str(c).lower()), None)
            col_prog = next((c for c in df.columns if "program" in str(c).lower() or "branch" in str(c).lower()), None)
            col_quota = next((c for c in df.columns if "quota" in str(c).lower()), None)
            col_seat = next((c for c in df.columns if "seat" in str(c).lower() or "category" in str(c).lower()), None)
            col_gender = next((c for c in df.columns if "gender" in str(c).lower()), None)
            col_open = next((c for c in df.columns if "open" in str(c).lower() and "rank" in str(c).lower()), None)
            col_close = next((c for c in df.columns if "clos" in str(c).lower() and "rank" in str(c).lower()), None)
            
            if not col_inst or not col_prog or not col_seat or not col_open or not col_close:
                raise ValueError(f"Missing essential columns in sheet {s_name}")
                
            institutes_data = []
            
            # Group by Institute preserving original order of appearance
            seen_institutes = []
            for inst in df[col_inst].dropna():
                inst_clean = str(inst).strip()
                if inst_clean not in seen_institutes:
                    seen_institutes.append(inst_clean)
                    
            for sr_no, inst_name in enumerate(seen_institutes, start=1):
                sub_df = df[df[col_inst].str.strip() == inst_name]
                meta = get_institute_metadata(inst_name)
                
                # In NIT, Quota is essential (HS, OS, etc.).
                # In IIT / IIIT, Quota is always AI.
                programs_dict = {}
                programs_order = []
                
                for _, row in sub_df.iterrows():
                    prog_name = str(row[col_prog]).strip()
                    quota_val = str(row[col_quota]).strip() if col_quota and pd.notna(row[col_quota]) else "AI"
                    seat_val = row[col_seat]
                    gender_val = row[col_gender] if col_gender else None
                    
                    token, is_allowed, raw = normalize_category_and_gender(seat_val, gender_val)
                    if not is_allowed or not token:
                        continue
                        
                    open_rk = parse_rank_value(row[col_open])
                    close_rk = parse_rank_value(row[col_close])
                    
                    # Row Key:
                    # For NIT, append [Quota] e.g. [HS] or [OS] to distinguish rows
                    if kind == "NIT":
                        row_display_name = f"{prog_name} [{quota_val}]"
                        row_key = (prog_name, quota_val)
                    else:
                        row_display_name = prog_name
                        row_key = (prog_name, "AI")
                        
                    if row_key not in programs_dict:
                        programs_dict[row_key] = {
                            "program_name": prog_name,
                            "display_name": row_display_name,
                            "quota": quota_val,
                            "ranks": {}
                        }
                        programs_order.append(row_key)
                        
                    curr_ranks = programs_dict[row_key]["ranks"]
                    if token not in curr_ranks:
                        curr_ranks[token] = (open_rk, close_rk)
                    else:
                        # Consolidate across rounds: min Opening Rank, max Closing Rank
                        prev_op, prev_cl = curr_ranks[token]
                        
                        new_op = prev_op
                        if open_rk is not None:
                            new_op = open_rk if prev_op is None else min(prev_op, open_rk)
                            
                        new_cl = prev_cl
                        if close_rk is not None:
                            new_cl = close_rk if prev_cl is None else max(prev_cl, close_rk)
                            
                        curr_ranks[token] = (new_op, new_cl)
                        
                branches_list = [programs_dict[k] for k in programs_order]
                institutes_data.append({
                    "sr_no": sr_no,
                    "institute_name": inst_name,
                    "meta": meta,
                    "branches": branches_list
                })
                
            results[kind] = {
                "category_order": cat_order,
                "institutes": institutes_data
            }
            
        self.processed_data = results
        return results
