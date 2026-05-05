# -*- coding: utf-8 -*-

import pandas as pd
import re
import sys
from openpyxl import load_workbook


# =========================================================
# COMMON
# =========================================================
def validate_column(df, col):
    if col not in df.columns:
        raise Exception(f"ไม่พบคอลัมน์: {col}")


# =========================================================
# ======================= SKC =============================
# =========================================================
def clean_desc(text):
    if pd.isna(text):
        return ""

    text = str(text).lstrip()

    if text.startswith("ST_"):
        text = re.sub(r'(\d+)\s*ํ', r'\1°', text)
        return re.sub(r"\s+", " ", text).strip()

    text = re.sub(r"^[_-]+", "", text)
    text = re.sub(r"^[^_\s]{1,50}_", "", text)
    text = re.sub(r'(\d+)\s*ํ', r'\1°', text)

    return re.sub(r"\s+", " ", text).strip()


def check_reason(text):
    if pd.isna(text) or str(text).strip() == "":
        return "blank"

    text = str(text).strip()

    if text.startswith("ST_"):
        return ""

    if re.search(r"^[_-]+", text):
        return "has_leading_symbol"

    if re.search(r"^[^_\s]{1,50}_", text):
        return "has_prefix_left"

    if len(text) <= 2:
        return "too_short"

    if re.fullmatch(r"\d+", text):
        return "numeric_only"

    if re.search(r'[^ก-๙a-zA-Z0-9°"()/\-\+\.\,\%*&]', text):
        return "weird_char"

    return ""


def process_skc(file_name):
    sheet_name = "Data"
    source_col = "Description  (Name Thai)"

    df = pd.read_excel(file_name, sheet_name=sheet_name)
    validate_column(df, source_col)

    df = df[~df[source_col].astype(str).str.contains("ยกเลิก", na=False)].copy()

    df["Description_Clean"] = df[source_col].apply(clean_desc)
    df["Changed"] = (df[source_col] != df["Description_Clean"]).map({True: "Yes", False: ""})

    df["Reason"] = df["Description_Clean"].apply(check_reason)
    df["Need_Review"] = df["Reason"].apply(lambda x: "Yes" if x != "" else "")

    return df


# =========================================================
# ======================= SNT =============================
# =========================================================
def detect_type(text):
    if pd.isna(text):
        return ""

    text = str(text).lstrip()
    text_check = re.sub(r"^_+", "", text)
    text_check = re.sub(r'^"+\s*', '', text_check)

    if re.match(r'^\(NO\.\d+\)', text_check, re.IGNORECASE):
        return "BRACKET_NO"

    if re.match(r'^\([^)]*\)', text_check):
        return "BRACKET_TEXT"

    if re.match(r'^\d+(?:/\d+)?x\d+(?:/\d+)?[A-Z]_', text_check, re.IGNORECASE):
        return "SIZE_PREFIX"

    if "_" in text_check:
        return "HAS_UNDERSCORE"

    return "CLEAN"


def clean_by_type(text, type_name):
    if pd.isna(text):
        return ""

    text = str(text).lstrip()
    text = re.sub(r"^_+", "", text)
    text = re.sub(r'^"+\s*', '', text)

    if type_name == "NUMERIC_PREFIX":
        text = re.sub(r'^\d+\s+', '', text)

    text = re.sub(r'(\d+)\s*ํ', r'\1°', text)
    return re.sub(r'\s+', ' ', text).strip()


def process_snt(file_name):
    sheet_name = "SNT"
    source_col = "Description  (Name Thai)"

    df = pd.read_excel(file_name, sheet_name=sheet_name)
    validate_column(df, source_col)

    df = df[~df[source_col].astype(str).str.contains("ยกเลิก", case=False, na=False)].copy()

    df["Type"] = df[source_col].apply(detect_type)

    df["Description_Clean"] = df.apply(
        lambda r: clean_by_type(r[source_col], r["Type"]),
        axis=1
    )

    return df


# =========================================================
# ======================= SPN =============================
# =========================================================
def base_clean(text):
    text = str(text).strip()
    text = text.replace('"', '')
    text = re.sub(r"^[\s_'`~\.,;:\|\?!\+=\-\*]+", "", text)
    text = re.sub(r'^[่-๋็์ํ]+', '', text)
    text = re.sub(r'^[^A-Za-z0-9ก-๙\(]+', '', text)
    return text


def final_clean(text):
    text = str(text)
    text = text.replace('"', '')
    text = re.sub(r'^\s+', '', text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def detect_type_spn(text):
    t = base_clean(text)

    if re.match(r'^ACCOUNTING\s*-\s*\d+$', t, re.IGNORECASE):
        return "KEEP"

    if re.match(r'^SPA\s*-\s*\d+\s*$', t):
        return "KEEP"

    if re.match(r'^IT-(08|09|10)\b', t, re.IGNORECASE):
        return "KEEP"

    if re.match(r'^TOA\s*-', t, re.IGNORECASE):
        return "KEEP"

    if re.match(r'^(TP-Link|TP-LINK)\b', t):
        return "KEEP"

    if re.match(r'^DH-IPC-HFW', t, re.IGNORECASE):
        return "KEEP"

    if re.match(r'^IPC-HFW', t, re.IGNORECASE):
        return "KEEP"

    if re.match(r'^USB-C\b', t, re.IGNORECASE):
        return "KEEP"

    return "CLEAN"


def clean_by_type_spn(text, typ):
    t = base_clean(text)

    if typ == "KEEP":
        return t

    t = re.sub(r'(\d+)\s*ํ', r'\1°', t)
    return re.sub(r'\s+', ' ', t).strip()


def process_spn(file_name):
    sheet_name = "SPN"
    source_col = "Name Thai"

    df = pd.read_excel(file_name, sheet_name=sheet_name)
    df.columns = df.columns.astype(str).str.replace(r"\s+", " ", regex=True).str.strip()

    validate_column(df, source_col)

    df = df[
        ~df[source_col].astype(str).str.contains(r"ยกเลิก|ขกเลิก", case=False, na=False)
    ].copy()

    df["Type"] = df[source_col].apply(detect_type_spn)

    df["Description_Clean"] = df.apply(
        lambda r: final_clean(clean_by_type_spn(r[source_col], r["Type"])),
        axis=1
    )

    return df


# =========================================================
# ======================= MAIN ============================
# =========================================================
def main():
    if len(sys.argv) < 2:
        raise Exception("python script.py input.xlsx")

    file_name = sys.argv[1]
    print("ใช้ไฟล์:", file_name)

    skc = process_skc(file_name)
    snt = process_snt(file_name)
    spn = process_spn(file_name)

    output = file_name.replace(".xlsx", "_clean.xlsx")

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        skc.to_excel(writer, sheet_name="SKC_clean", index=False)
        snt.to_excel(writer, sheet_name="SNT_clean", index=False)
        spn.to_excel(writer, sheet_name="SPN_clean", index=False)

    print("บันทึกไฟล์แล้ว:", output)


if __name__ == "__main__":
    main()