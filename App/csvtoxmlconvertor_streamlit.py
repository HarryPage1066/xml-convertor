"""
Salesforce Payments2Us CSV to Synergetic Donation XML Converter
----------------------------------------------------------------
Streamlit app for Finance users to validate and convert Salesforce exported
Payments2Us donation receipt data into Synergetic Standard Receipt XML.

Key behaviour:
- Generates Synergetic <Receipt> nodes with <DonationPayment> child nodes.
- Excludes rows without Contact Synergetic ID from XML.
- Produces an exceptions CSV for rows requiring remediation.
- Enriches rows with Fund and Appeal codes based on Merchant Facility.
- Maps Salesforce card descriptions to Synergetic luCreditCard codes.

Run:
    streamlit run csvtoxmlconvertor_streamlit.py
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any, Dict, Iterable, List, Optional, Tuple
from xml.dom.minidom import parseString
from xml.etree.ElementTree import Element, SubElement, tostring

import pandas as pd
import streamlit as st


# -----------------------------------------------------------------------------
# Configuration
# -----------------------------------------------------------------------------

APP_TITLE = "Salesforce Payments2Us → Synergetic Donation XML"
DEFAULT_SUPPLIER_REFERENCE = "KINGS"

# The Synergetic PDF's Receipt table says the Receipt date attribute is "Date"
# the sample donation/GL XML in the same PDF uses "ReceiptDate" on Receipt.
# Including both with the same value gives compatibility with either import parser.
INCLUDE_RECEIPT_DATE_ALIAS = True

# Do not include BankCode by default because this Synergetic environment does not
# have luBank available and the Synergetic documentation describes BankCode as a
# luBank lookup.
INCLUDE_BANK_CODE = False

# Do not include masked card numbers by default. The Synergetic documentation
# describes CreditCardNumber as the full card number, while Salesforce provides
# masked card numbers.
INCLUDE_MASKED_CARD_NUMBER = False

COL_FACILITY = "Merchant Facility: Merchant Facility Name"
COL_BANK_CODE = "Receipt Bank Code"
COL_CARD_NAME = "Name On Card"
COL_MASKED_CARD = "Masked Credit Card No."
COL_CARD_TYPE = "Credit Card Type"
COL_DEPOSIT_DATE = "Bank Deposit Date"
COL_CONTACT_ID = "Contact Synergetic ID"
COL_RECEIPT_NO = "Receipt No."
COL_TOTAL_AMOUNT = "Total Amount Charged"
COL_BANKED_AMOUNT = "Banked Amount"
COL_GL_CODE = "GL Code"
COL_GL_DESCRIPTION = "GL Description"
COL_GL_RECEIPT_NAME = "GL Receipt Name"

REQUIRED_COLUMNS = [
    COL_FACILITY,
    COL_CARD_NAME,
    COL_CARD_TYPE,
    COL_DEPOSIT_DATE,
    COL_CONTACT_ID,
    COL_RECEIPT_NO,
    COL_TOTAL_AMOUNT,
    COL_BANKED_AMOUNT,
    COL_GL_RECEIPT_NAME,
]

# Merchant Facility enrichment. These keys must match Salesforce export values.
PROJECT_MAPPING: Dict[str, Dict[str, str]] = {
    "The Council of the King's School - PREP (GD 2026)": {
        "Project": "Prep Playground",
        "Entity": "TKS",
        "BankAccountCode": "PREPPLAY",
        "BankAccountGL": "1510.5.000",
        "FundCode": "PP",
        "FundGL": "1510.5.000",
        "AppealCode": "PP",
        "AppealGL": "1510.5.000",
    },
    "The King’s School General Building Fund - Swimming Pool Stadium - Tax Deductible": {
        "Project": "Swimming Pool Stadium",
        "Entity": "TKSF",
        "BankAccountCode": "BUICHQ",
        "BankAccountGL": "1100.72.100",
        "FundCode": "FBFTAC",
        "FundGL": "5105.72.200",
        "AppealCode": "TAC",
        "AppealGL": "5105.72.200",
    },
    "Scholarships & Bursaries Fund - Next Crop Support Program 2026 - Tax Deductible": {
        "Project": "Next Crop",
        "Entity": "TKSF",
        "BankAccountCode": "SBPCHQ",
        "BankAccountGL": "1100.73.100",
        "FundCode": "FSBPNC",
        "FundGL": "5100.73.304",
        "AppealCode": "NCSBP",
        "AppealGL": "5100.73.304",
    },
    "Tudor House General Building Fund - Tax deductible": {
        "Project": "Tudor House Building",
        "Entity": "THF",
        "BankAccountCode": "BUICHQ",
        "BankAccountGL": "1100.62.200",
        "FundCode": "TBFSHED",
        "FundGL": "5106.62.200",
        "AppealCode": "RECSHED",
        "AppealGL": "5106.62.200",
    },
}

# Card mappings from synergetic [pvCreditCards] view in production
CARD_TYPE_MAPPING = {
    "visa": "VISA",
    "visa card": "VISA",
    "mastercard": "MCARD",
    "master card": "MCARD",
    "m/card": "MCARD",
    "mc": "MCARD",
    "american express": "AMEX",
    "amex": "AMEX",
    "bankcard": "BCARD",
    "bank card": "BCARD",
    "diners card": "DINER",
    "diners": "DINER",
    "school easy pay": "SEP",
}

MONEY_QUANT = Decimal("0.01")


# -----------------------------------------------------------------------------
# Data classes
# -----------------------------------------------------------------------------

@dataclass
class ProcessedResult:
    valid_rows: pd.DataFrame
    exceptions: pd.DataFrame
    warnings: pd.DataFrame
    project_totals: pd.DataFrame
    xml_text: str
    total_source_amount: Decimal
    total_xml_amount: Decimal
    missing_id_count: int


# -----------------------------------------------------------------------------
# Utility functions
# -----------------------------------------------------------------------------

def is_blank(value: Any) -> bool:
    """Return True for None, NaN, empty strings, and textual null-like values."""
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except TypeError:
        pass
    text = str(value).strip()
    return text == "" or text.lower() in {"nan", "none", "null", "nat"}


def clean_text(value: Any) -> str:
    """Return a trimmed string, or an empty string for blank/null values."""
    if is_blank(value):
        return ""
    return str(value).strip()


def truncate_text(value: str, max_len: int) -> str:
    """Trim text to a Synergetic field limit without raising an import-blocking error."""
    return value[:max_len] if len(value) > max_len else value


def first_non_blank(*values: Any) -> str:
    for value in values:
        text = clean_text(value)
        if text:
            return text
    return ""


def parse_integer(value: Any, field_name: str) -> str:
    """Parse an integer-like value and return it without decimals or separators."""
    if is_blank(value):
        raise ValueError(f"Missing {field_name}")

    text = clean_text(value).replace(",", "")
    try:
        decimal_value = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f"{field_name} is not a valid integer: {clean_text(value)}") from exc

    if decimal_value != decimal_value.to_integral_value():
        raise ValueError(f"{field_name} must be a whole number: {clean_text(value)}")

    return str(int(decimal_value))


def parse_money(value: Any, field_name: str) -> Decimal:
    """Parse money values, accepting $, commas, spaces, and accounting negatives."""
    if is_blank(value):
        raise ValueError(f"Missing {field_name}")

    text = clean_text(value)
    is_negative = text.startswith("(") and text.endswith(")")
    text = text.replace("$", "").replace(",", "").replace(" ", "")
    text = text.strip("()")

    try:
        amount = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f"{field_name} is not a valid money amount: {clean_text(value)}") from exc

    if is_negative:
        amount = -amount

    return amount.quantize(MONEY_QUANT, rounding=ROUND_HALF_UP)


def format_money(amount: Decimal) -> str:
    return f"{amount.quantize(MONEY_QUANT, rounding=ROUND_HALF_UP):.2f}"


def parse_au_date(value: Any, field_name: str) -> str:
    """Parse dates as Australian dates and return DD/MM/YYYY."""
    if is_blank(value):
        raise ValueError(f"Missing {field_name}")

    parsed = pd.to_datetime(clean_text(value), dayfirst=True, errors="coerce")
    if pd.isna(parsed):
        raise ValueError(f"{field_name} is not a valid date: {clean_text(value)}")

    return parsed.strftime("%d/%m/%Y")


def normalise_card_type(value: Any) -> Tuple[Optional[str], Optional[str]]:
    """Map Salesforce card descriptions to Synergetic luCreditCard.Code values."""
    if is_blank(value):
        return None, None

    source = clean_text(value)
    code = CARD_TYPE_MAPPING.get(source.lower())
    if not code:
        return None, f"Unknown credit card type '{source}'. CreditCardType was omitted from XML."
    return code, None


def add_optional_attr(attrs: Dict[str, str], key: str, value: Any, max_len: Optional[int] = None) -> None:
    text = clean_text(value)
    if not text:
        return
    attrs[key] = truncate_text(text, max_len) if max_len else text


def get_duplicate_receipt_numbers(df: pd.DataFrame) -> set[str]:
    normalised: List[str] = []
    for value in df.get(COL_RECEIPT_NO, pd.Series(dtype=str)):
        try:
            normalised.append(parse_integer(value, COL_RECEIPT_NO))
        except ValueError:
            continue
    counts = Counter(normalised)
    return {receipt_no for receipt_no, count in counts.items() if count > 1}


def required_action_for(reasons: Iterable[str]) -> str:
    reason_text = " | ".join(reasons)
    actions: List[str] = []

    if "Contact Synergetic ID" in reason_text:
        actions.append(
            "Create a Synergetic Community record for the individual, then back-populate "
            "Contact Synergetic ID in this CSV and rerun the conversion."
        )

    if "Merchant Facility" in reason_text:
        actions.append(
            "Add or correct the Merchant Facility mapping before rerunning the conversion."
        )

    if "Receipt No." in reason_text or "duplicate" in reason_text.lower():
        actions.append("Correct Receipt No. so each receipt number is present and unique in the file.")

    if "Amount" in reason_text or "amount" in reason_text:
        actions.append("Correct the amount fields so Total Amount Charged and Banked Amount are valid and equal.")

    if "date" in reason_text.lower():
        actions.append("Correct the date so it can be read as DD/MM/YYYY.")

    if not actions:
        actions.append("Correct the source row and rerun the conversion.")

    # Preserve order while removing duplicates.
    return " ".join(dict.fromkeys(actions))


# -----------------------------------------------------------------------------
# Processing and validation
# -----------------------------------------------------------------------------

def read_salesforce_csv(uploaded_file: Any) -> pd.DataFrame:
    """Read uploaded CSV as strings so IDs and receipt numbers are not damaged."""
    df = pd.read_csv(uploaded_file, dtype=str, keep_default_na=False)
    df.columns = [str(col).strip() for col in df.columns]
    return df


def validate_required_columns(df: pd.DataFrame) -> None:
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        missing_list = ", ".join(missing)
        raise ValueError(f"The uploaded CSV is missing required column(s): {missing_list}")


def process_dataframe(df: pd.DataFrame, supplier_reference: str) -> ProcessedResult:
    validate_required_columns(df)

    duplicate_receipts = get_duplicate_receipt_numbers(df)
    valid_records: List[Dict[str, Any]] = []
    exception_records: List[Dict[str, Any]] = []
    warning_records: List[Dict[str, Any]] = []

    total_source_amount = Decimal("0.00")
    missing_id_count = 0

    for idx, row in df.iterrows():
        original_row_number = idx + 2  # CSV row number, allowing for header row.
        reasons: List[str] = []
        warnings: List[str] = []
        project_info: Optional[Dict[str, str]] = None

        row_dict = row.to_dict()
        facility = clean_text(row_dict.get(COL_FACILITY))
        if not facility:
            reasons.append("Missing Merchant Facility")
        else:
            project_info = PROJECT_MAPPING.get(facility)
            if project_info is None:
                reasons.append(f"Unknown Merchant Facility: {facility}")

        # Receipt number
        receipt_number = ""
        try:
            receipt_number = parse_integer(row_dict.get(COL_RECEIPT_NO), COL_RECEIPT_NO)
            if receipt_number in duplicate_receipts:
                reasons.append(f"Duplicate Receipt No. within file: {receipt_number}")
        except ValueError as exc:
            reasons.append(str(exc))

        # Contact/Synergetic ID
        contact_id = ""
        try:
            contact_id = parse_integer(row_dict.get(COL_CONTACT_ID), COL_CONTACT_ID)
        except ValueError as exc:
            reasons.append(str(exc))
            if "Missing" in str(exc):
                missing_id_count += 1

        # Date
        receipt_date = ""
        try:
            receipt_date = parse_au_date(row_dict.get(COL_DEPOSIT_DATE), COL_DEPOSIT_DATE)
        except ValueError as exc:
            reasons.append(str(exc))

        # Amounts
        total_amount = Decimal("0.00")
        banked_amount = Decimal("0.00")
        try:
            total_amount = parse_money(row_dict.get(COL_TOTAL_AMOUNT), COL_TOTAL_AMOUNT)
            total_source_amount += total_amount
        except ValueError as exc:
            reasons.append(str(exc))

        try:
            banked_amount = parse_money(row_dict.get(COL_BANKED_AMOUNT), COL_BANKED_AMOUNT)
        except ValueError as exc:
            reasons.append(str(exc))

        if total_amount and banked_amount and total_amount != banked_amount:
            reasons.append(
                f"Total Amount Charged ({format_money(total_amount)}) does not match "
                f"Banked Amount ({format_money(banked_amount)})"
            )

        # Names
        drawer = first_non_blank(row_dict.get(COL_GL_RECEIPT_NAME), row_dict.get(COL_CARD_NAME))
        if not drawer:
            reasons.append("Missing Drawer/Receipt name. GL Receipt Name or Name On Card is required.")

        receipt_name = first_non_blank(row_dict.get(COL_GL_RECEIPT_NAME), drawer)
        credit_card_name = clean_text(row_dict.get(COL_CARD_NAME))

        # Credit card type is optional in the Synergetic Receipt node, but bad lookup
        # values should not be written into XML.
        credit_card_type, card_warning = normalise_card_type(row_dict.get(COL_CARD_TYPE))
        if card_warning:
            warnings.append(card_warning)

        if reasons:
            exception_record = dict(row_dict)
            exception_record["OriginalRowNumber"] = original_row_number
            exception_record["ExclusionReason"] = " | ".join(reasons)
            exception_record["RequiredAction"] = required_action_for(reasons)
            exception_records.append(exception_record)
            continue

        assert project_info is not None  # for type-checkers; guaranteed by reasons check above.

        valid_record: Dict[str, Any] = dict(row_dict)
        valid_record.update(
            {
                "OriginalRowNumber": original_row_number,
                "ReceiptNumber": receipt_number,
                "ReceiptDate": receipt_date,
                "ContactSynergeticID": contact_id,
                "Drawer": truncate_text(drawer, 100),
                "ReceiptName": truncate_text(receipt_name, 100),
                "CreditCardName": truncate_text(credit_card_name, 100),
                "CreditCardTypeCode": credit_card_type or "",
                "AmountDecimal": total_amount,
                "Amount": format_money(total_amount),
                "Project": project_info["Project"],
                "Entity": project_info["Entity"],
                "BankAccountCode": project_info["BankAccountCode"],
                "BankAccountGL": project_info["BankAccountGL"],
                "ReceiptFundCode": project_info["FundCode"],
                "ReceiptFundGL": project_info["FundGL"],
                "ReceiptAppealCode": project_info["AppealCode"],
                "ReceiptAppealGL": project_info["AppealGL"],
            }
        )
        valid_records.append(valid_record)

        for warning in warnings:
            warning_records.append(
                {
                    "OriginalRowNumber": original_row_number,
                    "ReceiptNumber": receipt_number,
                    "Warning": warning,
                }
            )

    valid_df = pd.DataFrame(valid_records)
    exceptions_df = pd.DataFrame(exception_records)
    warnings_df = pd.DataFrame(warning_records)

    total_xml_amount = (
        sum(valid_df["AmountDecimal"], Decimal("0.00")) if not valid_df.empty else Decimal("0.00")
    )

    if not valid_df.empty:
        project_totals = (
            valid_df.groupby(
                ["Project", "Entity", "BankAccountCode", "ReceiptFundCode", "ReceiptAppealCode"],
                as_index=False,
            )
            .agg(Records=("ReceiptNumber", "count"), TotalAmount=("AmountDecimal", "sum"))
            .sort_values(["Project"])
        )
        project_totals["TotalAmount"] = project_totals["TotalAmount"].apply(format_money)
    else:
        project_totals = pd.DataFrame(
            columns=[
                "Project",
                "Entity",
                "BankAccountCode",
                "ReceiptFundCode",
                "ReceiptAppealCode",
                "Records",
                "TotalAmount",
            ]
        )

    xml_text = build_synergetic_xml(valid_df, supplier_reference=supplier_reference)

    return ProcessedResult(
        valid_rows=valid_df,
        exceptions=exceptions_df,
        warnings=warnings_df,
        project_totals=project_totals,
        xml_text=xml_text,
        total_source_amount=total_source_amount,
        total_xml_amount=total_xml_amount,
        missing_id_count=missing_id_count,
    )


# -----------------------------------------------------------------------------
# XML generation
# -----------------------------------------------------------------------------

def build_synergetic_xml(valid_df: pd.DataFrame, supplier_reference: str) -> str:
    root = Element(
        "Synergetic",
        {
            "FileVersion": "1.00",
            "FileCreated": datetime.now().strftime("%d/%m/%Y"),
            "SupplierReference": supplier_reference or DEFAULT_SUPPLIER_REFERENCE,
        },
    )

    if valid_df.empty:
        return pretty_xml(root)

    for _, row in valid_df.iterrows():
        receipt_attrs: Dict[str, str] = {
            "ReceiptNumber": str(row["ReceiptNumber"]),
            "Date": str(row["ReceiptDate"]),
            "ID": str(row["ContactSynergeticID"]),
            "Drawer": str(row["Drawer"]),
            "TotalAmount": str(row["Amount"]),
        }

        if INCLUDE_RECEIPT_DATE_ALIAS:
            receipt_attrs["ReceiptDate"] = str(row["ReceiptDate"])

        if INCLUDE_BANK_CODE:
            add_optional_attr(receipt_attrs, "BankCode", row.get(COL_BANK_CODE), 5)

        add_optional_attr(receipt_attrs, "CreditCardType", row.get("CreditCardTypeCode"), 5)
        add_optional_attr(receipt_attrs, "CreditCardName", row.get("CreditCardName"), 100)

        if INCLUDE_MASKED_CARD_NUMBER:
            add_optional_attr(receipt_attrs, "CreditCardNumber", row.get(COL_MASKED_CARD), 20)

        receipt = SubElement(root, "Receipt", receipt_attrs)

        donation_attrs: Dict[str, str] = {
            "ReceiptID": str(row["ContactSynergeticID"]),
            "ReceiptIDJoint": "0",
            "ReceiptDate": str(row["ReceiptDate"]),
            "ReceiptFundCode": str(row["ReceiptFundCode"]),
            "ReceiptAppealCode": str(row["ReceiptAppealCode"]),
            "ReceiptAmount": str(row["Amount"]),
            "ReceiptAnonymousFlag": "0",
            "ReceiptRecognitionFlag": "1",
        }
        add_optional_attr(donation_attrs, "ReceiptName", row.get("ReceiptName"), 100)
        add_optional_attr(
            donation_attrs,
            "ReceiptComments",
            f"{row.get('Project', '')} - Salesforce receipt {row.get(COL_RECEIPT_NO, '')}",
            255,
        )

        SubElement(receipt, "DonationPayment", donation_attrs)

    return pretty_xml(root)


def pretty_xml(root: Element) -> str:
    rough_xml = tostring(root, encoding="utf-8")
    pretty = parseString(rough_xml).toprettyxml(indent="  ", encoding="utf-8").decode("utf-8")

    # Synergetic's sample files start directly with <Synergetic>, so remove
    # the XML declaration generated by minidom.
    xml_lines = [
        line
        for line in pretty.splitlines()
        if line.strip() and not line.strip().startswith("<?xml")
    ]

    return "\n".join(xml_lines)


# -----------------------------------------------------------------------------
# Streamlit presentation helpers
# -----------------------------------------------------------------------------

def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    export_df = df.copy()
    for column in export_df.columns:
        if export_df[column].dtype == "object":
            export_df[column] = export_df[column].apply(
                lambda value: format_money(value) if isinstance(value, Decimal) else value
            )
    return export_df.to_csv(index=False).encode("utf-8-sig")


def build_audit_dataframe(valid_df: pd.DataFrame) -> pd.DataFrame:
    if valid_df.empty:
        return pd.DataFrame()

    audit_cols = [
        "OriginalRowNumber",
        COL_RECEIPT_NO,
        "ReceiptNumber",
        COL_DEPOSIT_DATE,
        "ReceiptDate",
        COL_CONTACT_ID,
        "ContactSynergeticID",
        COL_GL_RECEIPT_NAME,
        "Drawer",
        COL_FACILITY,
        "Project",
        "Entity",
        "BankAccountCode",
        "BankAccountGL",
        "ReceiptFundCode",
        "ReceiptFundGL",
        "ReceiptAppealCode",
        "ReceiptAppealGL",
        COL_CARD_TYPE,
        "CreditCardTypeCode",
        COL_TOTAL_AMOUNT,
        COL_BANKED_AMOUNT,
        "Amount",
    ]
    existing_cols = [col for col in audit_cols if col in valid_df.columns]
    return valid_df[existing_cols].copy()


def display_metric_row(result: ProcessedResult, source_row_count: int) -> None:
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Rows uploaded", f"{source_row_count:,}")
    col2.metric("Rows in XML", f"{len(result.valid_rows):,}")
    col3.metric("Rows excluded", f"{len(result.exceptions):,}")
    col4.metric("XML total", f"${format_money(result.total_xml_amount)}")


def display_missing_id_message(result: ProcessedResult) -> None:
    if result.missing_id_count <= 0:
        return

    st.warning(
        f"{result.missing_id_count:,} record(s) were excluded because they do not have a "
        "Contact Synergetic ID. Create a Synergetic Community record for each individual, "
        "then back-populate the Contact Synergetic ID into the exceptions CSV and rerun "
        "the conversion using that updated CSV."
    )


def display_download_buttons(result: ProcessedResult, audit_df: pd.DataFrame) -> None:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    col1, col2, col3 = st.columns(3)

    with col1:
        st.download_button(
            label="Download Synergetic XML",
            data=result.xml_text,
            file_name=f"synergetic_donation_receipts_{timestamp}.xml",
            mime="application/xml",
            type="primary",
            use_container_width=True,
            disabled=result.valid_rows.empty,
        )

    with col2:
        st.download_button(
            label="Download exceptions CSV",
            data=dataframe_to_csv_bytes(result.exceptions),
            file_name=f"synergetic_donation_exceptions_{timestamp}.csv",
            mime="text/csv",
            use_container_width=True,
            disabled=result.exceptions.empty,
        )

    with col3:
        st.download_button(
            label="Download audit CSV",
            data=dataframe_to_csv_bytes(audit_df),
            file_name=f"synergetic_donation_audit_{timestamp}.csv",
            mime="text/csv",
            use_container_width=True,
            disabled=audit_df.empty,
        )


def format_result_tables(result: ProcessedResult, audit_df: pd.DataFrame) -> None:
    tabs = st.tabs(["Project totals", "Exceptions", "Audit", "XML preview", "Warnings"])

    with tabs[0]:
        st.subheader("Project totals included in XML")
        if result.project_totals.empty:
            st.info("No valid rows are available for XML generation.")
        else:
            st.dataframe(result.project_totals, use_container_width=True, hide_index=True)

    with tabs[1]:
        st.subheader("Rows excluded from XML")
        if result.exceptions.empty:
            st.success("No rows were excluded.")
        else:
            preview_cols = [
                "OriginalRowNumber",
                COL_RECEIPT_NO,
                COL_GL_RECEIPT_NAME,
                COL_CONTACT_ID,
                COL_TOTAL_AMOUNT,
                COL_FACILITY,
                "ExclusionReason",
                "RequiredAction",
            ]
            cols = [col for col in preview_cols if col in result.exceptions.columns]
            st.dataframe(result.exceptions[cols], use_container_width=True, hide_index=True)

    with tabs[2]:
        st.subheader("Audit of rows included in XML")
        if audit_df.empty:
            st.info("No audit rows are available.")
        else:
            st.dataframe(audit_df, use_container_width=True, hide_index=True)

    with tabs[3]:
        st.subheader("XML preview")
        if result.valid_rows.empty:
            st.info("No XML receipts were generated because there are no valid rows.")
        else:
            preview = result.xml_text[:15000]
            if len(result.xml_text) > len(preview):
                preview += "\n<!-- Preview truncated. Download the XML to view the full file. -->"
            st.code(preview, language="xml")

    with tabs[4]:
        st.subheader("Non-blocking warnings")
        if result.warnings.empty:
            st.success("No non-blocking warnings.")
        else:
            st.dataframe(result.warnings, use_container_width=True, hide_index=True)


def show_sidebar() -> str:
    st.sidebar.header("Import settings")
    supplier_reference = st.sidebar.text_input(
        "Supplier reference",
        value=DEFAULT_SUPPLIER_REFERENCE,
        help="Written to the Synergetic root XML SupplierReference attribute.",
    ).strip() or DEFAULT_SUPPLIER_REFERENCE

    st.sidebar.divider()
    st.sidebar.caption("This converter always creates Receipt + DonationPayment XML.")
    st.sidebar.caption("Rows without Contact Synergetic ID are excluded and written to exceptions CSV.")
    st.sidebar.caption("BankCode and masked card number are omitted by default.")

    return supplier_reference


# -----------------------------------------------------------------------------
# Streamlit app
# -----------------------------------------------------------------------------

def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="🧾", layout="wide", initial_sidebar_state="collapsed")

    supplier_reference = show_sidebar()

    st.title(APP_TITLE)
    st.write(
        "Upload a Salesforce Payments2Us donation CSV. The app validates the file, "
        "excludes records that cannot be imported, enriches each valid donation with the "
        "correct Synergetic Fund and Appeal, and generates a Synergetic XML import file."
    )

    with st.expander("What this converter does", expanded=False):
        st.markdown(
            """
            - Generates **Synergetic `Receipt` + `DonationPayment` XML** for donation imports.
            - Uses `Merchant Facility: Merchant Facility Name` to derive Project, Entity, Bank Account, Fund and Appeal.
            - Maps Salesforce card types to Synergetic `luCreditCard.Code` values: `VISA`, `MCARD`, `AMEX`, etc.
            - Excludes rows without `Contact Synergetic ID` and writes them to an exceptions CSV for remediation.
            - Does **not** create GLPayment XML for this process, because the Fund and Appeal should drive the donation posting.
            """
        )

    uploaded_file = st.file_uploader("Upload Salesforce donation CSV", type=["csv"])

    if uploaded_file is None:
        st.info("Upload a CSV file to validate and generate Synergetic XML.")
        return

    try:
        df = read_salesforce_csv(uploaded_file)
    except Exception as exc:
        st.error(f"The uploaded file could not be read as a CSV: {exc}")
        return

    if df.empty:
        st.error("The uploaded CSV is empty.")
        return

    st.success(f"Loaded {len(df):,} row(s) from `{uploaded_file.name}`.")

    try:
        with st.spinner("Validating and generating XML..."):
            result = process_dataframe(df, supplier_reference=supplier_reference)
            audit_df = build_audit_dataframe(result.valid_rows)
    except ValueError as exc:
        st.error(str(exc))
        return
    except Exception as exc:
        st.error(f"An unexpected error occurred while processing the file: {exc}")
        return

    display_metric_row(result, source_row_count=len(df))
    display_missing_id_message(result)

    if result.valid_rows.empty:
        st.error(
            "No valid rows are available to write to XML. Download the exceptions CSV, "
            "fix the listed issues, and rerun the conversion."
        )
    elif result.exceptions.empty:
        st.success("All rows passed validation and were included in the XML.")
    else:
        st.info(
            f"{len(result.valid_rows):,} row(s) were included in the XML and "
            f"{len(result.exceptions):,} row(s) were excluded."
        )

    display_download_buttons(result, audit_df)
    format_result_tables(result, audit_df)


if __name__ == "__main__":
    main()
