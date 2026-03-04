from __future__ import annotations

import os
from datetime import date, datetime
from pathlib import Path

import streamlit as st

OCR_ENABLED = os.getenv("ENABLE_OCR", "false").lower() == "true"

from contract_utils import (
    DEFAULT_PROPERTY_CSV_PATH,
    DEFAULT_TEMPLATE_PATH,
    FIXED_AUTHORISED_SIGNATORY_NAME,
    FIXED_COMPANY_NAME,
    FIXED_OWNER_MOBILE_NUMBER,
    FIXED_OWNER_REGISTERED_ADDRESS,
    PAYMENT_DAY_OPTIONS,
    PLACEHOLDER_DESCRIPTIONS,
    TENANT_DOC_TYPES,
    TenantData,
    build_placeholder_mapping,
    convert_docx_to_pdf,
    discover_template_placeholders,
    discover_template_placeholders_from_bytes,
    extract_id_details,
    read_properties,
    read_properties_from_bytes,
    render_contract,
    render_contract_from_bytes,
)


st.set_page_config(
    page_title="UK Tenancy Agreement Generator",
    layout="wide",
)


def main() -> None:
    st.title("UK Tenancy Agreement Generator")
    st.caption("Generate a tenancy agreement in British English from the supplied DOCX template and property register.")

    # --- Template file ---
    uploaded_template = st.sidebar.file_uploader(
        "Replace template (.docx)",
        type=["docx"],
        help="Leave blank to use the default VJProp template bundled with the app.",
    )
    if uploaded_template is not None:
        template_bytes = uploaded_template.read()
    elif DEFAULT_TEMPLATE_PATH.exists():
        template_bytes = DEFAULT_TEMPLATE_PATH.read_bytes()
    else:
        st.error("Default template not found. Please upload a .docx template using the sidebar.")
        st.stop()

    # --- Property CSV ---
    uploaded_csv = st.sidebar.file_uploader(
        "Replace property list (.csv)",
        type=["csv"],
        help="Leave blank to use the default property list bundled with the app.",
    )
    if uploaded_csv is not None:
        csv_bytes = uploaded_csv.read()
    elif DEFAULT_PROPERTY_CSV_PATH.exists():
        csv_bytes = DEFAULT_PROPERTY_CSV_PATH.read_bytes()
    else:
        st.error("Default property list not found. Please upload a CSV using the sidebar.")
        st.stop()

    try:
        properties = read_properties_from_bytes(csv_bytes)
        discovered_placeholders = discover_template_placeholders_from_bytes(template_bytes)
    except Exception as exc:  # noqa: BLE001
        st.exception(exc)
        st.stop()

    property_titles = [record.title for record in properties]
    property_lookup = {record.title: record for record in properties}

    st.subheader("Template Mapping")
    st.write("These are the merge fields detected in the supplied template. Each value is entered once and populated everywhere it appears.")
    mapping_rows = [
        {
            "Placeholder": placeholder,
            "Description": PLACEHOLDER_DESCRIPTIONS.get(placeholder, "Mapped automatically by the generator."),
        }
        for placeholder in discovered_placeholders
        if placeholder not in {"Owner's Mobile Number", "Apartment Description"}
    ]
    st.dataframe(mapping_rows, use_container_width=True, hide_index=True)

    left_column, right_column = st.columns(2)

    with left_column:
        st.subheader("Property")
        property_title = st.selectbox("Select a property", options=property_titles, index=0)
        property_record = property_lookup[property_title]
        st.text_input("Development name", value=property_record.building_name, disabled=True)
        st.text_input("Apartment number", value=property_record.apartment_number, disabled=True)
        st.text_area("Full property address", value=property_record.full_address, height=90, disabled=True)

        st.subheader("Landlord")
        company_name = FIXED_COMPANY_NAME
        owner_registered_address = FIXED_OWNER_REGISTERED_ADDRESS
        owner_mobile_number = FIXED_OWNER_MOBILE_NUMBER
        authorised_signatory_name = FIXED_AUTHORISED_SIGNATORY_NAME
        st.text_input("Landlord or company name", value=company_name, disabled=True)
        st.text_area("Registered address", value=owner_registered_address, height=90, disabled=True)
        st.text_input("Authorised signatory name", value=authorised_signatory_name, disabled=True)

        st.subheader("Financial Terms")
        deposit_amount = st.text_input("Deposit amount (£)", placeholder="e.g. 1200")
        monthly_rent = st.text_input("Monthly rent (£)", placeholder="e.g. 850")
        payment_day = st.selectbox("First payment day", options=PAYMENT_DAY_OPTIONS, index=0)
        notice_period = st.text_input("Notice period", value="one calendar month")
        minimum_occupation_period = st.text_input(
            "Minimum occupation period before notice",
            value="11 months",
        )

    with right_column:
        st.subheader("Dates")
        today_string = date.today().strftime("%d/%m/%Y")
        agreement_date_raw = st.text_input(
            "Agreement date (DD/MM/YYYY)",
            value=st.session_state.get("agreement_date_raw", today_string),
        )
        start_date_raw = st.text_input(
            "First payment date / tenancy start (DD/MM/YYYY)",
            value=st.session_state.get("start_date_raw", today_string),
        )
        end_date_raw = st.text_input(
            "End date of term (DD/MM/YYYY)",
            value=st.session_state.get("end_date_raw", today_string),
        )

        st.subheader("Tenants")
        tenant_count = st.number_input(
            "How many tenants (excluding the landlord) will be on the contract?",
            min_value=1,
            max_value=6,
            value=1,
            step=1,
        )
        tenant_count = int(tenant_count)
        tenants = collect_tenants(tenant_count)

    download_format = st.radio("Preferred download format", options=("Word (.docx)", "PDF"), horizontal=True)
    generate_clicked = st.button("Generate agreement", type="primary")

    if not generate_clicked:
        return

    if not tenants or not all(tenant.full_name.strip() for tenant in tenants):
        st.error("Enter the full name for each tenant.")
        return

    try:
        agreement_date = parse_ui_date(agreement_date_raw)
        start_date = parse_ui_date(start_date_raw)
        end_date = parse_ui_date(end_date_raw)
    except ValueError as exc:
        st.error(str(exc))
        return

    placeholder_mapping = build_placeholder_mapping(
        agreement_date=agreement_date,
        start_date=start_date,
        end_date=end_date,
        payment_day=payment_day,
        deposit_amount=deposit_amount,
        monthly_rent=monthly_rent,
        notice_period=notice_period,
        minimum_occupation_period=minimum_occupation_period,
        company_name=company_name,
        owner_registered_address=owner_registered_address,
        owner_mobile_number=owner_mobile_number,
        authorised_signatory_name=authorised_signatory_name,
        property_record=property_record,
        tenants=tenants,
    )

    try:
        docx_bytes = render_contract_from_bytes(template_bytes, placeholder_mapping, tenants)
    except Exception as exc:  # noqa: BLE001
        st.exception(exc)
        return

    st.success("Agreement generated successfully.")

    st.download_button(
        "Download Word (.docx)",
        data=docx_bytes,
        file_name="uk-tenancy-agreement.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    if download_format == "PDF":
        try:
            pdf_bytes = convert_docx_to_pdf(docx_bytes)
            st.download_button(
                "Download PDF",
                data=pdf_bytes,
                file_name="uk-tenancy-agreement.pdf",
                mime="application/pdf",
            )
        except Exception as exc:  # noqa: BLE001
            st.warning(f"PDF conversion is unavailable right now: {exc}")
    else:
        st.info("Select PDF above if you also want a PDF download.")


def collect_tenants(tenant_count: int) -> list[TenantData]:
    tenants: list[TenantData] = []
    for index in range(tenant_count):
        tenant_number = index + 1
        tenant_key = f"tenant_{tenant_number}"
        apply_ocr_prefill(tenant_key)
        with st.expander(f"Tenant {tenant_number}", expanded=True):
            full_name = st.text_input("Full name", key=f"{tenant_key}_name")
            date_of_birth = st.text_input("Date of birth (DD/MM/YYYY)", key=f"{tenant_key}_dob")
            has_national_insurance = st.selectbox(
                "National Insurance number?",
                options=("No", "Yes"),
                key=f"{tenant_key}_ni_enabled",
            )
            national_insurance_number = ""
            if has_national_insurance == "Yes":
                national_insurance_number = st.text_input(
                    "National Insurance number",
                    key=f"{tenant_key}_ni",
                )
            id_document_type = st.selectbox(
                "ID document type",
                options=TENANT_DOC_TYPES,
                format_func=lambda value: value if value else "Leave blank",
                key=f"{tenant_key}_id_type",
            )
            id_number = ""
            if id_document_type:
                id_number = st.text_input("ID number", key=f"{tenant_key}_id_number")
            if OCR_ENABLED:
                uploaded_id = st.file_uploader(
                    "Upload ID image for OCR",
                    type=["png", "jpg", "jpeg"],
                    key=f"{tenant_key}_upload",
                    help="Upload a clear image. OCR is best-effort and should be reviewed before generating the contract.",
                )
                if st.button(f"Extract from ID for tenant {tenant_number}", key=f"{tenant_key}_ocr"):
                    if uploaded_id is None:
                        st.warning("Upload an ID image before running OCR.")
                    elif not id_document_type:
                        st.warning("Select an ID document type before running OCR.")
                    else:
                        try:
                            extracted = extract_id_details(uploaded_id, uploaded_id.name, id_document_type)
                            st.session_state[f"{tenant_key}_ocr_prefill"] = extracted
                            st.rerun()
                        except Exception as exc:  # noqa: BLE001
                            st.warning(f"OCR could not extract the ID details: {exc}")
            else:
                st.info("💡 OCR is only available on the internal desktop version of this tool.")

            tenants.append(
                TenantData(
                    full_name=full_name,
                    date_of_birth=date_of_birth,
                    national_insurance_number=national_insurance_number,
                    id_document_type=id_document_type,
                    id_number=id_number,
                )
            )
    return tenants


def apply_ocr_prefill(tenant_key: str) -> None:
    extracted = st.session_state.pop(f"{tenant_key}_ocr_prefill", None)
    if not extracted:
        return
    if extracted.get("full_name"):
        st.session_state[f"{tenant_key}_name"] = extracted["full_name"]
    if extracted.get("date_of_birth"):
        st.session_state[f"{tenant_key}_dob"] = extracted["date_of_birth"]
    if extracted.get("id_number"):
        st.session_state[f"{tenant_key}_id_number"] = extracted["id_number"]


def parse_ui_date(raw_value: str) -> date:
    value = raw_value.strip()
    try:
        return datetime.strptime(value, "%d/%m/%Y").date()
    except ValueError as exc:
        raise ValueError(f"Use the date format DD/MM/YYYY. Invalid value: {value}") from exc


if __name__ == "__main__":
    main()
