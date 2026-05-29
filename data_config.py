import json
import pathlib
import re


CONFIG_PATH = pathlib.Path(__file__).parent / "data_fields.json"
DEFAULT_TEMPLATE_PATH = "templates/HDNT.docx"
DEFAULT_PDF_ENGINE = "libreoffice"
PDF_ENGINES = {"libreoffice", "word"}


DEFAULT_DATA_FIELDS = [
    {"template_name": "so_hd", "excel_column": "B", "display_name": "Số HĐ", "role": "contract_number"},
    {"template_name": "ten_hd", "excel_column": "C", "display_name": "Tên HĐ", "role": "contract_name"},
    {"template_name": "ten_cong_ty", "excel_column": "D", "display_name": "Tên công ty", "role": "company"},
    {"template_name": "ma_so_thue", "excel_column": "E", "display_name": "Mã số thuế", "role": "tax_code"},
    {"template_name": "dia_chi", "excel_column": "F", "display_name": "Địa chỉ", "role": "address"},
    {"template_name": "so_tai_khoan", "excel_column": "H", "display_name": "Số tài khoản", "role": "account_number"},
    {"template_name": "ngan_hang", "excel_column": "I", "display_name": "Ngân hàng", "role": "bank"},
    {"template_name": "nguoi_phu_trach", "excel_column": "K", "display_name": "Sale", "role": "responsible_person"},
    {"template_name": "nguoi_dai_dien", "excel_column": "L", "display_name": "Đại diện", "role": "representative"},
    {"template_name": "chuc_vu", "excel_column": "M", "display_name": "Chức vụ", "role": "position"},
]


ROLE_FALLBACKS = {
    "company": "ten_cong_ty",
    "responsible_person": "nguoi_phu_trach",
    "account_number": "so_tai_khoan",
    "bank": "ngan_hang",
}


def normalize_template_name(value):
    value = (value or "").strip()
    if value.startswith("{{") and value.endswith("}}"):
        value = value[2:-2].strip()
    return re.sub(r"\s+", "_", value)


def normalize_excel_column(value):
    return re.sub(r"\s+", "", (value or "").upper())


def normalize_field(field, default_role=""):
    return {
        "template_name": normalize_template_name(field.get("template_name")),
        "excel_column": normalize_excel_column(field.get("excel_column")),
        "display_name": (field.get("display_name") or "").strip(),
        "role": field.get("role", default_role) or "",
    }


def default_fields():
    return [dict(field) for field in DEFAULT_DATA_FIELDS]


def default_template_path():
    return DEFAULT_TEMPLATE_PATH


def default_pdf_engine():
    return DEFAULT_PDF_ENGINE


def _read_config_data():
    if not CONFIG_PATH.exists():
        return {}
    data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    return data if isinstance(data, dict) else {"fields": data}


def _normalize_template_path(template_path):
    return (template_path or DEFAULT_TEMPLATE_PATH).strip() or DEFAULT_TEMPLATE_PATH


def _normalize_pdf_engine(pdf_engine):
    value = (pdf_engine or DEFAULT_PDF_ENGINE).strip().lower()
    return value if value in PDF_ENGINES else DEFAULT_PDF_ENGINE


def save_data_settings(fields, template_path=None, pdf_engine=None):
    cleaned = []
    seen_names = set()
    for field in fields:
        normalized = normalize_field(field)
        if not normalized["template_name"] or not normalized["excel_column"]:
            continue
        if normalized["template_name"] in seen_names:
            continue
        seen_names.add(normalized["template_name"])
        if not normalized["display_name"]:
            normalized["display_name"] = normalized["template_name"]
        cleaned.append(normalized)

    if not cleaned:
        cleaned = default_fields()

    if template_path is None:
        try:
            data = _read_config_data()
            template_path = data.get("template_path", DEFAULT_TEMPLATE_PATH)
        except Exception:
            template_path = DEFAULT_TEMPLATE_PATH

    if pdf_engine is None:
        try:
            data = _read_config_data()
            pdf_engine = data.get("pdf_engine", DEFAULT_PDF_ENGINE)
        except Exception:
            pdf_engine = DEFAULT_PDF_ENGINE

    CONFIG_PATH.write_text(
        json.dumps(
            {
                "template_path": _normalize_template_path(template_path),
                "pdf_engine": _normalize_pdf_engine(pdf_engine),
                "fields": cleaned,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return cleaned


def save_data_fields(fields):
    return save_data_settings(fields)


def load_data_settings():
    if not CONFIG_PATH.exists():
        fields = save_data_settings(default_fields(), DEFAULT_TEMPLATE_PATH, DEFAULT_PDF_ENGINE)
        return {"template_path": DEFAULT_TEMPLATE_PATH, "pdf_engine": DEFAULT_PDF_ENGINE, "fields": fields}

    try:
        data = _read_config_data()
        raw_fields = data.get("fields", data if isinstance(data, list) else [])
        fields = []
        defaults_by_template = {field["template_name"]: field for field in DEFAULT_DATA_FIELDS}
        for raw in raw_fields:
            if not isinstance(raw, dict):
                continue
            normalized = normalize_field(raw)
            if not normalized["role"]:
                default = defaults_by_template.get(normalized["template_name"], {})
                normalized["role"] = default.get("role", "")
            if normalized["template_name"] and normalized["excel_column"]:
                normalized["display_name"] = normalized["display_name"] or normalized["template_name"]
                fields.append(normalized)
        if not fields:
            fields = save_data_settings(default_fields(), data.get("template_path", DEFAULT_TEMPLATE_PATH))
        return {
            "template_path": _normalize_template_path(data.get("template_path", DEFAULT_TEMPLATE_PATH)),
            "pdf_engine": _normalize_pdf_engine(data.get("pdf_engine", DEFAULT_PDF_ENGINE)),
            "fields": fields,
        }
    except Exception:
        fields = save_data_settings(default_fields(), DEFAULT_TEMPLATE_PATH, DEFAULT_PDF_ENGINE)
        return {"template_path": DEFAULT_TEMPLATE_PATH, "pdf_engine": DEFAULT_PDF_ENGINE, "fields": fields}


def load_data_fields():
    return load_data_settings()["fields"]


def load_template_path():
    return load_data_settings()["template_path"]


def load_pdf_engine():
    return load_data_settings()["pdf_engine"]


def column_letter_to_index(column):
    column = normalize_excel_column(column)
    if not column or not column.isalpha():
        raise ValueError(f"Cot Excel khong hop le: {column}")

    total = 0
    for char in column:
        total = total * 26 + (ord(char) - ord("A") + 1)
    return total - 1


def column_indexes(value):
    parts = [part for part in re.split(r"[,;+]", normalize_excel_column(value)) if part]
    return [column_letter_to_index(part) for part in parts]


def get_cell_value(row, col_idx):
    if col_idx >= len(row):
        return ""
    val = row[col_idx]
    try:
        import pandas as pd

        return str(val).strip() if pd.notna(val) else ""
    except Exception:
        return "" if val is None else str(val).strip()


def get_field_value(row, field):
    values = [get_cell_value(row, idx) for idx in column_indexes(field["excel_column"])]
    values = [value for value in values if value]
    return " ".join(values)


def make_row_context(row, fields):
    context = {}
    for field in fields:
        context[field["template_name"]] = get_field_value(row, field)

    account = get_role_value(context, fields, "account_number")
    bank = get_role_value(context, fields, "bank")
    if "so_tai_khoan_ngan_hang" not in context:
        if account and bank:
            context["so_tai_khoan_ngan_hang"] = f"{account} Mở tại {bank}"
        else:
            context["so_tai_khoan_ngan_hang"] = account or bank

    return context


def get_role_field(fields, role):
    fallback_name = ROLE_FALLBACKS.get(role)
    for field in fields:
        if field.get("role") == role:
            return field
    if fallback_name:
        for field in fields:
            if field.get("template_name") == fallback_name:
                return field
    return None


def get_role_value(context, fields, role):
    field = get_role_field(fields, role)
    if not field:
        return ""
    return context.get(field["template_name"], "")
