import os
import shutil
import subprocess
import sys
if sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
import pathlib
import unicodedata
from data_config import get_role_value, load_data_fields, load_template_path, make_row_context

try:
    import win32com.client as win32
except ImportError:
    win32 = None


def find_libreoffice():
    """Return a LibreOffice/soffice executable path when available."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        return soffice

    if sys.platform == "darwin":
        macos_soffice = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if os.path.exists(macos_soffice):
            return macos_soffice

    return None


def _is_cancelled(cancel_event):
    return bool(cancel_event and cancel_event.is_set())


def convert_docx_to_pdfs_with_libreoffice(generated_docs, log=print, cancel_event=None):
    """Export DOCX files to PDF using LibreOffice in headless mode."""
    soffice = find_libreoffice()
    if not soffice:
        return None

    log("Đan xuất PDF bằng LibreOffice headless...")
    pdf_count = 0
    for docx_path in generated_docs:
        if _is_cancelled(cancel_event):
            log("Đã hủy xuất PDF.")
            return "cancelled"
        try:
            output_dir = os.path.dirname(docx_path)
            result = subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", output_dir, docx_path],
                check=False,
                capture_output=True,
                text=True,
            )
            if result.returncode != 0:
                log(f"Lỗi xuất PDF LibreOffice '{os.path.basename(docx_path)}': {result.stderr.strip()}")
                continue
                
            pdf_path = docx_path.replace(".docx", ".pdf")
            if os.path.exists(pdf_path):
                log(f"PDF: {os.path.basename(pdf_path)}")
                pdf_count += 1
        except Exception as e:
            log(f"Lỗi xuất PDF bằng LibreOffice '{os.path.basename(docx_path)}': {e}")

    if pdf_count > 0:
        log(f"\nHoàn tất! Đã xuất {pdf_count} file PDF.")
        return True
    return False


def convert_docx_to_pdfs_with_docx2pdf(generated_docs, log=print, cancel_event=None):
    """Export DOCX files to PDF using docx2pdf, which controls Microsoft Word."""
    try:
        # pyrefly: ignore [missing-import]
        from docx2pdf import convert as docx2pdf_convert
    except ImportError:
        log("Không tìm thấy docx2pdf.")
        return None

    log("Đang xuất PDF bằng Microsoft Word/docx2pdf...")
    pdf_count = 0
    for docx_path in generated_docs:
        if _is_cancelled(cancel_event):
            log("Đã hủy xuất PDF.")
            return "cancelled"
        pdf_path = docx_path.replace(".docx", ".pdf")
        try:
            docx2pdf_convert(docx_path, pdf_path)
            log(f"PDF:  {os.path.basename(pdf_path)}")
            pdf_count += 1
        except Exception as e:
            log(f"Lỗi xuất PDF bằng Microsoft Word/macOS '{os.path.basename(docx_path)}': {e}")
    if pdf_count:
        log(f"\nHoàn tất! Đã xuất {pdf_count} file PDF.")
        return True

    return False


def convert_docx_to_pdfs(generated_docs, log=print, cancel_event=None):
    """Export generated DOCX files to PDF using the best available platform tool."""
    if not generated_docs:
        log("\nKhông có file Word nào để xuất PDF.")
        return True

    if _is_cancelled(cancel_event):
        log("Đã hủy trước khi xuất PDF.")
        return "cancelled"

    log("\nBắt đầu xuất PDF (có thể mất vài phút)...")

    if os.name != "nt":
        libreoffice_result = convert_docx_to_pdfs_with_libreoffice(generated_docs, log, cancel_event)
        if libreoffice_result is not None:
            return libreoffice_result

        log("Không tìm thấy LibreOffice. sẽ thử Microsoft Word/docx2pdf.")
        docx2pdf_result = convert_docx_to_pdfs_with_docx2pdf(generated_docs, log, cancel_event)
        if docx2pdf_result is not None:
            return docx2pdf_result

        log("\nKhông thể xuất PDF trên máy này.")
        log("macOS nen cai LibreOffice de xuat PDF headless, không cần mở Microsoft Word.")
        return False

    if win32 is not None:
        word = None
        com_initialized = False
        try:
            try:
                import pythoncom

                pythoncom.CoInitialize()
                com_initialized = True
            except ImportError:
                pythoncom = None

            word = win32.Dispatch('Word.Application')
            word.Visible = False
            pdf_count = 0
            for docx_path in generated_docs:
                if _is_cancelled(cancel_event):
                    log("Đã hủy xuất PDF.")
                    return "cancelled"
                pdf_path = docx_path.replace(".docx", ".pdf")
                try:
                    doc_obj = word.Documents.Open(str(pathlib.Path(docx_path).absolute()))
                    # 17 = wdFormatPDF
                    doc_obj.SaveAs(str(pathlib.Path(pdf_path).absolute()), FileFormat=17)
                    doc_obj.Close(0)  # 0 = wdDoNotSaveChanges
                    pdf_name = os.path.basename(pdf_path)
                    log(f"PDF:  {pdf_name}")
                    pdf_count += 1
                except Exception as e:
                    log(f"Lỗi xuất PDF '{os.path.basename(docx_path)}': {e}")
            log(f"\nHoàn tất! Đã xuất {pdf_count} file PDF.")
            return True
        except Exception as e:
            log(f"\nLỗi khi khởi động Word để xuất PDF: {e}")
            return False
        finally:
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    docx2pdf_result = convert_docx_to_pdfs_with_docx2pdf(generated_docs, log, cancel_event)
    if docx2pdf_result is not None:
        return docx2pdf_result

    libreoffice_result = convert_docx_to_pdfs_with_libreoffice(generated_docs, log, cancel_event)
    if libreoffice_result is not None:
        return libreoffice_result

    log("\nKhông thể xuất PDF trên máy này.")
    log("Hãy cài Microsoft Word/docx2pdf hoặc LibreOffice để xuất PDF.")
    return False


def get_docx_path(template_path, log=print):
    """
    Check and convert .doc to .docx if needed,
    since docxtpl only works with .docx format.
    """
    if template_path.endswith('.docx'):
        return template_path

    docx_path = template_path + "x"
    if os.path.exists(docx_path):
        return docx_path

    if win32 is None:
        log("Không thể tự động chuyển đổi file .doc trên hệ điều hành này. Vui lòng dùng template .docx.")
        return template_path

    log(f"Đang tự động chuyển đổi '{template_path}' sang .docx ...")
    word = None
    doc = None
    com_initialized = False
    try:
        try:
            import pythoncom

            pythoncom.CoInitialize()
            com_initialized = True
        except ImportError:
            pythoncom = None

        word = win32.Dispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(str(pathlib.Path(template_path).absolute()))
        # wdFormatDocumentDefault = 16
        doc.SaveAs(str(pathlib.Path(docx_path).absolute()), FileFormat=16)
        return docx_path
    except Exception as e:
        log(f"Lỗi tự động chuyển đổi file .doc (cần Microsoft Word): {e}")
        return template_path
    finally:
        if doc is not None:
            try:
                doc.Close(0)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def _get_cell_value(row, col_idx):
    if col_idx >= len(row):
        return ""
    val = row[col_idx]
    return str(val).strip() if pd.notna(val) else ""


def get_contract_preview(excel_path, log=print):
    """
    Read Excel rows and return contract data for GUI preview.
    row_index is the original DataFrame index, used later to generate selected rows.
    """
    log("Đang đọc dữ liệu từ Excel...")
    try:
        df = pd.read_excel(excel_path, header=None)
    except Exception as e:
        log(f"Lỗi đọc file Excel: {e}")
        return []

    data_fields = load_data_fields()
    preview_rows = []
    start_idx = 1

    for index in range(start_idx, len(df)):
        row = df.iloc[index]
        context = make_row_context(row, data_fields)
        ten_cong_ty = get_role_value(context, data_fields, "company")

        if not ten_cong_ty:
            continue

        preview_row = {
            "row_index": index,
            "_context": context,
            "_company": ten_cong_ty,
            "_responsible_person": get_role_value(context, data_fields, "responsible_person"),
        }
        preview_row.update(context)
        preview_rows.append(preview_row)

    return preview_rows


def get_responsible_persons(excel_path, log=print):
    """
    Get list of unique responsible persons (Column I) and their statistics.
    :param excel_path: Path to the Excel data file.
    :param log: Callback function for logging (default: print).
    :return: Dictionary with responsible persons and their statistics
    """
    log("Đang đọc dữ liệu từ Excel...")
    try:
        df = pd.read_excel(excel_path, header=None)
    except Exception as e:
        log(f"Lỗi đọc file Excel: {e}")
        return {}

    start_idx = 1
    data_fields = load_data_fields()

    persons_data = {}
    
    for index in range(start_idx, len(df)):
        row = df.iloc[index]
        
        context = make_row_context(row, data_fields)
        ten_cong_ty = get_role_value(context, data_fields, "company")
        nguoi_phu_trach = get_role_value(context, data_fields, "responsible_person")

        if not ten_cong_ty or not nguoi_phu_trach:
            continue

        if nguoi_phu_trach not in persons_data:
            persons_data[nguoi_phu_trach] = []
        
        persons_data[nguoi_phu_trach].append(ten_cong_ty)

    return persons_data


def generate(excel_path, log=print, output_format="both", responsible_person=None, selected_rows=None, template_path=None, cancel_event=None):
    """
    Main contract generation logic.
    :param excel_path: Path to the Excel data file.
    :param log: Callback function for logging (default: print).
    :param output_format: "docx", "pdf" or "both".
    :param responsible_person: Filter by responsible person (None = all).
    :param selected_rows: Optional iterable of DataFrame row indexes selected in GUI.
    :param template_path: Optional Word template path. Uses saved settings when omitted.
    :param cancel_event: Optional threading.Event used to cancel long runs.
    """
    # Safe folder/file name (remove Vietnamese diacritics and special characters completely)
    # Normalize to NFKD form and remove combining marks (diacritics)
    def sanitize_name(name):
        if not name:
            return ""
        # Manual replacement for Đ/đ as NFKD doesn't convert it to D/d
        name = name.replace('đ', 'd').replace('Đ', 'D')
        normalized = unicodedata.normalize('NFKD', name)
        no_diacritics = ''.join(c for c in normalized if unicodedata.category(c) != 'Mn')
        sanitized = "".join([c for c in no_diacritics if c.isalnum() or c in (' ', '-', '_')]).strip()
        return " ".join(sanitized.split())

    script_dir = pathlib.Path(__file__).parent
    template_path = template_path or load_template_path()
    template_path_obj = pathlib.Path(template_path)
    if not template_path_obj.is_absolute():
        template_path_obj = script_dir / template_path_obj
    template_path = str(template_path_obj)
    base_output_dir = str(script_dir / "contracts")

    if _is_cancelled(cancel_event):
        log("Đã hủy trước khi bắt đầu.")
        return "cancelled"

    if not os.path.exists(base_output_dir):
        os.makedirs(base_output_dir)

    log("Đang đọc dữ liệu từ Excel...")
    try:
        # header=None to read all rows including the first, avoiding data loss
        df = pd.read_excel(excel_path, header=None)
    except Exception as e:
        log(f"Lỗi đọc file Excel: {e}")
        log("Hãy đảm bảo file không đang được mở trong Excel.")
        return False

    if not os.path.exists(template_path):
        log(f"Không tìm thấy file template: {template_path}")
        return False

    # Convert .doc template to .docx
    actual_template_path = get_docx_path(template_path, log)
    if not actual_template_path.endswith('.docx'):
        log("CẢNH BÁO: Không thể tự động chuyển đổi file .doc.")
        log("Vui lòng mở file templates/HDNT.doc bằng Microsoft Word, nhấn 'Save As' -> .docx")
        return False

    log("Bắt đầu tạo hợp đồng...")
    count = 0
    generated_docs = []
    selected_row_set = set(selected_rows) if selected_rows is not None else None
    data_fields = load_data_fields()

    # Bỏ qua dòng 1 (dòng tiêu đề)
    start_idx = 1

    for index in range(start_idx, len(df)):
        if _is_cancelled(cancel_event):
            log("Đã hủy tạo hợp đồng.")
            return "cancelled"

        row = df.iloc[index]

        if selected_row_set is not None and index not in selected_row_set:
            continue

        # Helper: get cell value, return empty string if nan/null
        def get_val(col_idx):
            if col_idx >= len(row): return ""
            val = row[col_idx]
            return str(val).strip() if pd.notna(val) else ""

        # Column mapping (0-indexed): 1=Số HĐ, 2=Tên HĐ, 3=Tên Công Ty, 4=Mã Số Thuế, 5=Địa Chỉ, 7=Số Tài Khoản, 8=Ngân Hàng, 10=Người Phụ Trách, 11=Người Đại Diện, 12=Chức Vụ
        so_hd = get_val(1)
        ten_hd = get_val(2)
        ten_cong_ty = get_val(3)
        ma_so_thue = get_val(4)
        dia_chi = get_val(5)
        so_tai_khoan = get_val(7)
        ngan_hang = get_val(8)
        nguoi_phu_trach = get_val(10)
        nguoi_dai_dien = get_val(11)
        chuc_vu = get_val(12)
        context = make_row_context(row, data_fields)
        ten_cong_ty = get_role_value(context, data_fields, "company")
        nguoi_phu_trach = get_role_value(context, data_fields, "responsible_person")

        # Skip rows with no company name
        if not ten_cong_ty:
            continue

        # Skip rows if filtering by responsible person
        if responsible_person and nguoi_phu_trach != responsible_person:
            continue

        # Combine account number and bank
        if so_tai_khoan and ngan_hang:
            so_tai_khoan_ngan_hang = f"{so_tai_khoan} Mở tại {ngan_hang}"
        else:
            so_tai_khoan_ngan_hang = so_tai_khoan or ngan_hang

        safe_company_name = sanitize_name(ten_cong_ty)
        safe_person_name = sanitize_name(nguoi_phu_trach) if nguoi_phu_trach else "Unknown"

        # Create folder structure: contracts/HDNT [người phụ trách]/[công ty]/
        person_dir = os.path.join(base_output_dir, f"HDNT {safe_person_name}")
        company_dir = os.path.join(person_dir, safe_company_name)
        output_filepath = os.path.join(company_dir, f"{safe_company_name}.docx")

        if output_format == "pdf":
            if os.path.exists(output_filepath):
                generated_docs.append(output_filepath)
                count += 1
            else:
                log(f"Chưa có file Word: HDNT {safe_person_name}/{safe_company_name}.docx (Hãy tạo Docx trước)")
            continue

        # If format is docx or both, generate the docx file
        try:
            doc = DocxTemplate(actual_template_path)
        except Exception as e:
            log(f"Lỗi nạp template: {e}")
            break

        if _is_cancelled(cancel_event):
            log("Đã hủy tạo hợp đồng.")
            return "cancelled"

        context = {
            "so_hd":                  so_hd,
            "ten_hd":                 ten_hd,
            "ten_cong_ty":            ten_cong_ty,
            "ma_so_thue":             ma_so_thue,
            "dia_chi":                dia_chi,
            "so_tai_khoan_ngan_hang": so_tai_khoan_ngan_hang,
            "nguoi_dai_dien":         nguoi_dai_dien,
            "chuc_vu":                chuc_vu
        }

        context = make_row_context(row, data_fields)
        doc.render(context)

        if not os.path.exists(company_dir):
            os.makedirs(company_dir)

        try:
            doc.save(output_filepath)
            generated_docs.append(output_filepath)
            log(f"DOCX: HDNT {safe_person_name}/{safe_company_name}.docx")
            count += 1
        except Exception as e:
            log(f"Lỗi lưu file '{safe_company_name}': {e}")

    if output_format == "pdf":
        log(f"\nĐã tìm thấy {count} file Word sẵn sàng để chuyển đổi sang PDF.")
    else:
        log(f"\nĐã tạo {count} file Word trong thư mục 'contracts'.")

    should_convert_pdf = output_format in ("pdf", "both")
    if not should_convert_pdf:
        return True

    if _is_cancelled(cancel_event):
        log("Đã hủy trước khi xuất PDF.")
        return "cancelled"

    if count == 0:
        log("Không có file DOCX để chuyển đổi sang PDF.")
        return False

    if output_format == "pdf":
        # When user explicitly requests PDF, conversion failure should be reported.
        return convert_docx_to_pdfs(generated_docs, log, cancel_event)

    # output_format == "both"
    if os.name == 'nt':
        return convert_docx_to_pdfs(generated_docs, log, cancel_event)
    else:
        if find_libreoffice():
            return convert_docx_to_pdfs(generated_docs, log, cancel_event)
        else:
            log("\nLibreOffice không được cài đặt. Bỏ qua chuyển đổi PDF.")
            log("Để convert PDF trên macOS, cài đặt LibreOffice:")
            log("https://www.libreoffice.org/download/download/")
            return True  # Still return True vì DOCX đã được tạo thành công


def main():
    """CLI entry point for running without GUI."""
    excel_path = "Thông tin làm HỢP ĐỒNG NGUYÊN TẮC.xlsx"
    generate(excel_path, log=print)


if __name__ == "__main__":
    main()
