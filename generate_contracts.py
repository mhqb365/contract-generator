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


def convert_docx_to_pdfs_with_libreoffice(generated_docs, log=print):
    """Export DOCX files to PDF using LibreOffice in headless mode."""
    soffice = find_libreoffice()
    if not soffice:
        return None

    log("Dang xuat PDF bang LibreOffice headless...")
    pdf_count = 0
    for docx_path in generated_docs:
        try:
            output_dir = os.path.dirname(docx_path)
            result = subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", output_dir, docx_path],
                check=False,
                capture_output=True,
                text=True,
            )
            if result.returncode != 0:
                log(f"   ⚠️ Loi xuat PDF LibreOffice '{os.path.basename(docx_path)}': {result.stderr.strip()}")
                continue
                
            pdf_path = docx_path.replace(".docx", ".pdf")
            if os.path.exists(pdf_path):
                log(f"   PDF:  {os.path.basename(pdf_path)}")
                pdf_count += 1
        except Exception as e:
            log(f"   Loi xuat PDF bang LibreOffice '{os.path.basename(docx_path)}': {e}")

    if pdf_count > 0:
        log(f"\nHoan tat! Da xuat {pdf_count} file PDF.")
        return True
    return False


def convert_docx_to_pdfs_with_docx2pdf(generated_docs, log=print):
    """Export DOCX files to PDF using docx2pdf, which controls Microsoft Word."""
    try:
        # pyrefly: ignore [missing-import]
        from docx2pdf import convert as docx2pdf_convert
    except ImportError:
        log("Khong tim thay docx2pdf.")
        return None

    log("Dang xuat PDF bang Microsoft Word/docx2pdf...")
    pdf_count = 0
    for docx_path in generated_docs:
        pdf_path = docx_path.replace(".docx", ".pdf")
        try:
            docx2pdf_convert(docx_path, pdf_path)
            log(f"   PDF:  {os.path.basename(pdf_path)}")
            pdf_count += 1
        except Exception as e:
            log(f"   Loi xuat PDF bang Microsoft Word/macOS '{os.path.basename(docx_path)}': {e}")
    if pdf_count:
        log(f"\nHoan tat! Da xuat {pdf_count} file PDF.")
        return True

    return False


def convert_docx_to_pdfs(generated_docs, log=print):
    """Export generated DOCX files to PDF using the best available platform tool."""
    if not generated_docs:
        log("\nKhong co file Word nao de xuat PDF.")
        return True

    log("\nBat dau xuat PDF (co the mat vai phut)...")

    if os.name != "nt":
        libreoffice_result = convert_docx_to_pdfs_with_libreoffice(generated_docs, log)
        if libreoffice_result is not None:
            return libreoffice_result

        log("Khong tim thay LibreOffice. Se thu Microsoft Word/docx2pdf.")
        docx2pdf_result = convert_docx_to_pdfs_with_docx2pdf(generated_docs, log)
        if docx2pdf_result is not None:
            return docx2pdf_result

        log("\nKhong the xuat PDF tren may nay.")
        log("   macOS nen cai LibreOffice de xuat PDF headless, khong can mo Microsoft Word.")
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
                pdf_path = docx_path.replace(".docx", ".pdf")
                try:
                    doc_obj = word.Documents.Open(str(pathlib.Path(docx_path).absolute()))
                    # 17 = wdFormatPDF
                    doc_obj.SaveAs(str(pathlib.Path(pdf_path).absolute()), FileFormat=17)
                    doc_obj.Close(0)  # 0 = wdDoNotSaveChanges
                    pdf_name = os.path.basename(pdf_path)
                    log(f"   PDF:  {pdf_name}")
                    pdf_count += 1
                except Exception as e:
                    log(f"   Loi xuat PDF '{os.path.basename(docx_path)}': {e}")
            log(f"\nHoan tat! Da xuat {pdf_count} file PDF.")
            return True
        except Exception as e:
            log(f"\nLoi khi khoi dong Word de xuat PDF: {e}")
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

    docx2pdf_result = convert_docx_to_pdfs_with_docx2pdf(generated_docs, log)
    if docx2pdf_result is not None:
        return docx2pdf_result

    libreoffice_result = convert_docx_to_pdfs_with_libreoffice(generated_docs, log)
    if libreoffice_result is not None:
        return libreoffice_result

    log("\nKhong the xuat PDF tren may nay.")
    log("   Hay cai Microsoft Word/docx2pdf hoac LibreOffice de xuat PDF.")
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
        log("Khong the tu dong chuyen doi file .doc tren he dieu hanh nay. Vui long dung template .docx.")
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


def get_responsible_persons(excel_path, log=print):
    """
    Get list of unique responsible persons (Column I) and their statistics.
    :param excel_path: Path to the Excel data file.
    :param log: Callback function for logging (default: print).
    :return: Dictionary with responsible persons and their statistics
    """
    log("📂 Đang đọc dữ liệu từ Excel...")
    try:
        df = pd.read_excel(excel_path, header=None)
    except Exception as e:
        log(f"❌ Lỗi đọc file Excel: {e}")
        return {}

    # Auto-detect header row by checking column B (index 1)
    start_idx = 1

    persons_data = {}
    
    for index in range(start_idx, len(df)):
        row = df.iloc[index]
        
        def get_val(col_idx):
            if col_idx >= len(row): return ""
            val = row[col_idx]
            return str(val).strip() if pd.notna(val) else ""

        ten_cong_ty = get_val(3)
        nguoi_phu_trach = get_val(10)

        if not ten_cong_ty or not nguoi_phu_trach:
            continue

        if nguoi_phu_trach not in persons_data:
            persons_data[nguoi_phu_trach] = []
        
        persons_data[nguoi_phu_trach].append(ten_cong_ty)

    return persons_data


def generate(excel_path, log=print, output_format="both", responsible_person=None):
    """
    Main contract generation logic.
    :param excel_path: Path to the Excel data file.
    :param log: Callback function for logging (default: print).
    :param output_format: "docx", "pdf" or "both".
    :param responsible_person: Filter by responsible person (None = all).
    """
    script_dir = pathlib.Path(__file__).parent
    template_path = str(script_dir / "templates" / "HDNT.doc")
    base_output_dir = str(script_dir / "contracts")

    if not os.path.exists(base_output_dir):
        os.makedirs(base_output_dir)

    log("📂 Đang đọc dữ liệu từ Excel...")
    try:
        # header=None to read all rows including the first, avoiding data loss
        df = pd.read_excel(excel_path, header=None)
    except Exception as e:
        log(f"❌ Lỗi đọc file Excel: {e}")
        log("   Hãy đảm bảo file không đang được mở trong Excel.")
        return False

    log(f"   Đọc được {len(df)} dòng dữ liệu.")

    # Convert .doc template to .docx
    actual_template_path = get_docx_path(template_path, log)
    if not actual_template_path.endswith('.docx'):
        log("❌ CẢNH BÁO: Không thể tự động chuyển đổi file .doc.")
        log("   Vui lòng mở file templates/HDNT.doc bằng Microsoft Word, nhấn 'Save As' -> .docx")
        return False

    log("📝 Bắt đầu tạo hợp đồng...")
    count = 0
    generated_docs = []

    # Bỏ qua dòng 1 (dòng tiêu đề)
    start_idx = 1

    for index in range(start_idx, len(df)):
        row = df.iloc[index]

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

        try:
            doc = DocxTemplate(actual_template_path)
        except Exception as e:
            log(f"❌ Lỗi nạp template: {e}")
            break

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

        doc.render(context)

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

        safe_company_name = sanitize_name(ten_cong_ty)
        safe_person_name = sanitize_name(nguoi_phu_trach) if nguoi_phu_trach else "Unknown"

        # Create folder structure: contracts/HDNT [người phụ trách]/[công ty]/
        person_dir = os.path.join(base_output_dir, f"HDNT {safe_person_name}")
        company_dir = os.path.join(person_dir, safe_company_name)
        
        if not os.path.exists(company_dir):
            os.makedirs(company_dir)

        output_filepath = os.path.join(company_dir, f"{safe_company_name}.docx")

        try:
            doc.save(output_filepath)
            generated_docs.append(output_filepath)
            log(f"   ✅ DOCX: HDNT {safe_person_name}/{safe_company_name}.docx")
            count += 1
        except Exception as e:
            log(f"   ⚠️ Lỗi lưu file '{safe_company_name}': {e}")

    log(f"\n✅ Đã tạo {count} file Word trong thư mục 'contracts'.")

    should_convert_pdf = output_format in ("pdf", "both")
    if not should_convert_pdf:
        return True

    if count == 0:
        log("❌ Không có file DOCX để chuyển đổi sang PDF.")
        return False

    if output_format == "pdf":
        # When user explicitly requests PDF, conversion failure should be reported.
        return convert_docx_to_pdfs(generated_docs, log)

    # output_format == "both"
    if os.name == 'nt':
        return convert_docx_to_pdfs(generated_docs, log)
    else:
        if find_libreoffice():
            return convert_docx_to_pdfs(generated_docs, log)
        else:
            log("\n⚠️ LibreOffice không được cài đặt. Bỏ qua chuyển đổi PDF.")
            log("   Để convert PDF trên macOS, cài đặt LibreOffice:")
            log("   https://www.libreoffice.org/download/download/")
            return True  # Still return True vì DOCX đã được tạo thành công


def main():
    """CLI entry point for running without GUI."""
    excel_path = "Thông tin làm HỢP ĐỒNG NGUYÊN TẮC.xlsx"
    generate(excel_path, log=print)


if __name__ == "__main__":
    main()
