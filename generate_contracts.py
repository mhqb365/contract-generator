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
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", output_dir, docx_path],
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            pdf_path = docx_path.replace(".docx", ".pdf")
            if os.path.exists(pdf_path):
                log(f"   PDF:  {os.path.basename(pdf_path)}")
                pdf_count += 1
        except Exception as e:
            log(f"   Loi xuat PDF bang LibreOffice '{os.path.basename(docx_path)}': {e}")

    log(f"\nHoan tat! Da xuat {pdf_count} file PDF.")
    return pdf_count > 0


def convert_docx_to_pdfs_with_docx2pdf(generated_docs, log=print):
    """Export DOCX files to PDF using docx2pdf, which controls Microsoft Word."""
    try:
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


def generate(excel_path, log=print):
    """
    Main contract generation logic.
    :param excel_path: Path to the Excel data file.
    :param log: Callback function for logging (default: print).
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

    now = datetime.now()
    current_date_str = f"ngày {now.strftime('%d')} tháng {now.strftime('%m')} năm {now.strftime('%Y')}"

    log("📝 Bắt đầu tạo hợp đồng...")
    count = 0
    generated_docs = []

    # Auto-detect header row by checking column B (index 1)
    start_idx = 0
    if len(df) > 0:
        val_b = str(df.iloc[0, 1]).lower()
        if 'công ty' in val_b or 'company' in val_b or 'tên' in val_b:
            start_idx = 1  # Skip header row

    for index in range(start_idx, len(df)):
        row = df.iloc[index]

        # Helper: get cell value, return empty string if nan/null
        def get_val(col_idx):
            if col_idx >= len(row): return ""
            val = row[col_idx]
            return str(val).strip() if pd.notna(val) else ""

        # Column mapping (0-indexed): A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9, K=10
        ten_cong_ty   = get_val(1)
        ma_so_thue    = get_val(2)
        dia_chi       = get_val(3)
        nguoi_dai_dien = get_val(9)
        chuc_vu       = get_val(10)

        # Account number: combine column F (5) and G (6)
        val_f = get_val(5)
        val_g = get_val(6)
        if val_f and val_g:
            so_tai_khoan = f"{val_f} Mở tại {val_g}"
        else:
            so_tai_khoan = val_f or val_g

        # Skip rows with no company name
        if not ten_cong_ty:
            continue

        try:
            doc = DocxTemplate(actual_template_path)
        except Exception as e:
            log(f"❌ Lỗi nạp template: {e}")
            break

        context = {
            "thoi_gian_tao":   current_date_str,
            "ten_cong_ty":     ten_cong_ty,
            "ma_so_thue":      ma_so_thue,
            "dia_chi":         dia_chi,
            "so_tai_khoan":    so_tai_khoan,
            "nguoi_dai_dien":  nguoi_dai_dien,
            "chuc_vu":         chuc_vu,
        }

        doc.render(context)

        # Safe folder/file name (remove special characters)
        safe_name = "".join([c for c in ten_cong_ty if c.isalnum() or c in (' ', '-', '_', '.', ',')]).strip()

        company_dir = os.path.join(base_output_dir, safe_name)
        if not os.path.exists(company_dir):
            os.makedirs(company_dir)

        output_filepath = os.path.join(company_dir, f"{safe_name}.docx")

        try:
            doc.save(output_filepath)
            generated_docs.append(output_filepath)
            log(f"   ✅ DOCX: {safe_name}.docx")
            count += 1
        except Exception as e:
            log(f"   ⚠️ Lỗi lưu file '{safe_name}': {e}")

    log(f"\n✅ Đã tạo {count} file Word trong thư mục 'contracts'.")

    return convert_docx_to_pdfs(generated_docs, log)


def main():
    """CLI entry point for running without GUI."""
    excel_path = "Thông tin làm HỢP ĐỒNG NGUYÊN TẮC.xlsx"
    generate(excel_path, log=print)


if __name__ == "__main__":
    main()
