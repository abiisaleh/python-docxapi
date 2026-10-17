import io
import os
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Depends
from fastapi.security.api_key import APIKeyHeader
from fastapi.responses import StreamingResponse
from docx import Document

app = FastAPI(
    title="Word Document Generator API",
    description="API dengan keamanan API Key untuk mengupdate dokumen Word",
    version="1.1.0"
)

# Nama Header yang akan digunakan untuk API Key
API_KEY_NAME = "X-API-KEY"
api_key_header = APIKeyHeader(name=API_KEY_NAME, auto_error=False)

def get_api_key(api_key: str = Depends(api_key_header)):
    expected_api_key = os.getenv("API_KEY")

    # Validasi kecocokan key yang dikirim client dengan server
    if api_key == expected_api_key:
        return api_key

    raise HTTPException(
        status_code=403, 
        detail="Akses Ditolak: API Key tidak valid atau tidak disertakan."
    )

def replace_placeholder_in_paragraph(paragraph, placeholder, replacement) -> bool:
    """
    Fungsi cerdas untuk mengganti placeholder di paragraf Word tanpa merusak format 
    meskipun teks terpecah menjadi beberapa Run oleh MS Word.
    Mengembalikan True jika ada penggantian yang dilakukan, False jika tidak.
    """
    replaced = False

    # Jika tidak ada placeholder sama sekali, cepat keluar
    if placeholder not in paragraph.text:
        return False

    # Coba ganti di setiap run (mengatasi kasus placeholder berada di satu atau beberapa run)
    for run in paragraph.runs:
        if placeholder in run.text:
            run.text = run.text.replace(placeholder, replacement)
            replaced = True

    if replaced:
        return True

    # Jika teks terpecah di beberapa run sehingga tidak ditemukan pada run tunggal,
    # gabungkan seluruh teks, ganti, lalu tulis ulang ke run pertama dan kosongkan yang lain.
    full_text = paragraph.text
    new_text = full_text.replace(placeholder, replacement)

    if paragraph.runs:
        paragraph.runs[0].text = new_text
        for run in paragraph.runs[1:]:
            run.text = ""
        return True

    return False


def process_docx(file_bytes: bytes, nomor_surat: str) -> io.BytesIO:
    try:
        doc = Document(io.BytesIO(file_bytes))

        # PASTIKAN tulisan di file Word Anda tepat sama dengan isi variabel ini (menggunakan huruf kecil/besar yang sama)
        placeholder = "{nomor}"

        any_replacement = False

        # 1. Proses Paragraf Biasa
        for paragraph in doc.paragraphs:
            if replace_placeholder_in_paragraph(paragraph, placeholder, nomor_surat):
                any_replacement = True

        # 2. Proses Paragraf di Dalam Tabel
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if replace_placeholder_in_paragraph(paragraph, placeholder, nomor_surat):
                            any_replacement = True

        if not any_replacement:
            # Jika tidak ada penggantian sama sekali, berikan pesan error yang jelas
            raise HTTPException(status_code=400, detail=f"Placeholder {placeholder} tidak ditemukan dalam dokumen.")

        output_buffer = io.BytesIO()
        doc.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer
    except HTTPException:
        # Biarkan HTTPException (mis. 400 placeholder) keluar apa adanya
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal memproses dokumen: {str(e)}")


@app.post("/api/update-surat", summary="Update Nomor Surat dalam Dokumen Word")
async def update_surat(
    file: UploadFile = File(..., description="File template .docx yang diunggah"),
    nomor_surat: str = Form(..., description="Nomor surat baru untuk menggantikan {nomor}"),
    api_key: str = Depends(get_api_key)
):
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="Format file harus berupa .docx")
    
    file_bytes = await file.read()
    final_document = process_docx(file_bytes, nomor_surat)
    
    clean_filename = nomor_surat.replace("/", "-").replace(" ", "_")
    output_filename = f"Surat_{clean_filename}.docx"
    
    return StreamingResponse(
        final_document,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename={output_filename}"}
    )

@app.get("/")
def read_root():
    return {"status": "API Berjalan Lancar"}
