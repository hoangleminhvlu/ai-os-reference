# -*- coding: utf-8 -*-
"""TL2 bản 0.15 (20/11/2026): đối chiếu AI Platform 6.11.18 (iThunder-AI-OS PR #3, #4) và README của ai-workspace (PR #5).
- Trang bìa: phiên bản tài liệu, phiên bản phần mềm; mục "Những điểm mới trong bản 0.15".
- 4.1: tệp mẫu đúng tên ai_runtimes.json.example.
- 11.1: lệnh chạy thêm test_chunking, test_keyword_db, test_vector_db; Bảng 3 theo kết quả ngày 20/11/2026.
- Phụ lục C: OLLAMA_API_KEY vẫn được đọc (Ollama của Tổ chức), chỉ khóa của nhà cung cấp On-Clouds không còn đọc.
- Phụ lục G: căn cứ mã nguồn ngày 20/11/2026.
Đầu vào: TL2 bản 0.14 ĐÃ điền mục lục. Đầu ra: bản 0.15 chưa điền mục lục ở <thư mục ra>/<tên>.src.docx, sau đó chép vào
_tools/toc_work/ và chạy build_toc.py TL2.
    python update_v015.py <TL2_v014.docx> <thư mục ra> [<tệp kết quả test_vector_db>]"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docedit import Doc  # noqa: E402
from doc_versioning import new_section, unfill_toc  # noqa: E402

T2 = "iThunder_AI_OS_TL2_Cai_dat_va_Van_hanh_AI_Platform.docx"
DATE = "20/11/2026"

TEST_LINES = [
    "AIP=.venv/bin/python     # hoặc /usr/local/venv/bin/python3",
    "$AIP tests/identity/run_all.py",
    "$AIP tests/comm/test_comm.py",
    "$AIP tests/mcp/test_mcp.py",
    "$AIP tests/knowledge/test_chunking.py",
    "$AIP tests/knowledge/test_agent_tools.py",
    "$AIP tests/platform_status/test_platform_status.py          # --live: dịch vụ thật, chỉ đọc",
    "$AIP tests/zalo/test_zalo_mcp.py",
    "$AIP tests/doc_export/test_export_docx.py",
    "$AIP tests/cloud_ai/test_cloud_ai.py     # khóa API AI On-Clouds (CSDL, nhà cung cấp giả lập)",
    "$AIP tests/ai_server/test_ai_server.py",
    "# Cần PostgreSQL thật (đọc .env), dùng collection tạm và dọn sạch:",
    "$AIP tests/knowledge/test_archive_db.py",
    "$AIP tests/knowledge/test_keyword_db.py",
    "$AIP tests/knowledge/test_vector_db.py   # có HNSW: chèn 20.000 vector, chạy nhiều phút; đừng ngắt giữa chừng",
]


def tl2(src, dst, vector_result):
    doc = Doc(src)
    unfill_toc(doc)
    t = doc.table("Phiên bản tài liệu")
    doc.set_cell(t, doc.find_row(t, "Phiên bản tài liệu"), 1,
                 f"0.15, bản dự thảo ngày {DATE} (cập nhật bản 0.14: đối chiếu AI Platform 6.11.18, kết quả kiểm thử mới)")
    doc.set_cell(t, doc.find_row(t, "Phiên bản phần mềm"), 1,
                 f"AI Workspace 6.11.16 và AI Platform 6.11.18 (nhánh main ngày {DATE}, sau khi hợp nhất PR #3, PR #4 của "
                 "iThunder-AI-OS và PR #5 của ai-workspace)")

    new_section(doc, "Những điểm mới trong bản 0.14", "Những điểm mới trong bản 0.15",
                f"Đối chiếu hai kho mã nguồn ngày {DATE}: AI Platform ghi đúng phiên bản 6.11.18 trong mã nguồn "
                "(iThunder-AI-OS PR #4), README của hai kho được cập nhật (iThunder-AI-OS PR #3, ai-workspace PR #5). "
                "Các thay đổi và nơi mô tả:",
                "Bản 0.14 là bản phát hành công khai",
                "Gỡ phụ lục nội bộ dùng cho đội phát triển",
                [("Số phiên bản 6.11.18 trong aiplatform.py và config.py (trước đây ghi 6.11.02)", "Trang bìa, G"),
                 ("Tệp mẫu danh mục Máy chủ AI đúng tên ai_runtimes.json.example", "4.1"),
                 ("OLLAMA_API_KEY vẫn được đọc cho Ollama của Tổ chức; chỉ khóa của nhà cung cấp On-Clouds không còn "
                  "đọc từ .env", "C"),
                 (f"Kết quả kiểm thử ngày {DATE}, thêm test_chunking, test_keyword_db, test_vector_db", "11.1")])

    # ---- 4.1: tên tệp mẫu
    code = doc.find("cp ai_runtimes.json.sample", "tbl")
    doc.replace_in(code, "ai_runtimes.json.sample", "ai_runtimes.json.example")

    # ---- 11.1: lệnh và Bảng 3
    code = doc.table("$AIP tests/identity/run_all.py")
    doc.set_code(code, TEST_LINES)
    cap = doc.find("Bảng 3. Kết quả chạy trong môi trường cô lập", "p")
    doc.replace_in(cap, "Kết quả chạy trong môi trường cô lập, không kết nối dịch vụ thật (ngày 18/11/2026)",
                   f"Kết quả kiểm thử ngày {DATE}: dịch vụ giả lập, trừ ba bộ của Kho Tri thức chạy trên PostgreSQL thật")
    t = doc.table("Bộ kiểm thửKết quả ngày 18/11/2026")
    doc.set_cell(t, 0, 1, f"Kết quả ngày {DATE}")
    r = doc.find_row(t, "tests/knowledge/test_agent_tools.py")
    doc.add_row(t, ["tests/knowledge/test_chunking.py", "Đạt (20 kiểm tra)", "Tách đoạn có ngữ cảnh: DOCX, PDF, MD, TXT"],
                like=r, after=r - 1)
    r = doc.find_row(t, "tests/knowledge/test_archive_db.py")
    doc.set_cell(t, r, 0, "tests/knowledge/test_archive_db.py, test_keyword_db.py", code=None)
    doc.set_cell(t, r, 1, "Đạt (test_keyword_db: 20 kiểm tra)", code=False)
    doc.set_cell(t, r, 2, "PostgreSQL thật trên máy phát triển, collection tạm, dọn sạch sau khi chạy", code=False)
    doc.add_row(t, ["tests/knowledge/test_vector_db.py", vector_result[0], vector_result[1]], like=r, after=r)

    # ---- Phụ lục C: biến AI On-Clouds
    t = doc.table("NhómBiến trong .env")
    r = doc.find_row(t, "CLOUD_AI_SECRET")
    doc.set_cell(t, r, 1, "ANTHROPIC_BASE_URL, OPENAI_BASE_URL, GROK_BASE_URL, GEMINI_BASE_URL, DEEPSEEK_BASE_URL, "
                          "PERPLEXITY_BASE_URL, GLMAI_BASE_URL, LLAMA_BASE_URL; CLOUD_AI_SECRET, CLOUD_AI_ADMIN_ROLES, "
                          "CLOUD_AI_TIMEOUT, CLOUD_AI_MAX_TOKENS. Không còn đọc khóa API của nhà cung cấp On-Clouds "
                          "(ANTHROPIC_API_KEY, OPENAI_API_KEY...; khóa nhập trên giao diện, mục 5.6), nên xóa khỏi .env. "
                          "OLLAMA_API_KEY (nhóm Mô hình) vẫn được đọc cho Ollama của Tổ chức. OPEN_WEBUI_* chưa được mã "
                          "nguồn dùng")

    # ---- Phụ lục G: căn cứ mã nguồn
    t = doc.table("Căn cứ chính của nội dung hiện trạng")
    r = doc.find_row(t, "Căn cứ chính của nội dung hiện trạng")
    doc.set_cell(t, r, 0, f"Mã nguồn ai-platform (AI Platform 6.11.18), ai-work (AI Workspace 6.11.16), nhánh main ngày {DATE}",
                 code=False)
    doc.save(dst)
    return doc


if __name__ == "__main__":
    src, out = sys.argv[1:3]
    vr = ("Đạt", "PostgreSQL thật, collection tạm; gồm phần chỉ mục HNSW với 20.000 vector")
    if len(sys.argv) > 3:
        vr = tuple(open(sys.argv[3], encoding="utf-8").read().strip().split("\n", 1))
    os.makedirs(out, exist_ok=True)
    d = tl2(src, os.path.join(out, T2[:-5] + ".src.docx"), vr)
    print("tl2", len(d.changes), "thay đổi")
    for c in d.changes:
        print("  -", c)
