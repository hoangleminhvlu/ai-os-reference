# -*- coding: utf-8 -*-
"""TL2 bản 0.19 (21/11/2026): thống kê AI Tokens (iThunder-AI-OS PR #8, ai-workspace PR #11).
- Trang bìa; mục "Những điểm mới trong bản 0.19"; 2.2 phân hệ.
- 8.9 (mới): sổ iAIP_ai_usage, API, phạm vi admin / manager, nạp số liệu cũ, giới hạn, ảnh màn hình.
- 6.6 nhật ký; 10.1; 11.1 lệnh, Bảng 3; Phụ lục A.4, B, C, G. Đánh lại số ảnh màn hình.
Đầu vào: TL2 bản 0.18 ĐÃ điền mục lục. Đầu ra: bản 0.19 chưa điền mục lục ở <thư mục ra>/<tên>.src.docx, sau đó chép vào
_tools/toc_work/ và chạy build_toc.py TL2.
    python update_v019.py <TL2_v018.docx> <thư mục ra>"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docedit import Doc  # noqa: E402
from doc_versioning import clone_heading, insert_picture, new_section, renumber_shots, unfill_toc  # noqa: E402

T2 = "iThunder_AI_OS_TL2_Cai_dat_va_Van_hanh_AI_Platform.docx"
DATE = "21/11/2026"
SHOT = os.path.join(HERE, "..", "tl2_07_ai_tokens.png")
POINTS = [
    "**Ghi sổ**: AnswerClock.finish ghi mỗi lượt mô hình trả lời (AI Chat/Work, Trợ lý AI của Trao đổi trực tuyến) vào bảng "
    "iAIP_ai_usage ở luồng phụ: người dùng, kênh, Kho, mô hình, Máy chủ AI, nơi chạy (onpremise | onclouds), nhà cung cấp "
    "On-Clouds, trạng thái, input_tokens, output_tokens, total_tokens, thời lượng. Lỗi ghi sổ chỉ ghi nhật ký, không ảnh hưởng "
    "câu trả lời. Sổ tách khỏi lịch sử hội thoại: người dùng xóa phiên, vai trò không lưu lịch sử vẫn được thống kê.",
    "**Số token** do mô hình báo (mục 8.8: Ollama prompt_eval_count / eval_count; On-Clouds: usage của nhà cung cấp). Mô hình "
    "không báo số token thì ghi measured = false và được đếm riêng. Lượt refused (không gọi mô hình) không được ghi; token tạo "
    "embedding khi nạp Kho Tri thức chưa được tính.",
    "**Thống kê** (GET /api/system/ai-usage): tổng theo On-Premise / On-Clouds và kênh, chuỗi theo ngày hoặc tháng (ngày, tháng "
    "theo AI_USAGE_TIMEZONE, mặc định Asia/Ho_Chi_Minh), theo mô hình, theo người dùng; lọc nơi chạy, kênh, người dùng. CSV chi "
    "tiết theo kỳ, người dùng, mô hình (ghi nhật ký ai_usage.export).",
    "**Phạm vi**: admin xem toàn Tổ chức. manager (Quản trị Đơn vị) xem số tổng của toàn Tổ chức, nhưng bảng theo người dùng và "
    "lọc theo người dùng chỉ gồm người dùng thuộc đơn vị mình quản lý; vai trò đọc lại từ OpenLDAP ở mỗi yêu cầu. Trang Giám sát "
    "Hệ thống mở cho admin và manager; tab AI On-Clouds vẫn chỉ admin.",
    "**Số liệu cũ**: POST /api/system/ai-usage/backfill (admin, nút Nạp số liệu cũ) nạp số token từ metadata câu trả lời đã lưu "
    "(iAIP_chat_messages, iAIP_comm_messages); chạy lại không trùng nhờ khóa source + source_id.",
]


def tl2(src, dst):
    doc = Doc(src)
    unfill_toc(doc)
    t = doc.table("Phiên bản tài liệu")
    doc.set_cell(t, doc.find_row(t, "Phiên bản tài liệu"), 1,
                 f"0.19, bản dự thảo ngày {DATE} (cập nhật bản 0.18: thống kê AI Tokens)")
    doc.set_cell(t, doc.find_row(t, "Phiên bản phần mềm"), 1,
                 f"AI Workspace 6.11.21 và AI Platform 6.11.18 (nhánh main ngày {DATE}, sau khi hợp nhất PR #5 đến #8 của "
                 "iThunder-AI-OS và PR #6 đến #11 của ai-workspace)")
    new_section(doc, "Những điểm mới trong bản 0.18", "Những điểm mới trong bản 0.19",
                "Bổ sung thống kê AI Tokens theo ngày, tháng làm căn cứ đánh giá hiệu quả sử dụng AI On-Premise và On-Clouds "
                "(iThunder-AI-OS PR #8, ai-workspace PR #11). Các thay đổi và nơi mô tả:",
                "Bổ sung kiểm tra người thật (liveness)",
                "Luồng đăng nhập: khung hình vuông 3 khung mỗi giây",
                [("Sổ iAIP_ai_usage, thống kê, phạm vi admin / manager, nạp số liệu cũ", "8.9, 2.2"),
                 ("API /api/system/ai-usage, bảng CSDL, biến cấu hình", "A.4, B, C"),
                 ("Nhật ký, cơ chế an toàn, kiểm thử", "6.6, 10.1, 11.1")])

    t = doc.table("Phân hệTệp chính")
    doc.add_row(t, ["Thống kê AI Tokens", "ai_usage/", "Sổ ghi mỗi lượt mô hình trả lời, thống kê theo ngày, tháng, người dùng, mô "
                    "hình; On-Premise / On-Clouds", "[[ĐÃ CÓ]]"])

    # ---- 8.9 mới, trước chương 9
    h9 = doc.heading("9. Vận hành")
    h = clone_heading(doc, "8.8. Bằng chứng của từng câu trả lời", h9.getprevious(), "8.9. Thống kê AI Tokens")
    intro_tpl = doc.find("Theo nguyên tắc truy vết của TARA, mỗi câu trả lời", "p")
    intro = doc.para_after(h, intro_tpl, "[[ĐÃ CÓ]] Số AI Tokens mô hình đã dùng theo ngày, tháng, tách On-Premise và On-Clouds, "
                                         "theo người dùng và mô hình là căn cứ đánh giá hiệu quả sử dụng nền tảng. Giao diện: "
                                         "**Giám sát Hệ thống › AI Tokens** (ai_usage/, AiUsagePanel.jsx):")
    bullet_tpl = doc.find("Khóa là DN. gidNumber", "p")
    last = intro
    for text in POINTS:
        last = doc.para_after(last, bullet_tpl, text)
    insert_picture(doc, last, SHOT, "Giám sát Hệ thống, tab AI Tokens: tổng theo On-Premise và On-Clouds, biểu đồ theo ngày có chú "
                                    "thích khi rê chuột, bảng theo người dùng và mô hình (dữ liệu mẫu)", 16.0, "Ảnh màn hình 2. ")

    # ---- 6.6, 10.1
    p = doc.find("Xem ảnh FaceID của người khác ghi user.faceid.photos.view", "p")
    doc.para_after(p, p, "Xuất CSV thống kê AI Tokens ghi ai_usage.export; nạp số liệu cũ ghi ai_usage.backfill (mục 8.9).")
    p = doc.find("FaceID có kiểm tra người thật do AI Platform phát thử thách", "p")
    doc.para_after(p, p, "Thống kê AI Tokens chỉ cho admin, manager; manager chỉ thấy chi tiết người dùng thuộc đơn vị mình; xuất "
                         "CSV và nạp số liệu cũ được ghi nhật ký (mục 8.9).")

    # ---- 11.1
    t = doc.table("AIP=.venv")
    lines = [x.text for x in t.cell(0, 0).paragraphs]
    i = lines.index("$AIP tests/knowledge/test_keyword_db.py") + 1
    lines[i:i] = ["$AIP tests/ai_usage/test_ai_usage_db.py   # sổ AI Tokens, mã Tổ chức tạm"]
    doc.set_code(t, lines)
    t = doc.table("Bộ kiểm thửKết quả ngày")
    r = doc.find_row(t, "tests/identity/run_all.py")
    doc.set_cell(t, r, 1, "Đạt (19 tệp, gồm test_14_rag_chat.py so sánh trực tiếp lời nhắc của AI Chat/Work và Trao đổi trực "
                          "tuyến; test_16_faceid_liveness.py mô phỏng hình học đầu người thật và ảnh phẳng; test_17 đến test_19 "
                          "ngày 21/11/2026)", code=False)
    r = doc.find_row(t, "tests/knowledge/test_archive_db.py")
    doc.add_row(t, ["tests/ai_usage/test_ai_usage_db.py", "Đạt", "PostgreSQL thật, mã Tổ chức tạm, dọn sạch: tổng, ranh giới ngày "
                    "theo giờ Việt Nam, theo tháng, lọc, CSV, nạp số liệu cũ"], like=r, after=r)

    # ---- Phụ lục A.4, B, C, G
    t = doc.table("POST /api/mcpMCP Server")
    doc.add_row(t, ["GET /api/system/ai-usage, /api/system/ai-usage.csv; POST /api/system/ai-usage/backfill",
                    "Thống kê AI Tokens (admin, manager; manager thấy chi tiết người dùng thuộc đơn vị mình); nạp số liệu cũ chỉ admin"])
    t = doc.table("BảngDùng cho")
    r = doc.find_row(t, "iAIP_storage_links")
    doc.add_row(t, ["iAIP_ai_usage", "Sổ AI Tokens: mỗi lượt mô hình trả lời (người dùng, mô hình, nơi chạy, số token)",
                    "init_database.sh, mã nguồn"], like=r, after=r)
    t = doc.table("NhómBiến trong .env")
    doc.add_row(t, ["Thống kê AI Tokens", "AI_USAGE_ENABLED, AI_USAGE_TIMEZONE (mục 8.9)"])
    t = doc.table("Giới hạn trong config.py")
    doc.add_row(t, ["Thống kê AI Tokens", "Mặc định 30 ngày (theo ngày), 12 tháng (theo tháng); tối đa 366 ngày, 36 tháng; bảng "
                    "người dùng 500 dòng, mô hình 50 dòng"])
    t = doc.table("Căn cứ chính của nội dung hiện trạng")
    r = doc.find_row(t, "Căn cứ chính của nội dung hiện trạng")
    doc.set_cell(t, r, 0, f"Mã nguồn ai-platform (AI Platform 6.11.18), ai-work (AI Workspace 6.11.21), nhánh main ngày {DATE} "
                          "sau khi hợp nhất FaceID liveness (iThunder-AI-OS PR #5, ai-workspace PR #6), Kết nối Kho Đám mây "
                          "(PR #6, PR #7), Ảnh FaceID trong hồ sơ (PR #7, PR #8), ẩn địa chỉ AI Gateway và phóng to (ai-workspace "
                          "PR #9), AI Workspace 6.11.21 (PR #10), thống kê AI Tokens (PR #8, PR #11)", code=False)

    renumber_shots(doc)
    doc.save(dst)
    return doc


if __name__ == "__main__":
    src, out = sys.argv[1:3]
    os.makedirs(out, exist_ok=True)
    d = tl2(src, os.path.join(out, T2[:-5] + ".src.docx"))
    print("tl2", len(d.changes), "thay đổi")
    for c in d.changes:
        if c != "-":
            print("  -", c)
