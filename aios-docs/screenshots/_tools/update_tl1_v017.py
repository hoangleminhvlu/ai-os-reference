# -*- coding: utf-8 -*-
"""TL1 bản 0.17 (21/11/2026): Giám sát Hệ thống › AI Tokens (iThunder-AI-OS PR #8, ai-workspace PR #11).
- Trang bìa; mục "Những điểm mới trong bản 0.17".
- C.5: tab AI Tokens, quyền của Quản trị Đơn vị, ảnh màn hình mới (đánh lại số các ảnh sau).
Đầu vào: TL1 bản 0.16 ĐÃ điền mục lục. Đầu ra: bản 0.17 chưa điền mục lục ở <thư mục ra>/<tên>.src.docx, sau đó chép vào
_tools/toc_work/ và chạy build_toc.py TL1.
    python update_tl1_v017.py <TL1_v016.docx> <thư mục ra>"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docedit import Doc  # noqa: E402
from doc_versioning import insert_picture, new_section, renumber_shots, unfill_toc  # noqa: E402

T1 = "iThunder_AI_OS_TL1_Huong_dan_su_dung_AI_Workspace.docx"
DATE = "21/11/2026"
SHOT = os.path.join(HERE, "..", "tl2_07_ai_tokens.png")


def tl1(src, dst):
    doc = Doc(src)
    unfill_toc(doc)
    t = doc.table("Phiên bản tài liệu")
    doc.set_cell(t, doc.find_row(t, "Phiên bản tài liệu"), 1,
                 f"0.17, bản dự thảo ngày {DATE} (cập nhật bản 0.16: thống kê AI Tokens trong Giám sát Hệ thống)")
    doc.set_cell(t, doc.find_row(t, "Phiên bản phần mềm"), 1,
                 f"AI Workspace 6.11.21 (Web, ứng dụng Windows và Android) và AI Platform 6.11.18, nhánh main ngày {DATE}")
    new_section(doc, "Những điểm mới trong bản 0.16", "Những điểm mới trong bản 0.17",
                "So với bản 0.16, trang Giám sát Hệ thống có thêm tab AI Tokens và mở cho cả Quản trị Đơn vị:",
                "So với bản 0.15, khi quét khuôn mặt",
                "Quét FaceID: nhắc khi khung hình chưa đạt",
                [("Thống kê AI Tokens theo ngày, tháng, tách On-Premise và On-Clouds, theo người dùng, mô hình", "C.5"),
                 ("Quản trị Đơn vị (manager) mở được Giám sát Hệ thống; tab AI On-Clouds vẫn chỉ Quản trị viên", "C.5")])

    first = doc.find("Giám sát Hệ thống, tab Máy chủ AI:", "p")
    p = doc.para_after(first, first,
                       "Tab **AI Tokens** [[ĐÃ CÓ]]: số AI Tokens mô hình AI đã dùng theo ngày hoặc tháng, tách **On-Premise** (Máy "
                       "chủ AI của Tổ chức) và **On-Clouds** (dịch vụ AI qua Internet), kèm số lượt trả lời, người dùng, bảng theo "
                       "người dùng, theo mô hình. Lọc theo nơi chạy, kênh (AI Chat/Work, Trao đổi trực tuyến), người dùng; bấm tên "
                       "một người trong bảng để xem riêng người đó; **Xuất CSV** để phân tích bằng Excel. Đây là căn cứ đánh giá "
                       "hiệu quả sử dụng AI On-Premise so với On-Clouds.")
    p = doc.para_after(p, first,
                       "Quản trị viên xem toàn Tổ chức và bấm **Nạp số liệu cũ** một lần để đưa số token của các câu trả lời đã lưu "
                       "trước đây vào thống kê. Quản trị Đơn vị (manager) mở được trang Giám sát Hệ thống, xem số tổng của Tổ chức, "
                       "còn bảng theo người dùng chỉ gồm người thuộc đơn vị mình quản lý; tab AI On-Clouds vẫn chỉ Quản trị viên.")
    insert_picture(doc, p, SHOT, "Giám sát Hệ thống, tab AI Tokens: tổng theo On-Premise và On-Clouds, biểu đồ theo ngày, bảng "
                                 "theo người dùng và mô hình (dữ liệu mẫu)", 16.0, "Ảnh màn hình 12. ")
    renumber_shots(doc)
    doc.save(dst)
    return doc


if __name__ == "__main__":
    src, out = sys.argv[1:3]
    os.makedirs(out, exist_ok=True)
    d = tl1(src, os.path.join(out, T1[:-5] + ".src.docx"))
    print("tl1", len(d.changes), "thay đổi")
    for c in d.changes:
        if c != "-":
            print("  -", c)
