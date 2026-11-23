# -*- coding: utf-8 -*-
"""TL1 bản 0.21 (23/11/2026): ứng dụng 6.11.23 đã đăng trang phát hành (v6.11.23).
- Trang bìa; mục "Những điểm mới trong bản 0.21"; 2.5, 2.8 (tên tệp, cập nhật), chú thích ảnh. Ảnh trang Đăng nhập giữ bản 6.11.21
  (giao diện trang này không đổi).
Đầu vào: TL1 bản 0.20 ĐÃ điền mục lục. Đầu ra: bản 0.21 chưa điền mục lục ở <thư mục ra>/<tên>.src.docx, sau đó chép vào
_tools/toc_work/ và chạy build_toc.py TL1.
    python update_tl1_v021.py <TL1_v020.docx> <thư mục ra>"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docx.oxml.ns import qn  # noqa: E402
from docedit import Doc  # noqa: E402
from doc_versioning import new_section, set_caption, unfill_toc  # noqa: E402

T1 = "iThunder_AI_OS_TL1_Huong_dan_su_dung_AI_Workspace.docx"
DATE, OLD, VER = "23/11/2026", "6.11.21", "6.11.23"


def bump(doc, el, count):
    """Đổi tên tệp / số phiên bản 6.11.21 trong mọi run của một phần tử; count: số chỗ phải thấy."""
    n = 0
    for t in el.iter(qn("w:t")):
        if OLD in (t.text or ""):
            n += t.text.count(OLD)
            t.text = t.text.replace(OLD, VER)
    assert n == count, (n, count, doc.text_of(el)[:80])
    doc.changes.append(f"Đổi phiên bản ({n} chỗ): {doc.text_of(el)[:60]}")


def tl1(src, dst):
    doc = Doc(src)
    unfill_toc(doc)
    t = doc.table("Phiên bản tài liệu")
    doc.set_cell(t, doc.find_row(t, "Phiên bản tài liệu"), 1,
                 f"0.21, bản dự thảo ngày {DATE} (cập nhật bản 0.20: ứng dụng {VER} trên trang phát hành)")
    doc.set_cell(t, doc.find_row(t, "Phiên bản phần mềm"), 1,
                 f"AI Workspace {VER} (Web, ứng dụng Windows và Android) và AI Platform 6.11.18, nhánh main ngày {DATE}")
    new_section(doc, "Những điểm mới trong bản 0.20", "Những điểm mới trong bản 0.21",
                f"So với bản 0.20, ứng dụng Windows và Android {VER} đã có trên trang phát hành:",
                "So với bản 0.19, cách quét khuôn mặt",
                "Bản Web vẫn kiểm tra người thật; ứng dụng Windows",
                [(f"Tải ứng dụng {VER}; cài đè lên bản trước", "2.8"),
                 (f"Ứng dụng {VER} khi AI Gateway của Tổ chức chưa cập nhật", "2.5")])

    # ---- 2.5
    box = doc.find("Ứng dụng 6.11.21 trở về trước cũng phải làm các bước này", "tbl")
    doc.replace_in(box, "Ứng dụng 6.11.21 trở về trước cũng phải làm các bước này.",
                   f"Ứng dụng 6.11.21 trở về trước, và ứng dụng {VER} khi AI Gateway của Tổ chức chưa cập nhật, cũng phải làm các bước này "
                   "và không hỏi mã xác nhận.")
    cap = doc.find("Ứng dụng Windows 6.11.23 (bản thử): nhập mã xác nhận", "p")
    set_caption(doc, cap, f"Ứng dụng Windows {VER}: nhập mã xác nhận thiết bị mới gửi qua Email (dữ liệu mẫu)")
    cap = doc.find("mục Thiết bị FaceID, trên ứng dụng Windows 6.11.23 (bản thử, dữ liệu mẫu)", "p")
    set_caption(doc, cap, f"Hồ sơ cá nhân, mục Thiết bị FaceID, trên ứng dụng Windows {VER} (dữ liệu mẫu)")

    # ---- 2.8
    bump(doc, doc.find("Ngoài trình duyệt web, bạn có thể dùng AI Workspace như một ứng dụng", "p"), 1)
    bump(doc, doc.find("Thiết bịYêu cầuTệp cài đặt", "tbl"), 2)
    box = doc.find("Bản thử nghiệm đã được kiểm tra cài đặt, gỡ cài đặt", "tbl")
    doc.replace_in(box, "Bản thử 6.11.23 đã được kiểm tra", f"Bản {VER} đã được kiểm tra")
    bump(doc, doc.find("tệp nén cho thiết bị của bạn: ", "p"), 2)
    bump(doc, doc.find("(Android): tệp cài đặt.", "p"), 2)
    bump(doc, doc.find("từ trang phát hành, hoặc nhận bộ cài từ Quản trị viên", "p"), 1)
    bump(doc, doc.find("rồi mở tệp .apk trên điện thoại", "p"), 2)
    p = doc.find("Nếu Google Play Protect cảnh báo ứng dụng chưa được xác minh", "p")
    doc.replace_in(p, "bản 6.11.16: cài đè lên", "bản trước (từ 6.11.16): cài đè lên")
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
