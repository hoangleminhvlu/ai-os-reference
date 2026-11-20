# -*- coding: utf-8 -*-
"""TL1 bản 0.14 (20/11/2026): đối chiếu lại với AI Workspace 6.11.16 và AI Platform 6.11.18 trên nhánh main ngày 20/11/2026
(giao diện không đổi so với bản 0.13); Phụ lục C chỉ tới TL2 công khai trên GitHub; C.5 hướng dẫn chuyển khóa API cũ trong
tập tin cấu hình sang tab AI On-Clouds và thu hồi khóa cũ.
Đầu vào: TL1 bản 0.13 ĐÃ điền mục lục. Đầu ra: bản 0.14 chưa điền mục lục ở <thư mục ra>/<tên>.src.docx, sau đó chép vào
_tools/toc_work/ và chạy build_toc.py TL1.
    python update_tl1_v014.py <TL1_v013.docx> <thư mục ra>"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docedit import Doc  # noqa: E402
from doc_versioning import new_section, unfill_toc  # noqa: E402

T1 = "iThunder_AI_OS_TL1_Huong_dan_su_dung_AI_Workspace.docx"
DATE = "20/11/2026"
TL2_URL = "https://github.com/hoangleminhvlu/ai-os-reference"


def tl1(src, dst):
    doc = Doc(src)
    unfill_toc(doc)
    t = doc.table("Phiên bản tài liệu")
    doc.set_cell(t, doc.find_row(t, "Phiên bản tài liệu"), 1,
                 f"0.14, bản dự thảo ngày {DATE} (cập nhật bản 0.13: đối chiếu lại mã nguồn, chỉ dẫn cho Quản trị viên)")
    doc.set_cell(t, doc.find_row(t, "Phiên bản phần mềm"), 1,
                 f"AI Workspace 6.11.16 (Web, ứng dụng Windows và Android) và AI Platform 6.11.18, nhánh main ngày {DATE}")

    new_section(doc, "Những điểm mới trong bản 0.13", "Những điểm mới trong bản 0.14",
                "So với bản 0.13, giao diện AI Workspace không thay đổi. Tài liệu được đối chiếu lại với mã nguồn và bổ sung "
                "chỉ dẫn cho Quản trị viên:",
                "So với bản 0.12, AI Workspace thay đổi theo góp ý của người dùng",
                "Mô hình On-Clouds dùng khóa API riêng của Tổ chức do Quản trị viên nhập",
                [(f"Đối chiếu lại với AI Workspace 6.11.16 và AI Platform 6.11.18, nhánh main ngày {DATE}", "Trang bìa"),
                 ("Nơi tìm Tài liệu Cài đặt và Vận hành AI Platform (TL2)", "C"),
                 ("Chuyển khóa API từng khai báo trong tập tin cấu hình sang tab AI On-Clouds, thu hồi khóa cũ", "C.5")])

    p = doc.find("Phụ lục này tóm tắt các thao tác trên giao diện. Chi tiết kỹ thuật nằm trong", "p")
    doc.set_para(p, "Phụ lục này tóm tắt các thao tác trên giao diện. Chi tiết kỹ thuật nằm trong **Tài liệu Cài đặt và Vận "
                    f"hành AI Platform** (TL2), bản công khai tại {TL2_URL}.")

    anchor = doc.find("Khóa không khai báo trong tập tin cấu hình.", "p")
    doc.para_after(anchor, anchor,
                   "Nếu trước đây khóa API được khai báo trong tập tin cấu hình của AI Platform: tạo khóa mới ở trang của "
                   "nhà cung cấp, nhập vào tab **AI On-Clouds**, rồi **thu hồi khóa cũ** ở trang của nhà cung cấp. Khóa cũ "
                   "từng nằm ở dạng đọc được trên máy chủ nên không nên dùng tiếp.")
    doc.save(dst)
    return doc


if __name__ == "__main__":
    src, out = sys.argv[1:3]
    os.makedirs(out, exist_ok=True)
    d = tl1(src, os.path.join(out, T1[:-5] + ".src.docx"))
    print("tl1", len(d.changes), "thay đổi")
    for c in d.changes:
        print("  -", c)
