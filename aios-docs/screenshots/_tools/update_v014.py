# -*- coding: utf-8 -*-
"""TL2 bản 0.14 (20/11/2026): bản phát hành công khai. Gỡ Phụ lục G (Vấn đề đã biết, NỘI BỘ) khỏi tài liệu, Phụ lục H
(Tài liệu tham chiếu) thành Phụ lục G; bỏ các tham chiếu G.2, G.3 trong bảng "Những điểm mới". Nội dung kỹ thuật không đổi.
Bản đầy đủ có Phụ lục G chỉ lưu nội bộ, ngoài kho công khai (~/ai-reference/noi-bo/, thư mục bị .gitignore bỏ qua).
Đầu vào: TL2 bản 0.13 ĐÃ điền mục lục. Đầu ra: bản 0.14 chưa điền mục lục ở <thư mục ra>/<tên>.src.docx, sau đó chép vào
_tools/toc_work/ và chạy build_toc.py TL2.
    python update_v014.py <TL2_v013.docx> <thư mục ra>"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docedit import Doc  # noqa: E402
from doc_versioning import new_section, unfill_toc  # noqa: E402

T2 = "iThunder_AI_OS_TL2_Cai_dat_va_Van_hanh_AI_Platform.docx"


def remove_appendix_g(doc):
    """Xóa từ tiêu đề 'Phụ lục G. Vấn đề đã biết' tới trước tiêu đề 'Phụ lục H. Tài liệu tham chiếu'."""
    g = doc.heading("Phụ lục G. Vấn đề đã biết")
    h = doc.heading("Phụ lục H. Tài liệu tham chiếu")
    els = doc.elements()
    part = els[els.index(g):els.index(h)]
    for e in part:
        doc.body.remove(e)
    doc.changes.append(f"Gỡ Phụ lục G ({len(part)} phần tử)")
    doc.replace_in(h, "Phụ lục H.", "Phụ lục G.")


def tl2(src, dst):
    doc = Doc(src)
    unfill_toc(doc)
    t = doc.table("Phiên bản tài liệu")
    doc.set_cell(t, doc.find_row(t, "Phiên bản tài liệu"), 1,
                 "0.14, bản dự thảo ngày 20/11/2026 (bản công khai: nội dung kỹ thuật như bản 0.13, bỏ phụ lục nội bộ)")

    # Bảng "Những điểm mới trong bản 0.13": bỏ tham chiếu tới Phụ lục G cũ
    t = doc.table("Quy tắc trả lời theo nơi chạy mô hình")
    doc.set_cell(t, doc.find_row(t, "Quy tắc trả lời theo nơi chạy mô hình"), 1, "5.7, 8.2, 8.6", code=False)
    doc.set_cell(t, doc.find_row(t, "Gói cloud_ai"), 1, "5.6, 10.1, 10.2, A.4, B, C", code=False)

    new_section(doc, "Những điểm mới trong bản 0.13", "Những điểm mới trong bản 0.14",
                "Bản 0.14 là bản phát hành công khai của tài liệu, nội dung kỹ thuật giữ nguyên như bản 0.13. Các thay đổi:",
                "Đối chiếu hai kho mã nguồn sau khi hợp nhất PR #2 của iThunder-AI-OS",
                "Quy tắc trả lời theo nơi chạy mô hình",
                [("Gỡ phụ lục nội bộ dùng cho đội phát triển; Phụ lục Tài liệu tham chiếu đổi từ H thành G", "G"),
                 ("Việc cần làm để vận hành an toàn vẫn nằm ở các chương 4, 9 và 10", "4, 9, 10")])

    remove_appendix_g(doc)
    for e in doc.elements():
        txt = doc.text_of(e)
        for bad in ("Vấn đề đã biết", "NỘI BỘ", "G.1.", "G.2.", "G.3.", "Phụ lục H"):
            if bad in txt and "TOC " not in txt:
                raise SystemExit(f"Còn sót '{bad}': {txt[:80]}")
    doc.save(dst)
    return doc


if __name__ == "__main__":
    src, out = sys.argv[1:3]
    os.makedirs(out, exist_ok=True)
    d = tl2(src, os.path.join(out, T2[:-5] + ".src.docx"))
    print("tl2", len(d.changes), "thay đổi")
    for c in d.changes:
        print("  -", c)
