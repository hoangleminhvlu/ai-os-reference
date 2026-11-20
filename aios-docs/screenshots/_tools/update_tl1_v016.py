# -*- coding: utf-8 -*-
"""TL1 bản 0.16 (21/11/2026): FaceID có kiểm tra người thật (liveness), iThunder-AI-OS PR #5, ai-workspace PR #6.
- Trang bìa; mục "Những điểm mới trong bản 0.16".
- 2.5: các bước quét mới (nhắc khung hình, thử thách ngẫu nhiên, thời hạn), thay Ảnh màn hình 2, hộp Lưu ý.
- 12.2: thông báo mới; C.2: người dùng không qua bước quay mặt; Phụ lục D: thuật ngữ Kiểm tra người thật.
Đầu vào: TL1 bản 0.15 ĐÃ điền mục lục. Đầu ra: bản 0.16 chưa điền mục lục ở <thư mục ra>/<tên>.src.docx, sau đó chép vào
_tools/toc_work/ và chạy build_toc.py TL1.
    python update_tl1_v016.py <TL1_v015.docx> <thư mục ra>"""
import copy
import os
import sys

from docx.oxml.ns import qn
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docedit import Doc  # noqa: E402
from doc_versioning import blank_para, new_section, set_caption, unfill_toc  # noqa: E402

T1 = "iThunder_AI_OS_TL1_Huong_dan_su_dung_AI_Workspace.docx"
DATE = "21/11/2026"
SHOT = os.path.join(HERE, "..", "tl1_02_faceid_live.png")
STEPS = [
    "Khi trình duyệt hoặc ứng dụng hỏi, chọn **Cho phép** dùng camera.",
    "Giữ khuôn mặt trong khung tròn, nhìn thẳng, ở nơi đủ sáng. Hệ thống tự chụp liên tục; dòng chữ màu vàng dưới khung nhắc "
    "khi khuôn mặt quá xa, lệch khỏi khung, quá tối, bị mờ hoặc có nhiều người trong khung.",
    "Làm theo hướng dẫn **kiểm tra người thật** hiện phía trên khung tròn: hai thử thách chọn ngẫu nhiên trong ba thử thách "
    "**từ từ quay mặt sang trái**, **sang phải** của bạn, **nhắm mắt khoảng một giây rồi mở**. Ảnh trong khung là ảnh gương, "
    "nên quay sang trái của bạn cũng là sang trái trên màn hình. Thanh tiến độ cho biết đã xong mấy bước.",
    "Cuối cùng **nhìn thẳng lại vào camera**. Khi khớp, bạn tự động vào AI Workspace.",
    "Nếu chưa khớp sau nhiều lần, hoặc quá 90 giây, hệ thống tạm dừng. Bấm **Quét lại khuôn mặt**, hoặc **Hủy quét, quay lại** "
    "để nhập lại PassCode.",
]


def replace_picture_fit(doc, caption_needle, path):
    """Thay ảnh trước chú thích, giữ bề rộng, đổi chiều cao theo tỷ lệ ảnh mới (ảnh không bị méo)."""
    cap = doc.find(caption_needle, "p")
    pic = cap.getprevious()
    blip = pic.find(".//" + qn("a:blip"))
    part = doc.d.part.related_parts[blip.get(qn("r:embed"))]
    with open(path, "rb") as f:
        part._blob = f.read()
    w, h = Image.open(path).size
    for tag in ("wp:extent", "a:ext"):
        for ext in pic.iter(qn(tag)):
            cx = int(ext.get("cx"))
            ext.set("cy", str(int(cx * h / w)))
    doc.changes.append(f"Thay ảnh: {caption_needle}")
    return cap


def tl1(src, dst):
    doc = Doc(src)
    unfill_toc(doc)
    blank = blank_para(doc)

    t = doc.table("Phiên bản tài liệu")
    doc.set_cell(t, doc.find_row(t, "Phiên bản tài liệu"), 1,
                 f"0.16, bản dự thảo ngày {DATE} (cập nhật bản 0.15: FaceID có kiểm tra người thật)")
    doc.set_cell(t, doc.find_row(t, "Phiên bản phần mềm"), 1,
                 f"AI Workspace 6.11.16 (Web, ứng dụng Windows và Android) và AI Platform 6.11.18, nhánh main ngày {DATE}")
    new_section(doc, "Những điểm mới trong bản 0.15", "Những điểm mới trong bản 0.16",
                "So với bản 0.15, khi quét khuôn mặt, AI Workspace yêu cầu kiểm tra người thật để ảnh chụp không dùng để "
                "đăng nhập được:",
                "So với bản 0.14, trang Hồ sơ cá nhân",
                "Xem các ảnh khuôn mặt đã lưu kèm đánh giá chất lượng",
                [("Quét FaceID: nhắc khi khung hình chưa đạt, hai thử thách ngẫu nhiên (quay trái, quay phải, nhắm mắt)", "2.5"),
                 ("Thông báo mới khi quét FaceID; ứng dụng bản cũ cần cập nhật", "12.2, 2.5"),
                 ("Người dùng không qua được bước quay mặt", "C.2")])

    # ---- 2.5: viết lại các bước
    h = doc.heading("2.5. Quét khuôn mặt (FaceID)")
    old = []
    el = h.getnext()
    while doc.text_of(el).startswith("Bước "):
        old.append(el)
        el = el.getnext()
    tpl = copy.deepcopy(old[0])
    for e in old:
        e.getparent().remove(e)
    last = h
    for n, text in enumerate(STEPS, start=1):
        last = doc.step_after(last, tpl, n, text)
    cap = replace_picture_fit(doc, "Ảnh màn hình 2. ", SHOT)
    set_caption(doc, cap, "Kích hoạt Trắc sinh học FaceID: bước 1/2 của kiểm tra người thật (ảnh camera được thay bằng "
                          "hình minh họa)")
    box = doc.find("Cảnh báo. Khuôn mặt quét đầu tiên", "tbl")
    gap = box.getnext()
    doc.note_after(gap, "Lưu ý", "Ảnh in, ảnh trên điện thoại hay màn hình không qua được bước kiểm tra người thật. Lần đầu "
                                 "đăng ký khuôn mặt cũng phải làm đủ các bước. Nếu ứng dụng báo AI Gateway yêu cầu kiểm tra "
                                 "người thật, hãy tải lại trang (bản Web) hoặc cài bản ứng dụng mới (Windows, Android).",
                   blank_template=copy.deepcopy(blank))

    # ---- 12.2
    t = doc.table("Mã PassCode không chính xác hoặc tài khoản chưa được phê duyệt")
    r = doc.find_row(t, "Không phát hiện khuôn mặt. Hãy thử lại!")
    doc.set_cell(t, r, 0, "Không thấy khuôn mặt... / Khuôn mặt quá nhỏ... / Ảnh quá tối... (chữ màu vàng dưới khung)")
    doc.set_cell(t, r, 1, "Camera chưa thấy rõ mặt. Đây chỉ là lời nhắc, không tính là lần thất bại.")
    doc.set_cell(t, r, 2, "Làm theo lời nhắc: đến nơi đủ sáng, nhìn thẳng, đưa mặt lại gần và vào giữa khung tròn, chỉ một "
                          "người trong khung.")
    for values in [
        ("Hết thời gian thực hiện. Hãy làm lại...", "Chưa làm xong các bước kiểm tra người thật trong khoảng 40 giây.",
         "Làm lại từ đầu theo hướng dẫn mới (thử thách có thể khác lần trước)."),
        ("Khuôn mặt thay đổi đột ngột. Hãy làm lại từ đầu...", "Khuôn mặt nhảy vị trí hoặc đổi cỡ quá nhanh, như khi đổi ảnh.",
         "Quay mặt từ từ, giữ mặt trong khung tròn, không di chuyển camera."),
        ("AI Gateway yêu cầu kiểm tra người thật khi quét FaceID...", "Trình duyệt hoặc ứng dụng đang dùng bản cũ.",
         "Tải lại trang (bản Web) hoặc cài bản ứng dụng mới."),
        ("Đã hết thời gian quét FaceID. Hãy bấm Quét lại...", "Quét quá 90 giây mà chưa xác thực xong.",
         "Bấm Quét lại khuôn mặt và làm theo hướng dẫn."),
    ]:
        r = doc.add_row(t, list(values), like=r, after=r)

    # ---- C.2
    t = doc.table("Người dùng nghi lộ PassCode")
    r = doc.find_row(t, "Khuôn mặt không nhận được")
    doc.add_row(t, ["Người dùng không qua được bước quay mặt", "Nhắc quay từ từ, vừa phải, giữ mặt trong khung, ở nơi đủ sáng. "
                    "Nếu nhiều người cùng gặp, báo đội vận hành điều chỉnh ngưỡng (Tài liệu Cài đặt và Vận hành, mục 8.3)."],
                like=r, after=r)

    # ---- Phụ lục D
    for tbl in doc.d.tables:
        if tbl.rows[0].cells[0].text == "Thuật ngữ":
            t = tbl
    r = doc.find_row(t, "FaceID")
    doc.add_row(t, ["Kiểm tra người thật (liveness)", "Bước của FaceID yêu cầu quay mặt, nhắm mắt theo hướng dẫn ngẫu nhiên, để "
                    "ảnh chụp không dùng để đăng nhập được."], like=r, after=r)

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
