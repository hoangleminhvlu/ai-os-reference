# -*- coding: utf-8 -*-
"""TL1 bản 0.22 (24/11/2026): thông báo khuyến nghị cài ứng dụng trước trang Đăng nhập của bản Web (ai-workspace PR #19).
- Trang bìa; mục "Những điểm mới trong bản 0.22"; 2.1, 2.2 (thay Ảnh màn hình 1), 2.4 (ảnh mới tl1_15_web_notice.png), 2.8, 12.1.
Đầu vào: TL1 bản 0.21 ĐÃ điền mục lục. Đầu ra: bản 0.22 chưa điền mục lục ở <thư mục ra>/<tên>.src.docx, sau đó chép vào
_tools/toc_work/ và chạy build_toc.py TL1.
    python update_tl1_v022.py <TL1_v021.docx> <thư mục ra>"""
import copy
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docedit import Doc  # noqa: E402
from doc_versioning import blank_para, insert_picture, new_section, renumber_shots, replace_picture, set_caption, unfill_toc  # noqa: E402

T1 = "iThunder_AI_OS_TL1_Huong_dan_su_dung_AI_Workspace.docx"
DATE, VER = "24/11/2026", "6.11.23"
SHOT_LOGIN = os.path.join(HERE, "..", "tl1_01_login.jpg")
SHOT_NOTICE = os.path.join(HERE, "..", "tl1_15_web_notice.png")


def tl1(src, dst):
    doc = Doc(src)
    unfill_toc(doc)
    t = doc.table("Phiên bản tài liệu")
    doc.set_cell(t, doc.find_row(t, "Phiên bản tài liệu"), 1,
                 f"0.22, bản dự thảo ngày {DATE} (cập nhật bản 0.21: thông báo khuyến nghị cài ứng dụng trên bản Web)")
    doc.set_cell(t, doc.find_row(t, "Phiên bản phần mềm"), 1,
                 f"AI Workspace Web nhánh main ngày {DATE}, ứng dụng Windows và Android {VER}, AI Platform 6.11.18")
    new_section(doc, "Những điểm mới trong bản 0.21", "Những điểm mới trong bản 0.22",
                "So với bản 0.21, bản Web khuyến nghị dùng ứng dụng trước khi đăng nhập:",
                "So với bản 0.20, ứng dụng Windows và Android",
                f"Tải ứng dụng {VER}; cài đè lên bản trước",
                [("Thông báo Nên dùng ứng dụng AI Workspace trước trang Đăng nhập của bản Web; nút Tải ứng dụng Windows / Android", "2.4"),
                 ("Câu hỏi thường gặp về thông báo và trang tải ứng dụng", "12.1")])

    # ---- 2.1
    p = doc.find("Địa chỉ AI Workspace do tổ chức cung cấp, ví dụ", "p")
    doc.replace_in(p, "chọn trong danh sách trên màn hình đăng nhập.",
                   "chọn trong danh sách trên màn hình đăng nhập. Mở bằng trình duyệt, bạn thấy thông báo khuyến nghị cài ứng dụng trước "
                   "trang Đăng nhập (mục 2.4).")

    # ---- 2.2
    cap = replace_picture(doc, "Trang Đăng nhập: khung AI Workspace của Tổ chức, ô PassCode, GetCode và Register", SHOT_LOGIN)
    set_caption(doc, cap, "Trang Đăng nhập của bản Web: khung AI Workspace của Tổ chức, ô PassCode, GetCode, Register và nút tải ứng dụng")

    # ---- 2.4
    h = doc.heading("2.4. Đăng nhập hằng ngày")
    para = "Trước khi nhập PassCode, hãy xem khung AI Workspace của Tổ chức"
    bullet = doc.find("AI Gateway chạy trên Internet: chọn Tổ chức của bạn trong danh sách", "p")
    q = doc.para_after(h, para, "**Thông báo trên bản Web.** Khi mở AI Workspace bằng trình duyệt, trước trang Đăng nhập có thông báo **Nên dùng "
                                "ứng dụng AI Workspace**. Bản Web bắt buộc kiểm tra người thật khi quét FaceID (mục 2.5). Ứng dụng Windows, Android "
                                "tiện hơn vì chỉ cần nhìn thẳng vào camera, và an toàn hơn vì mỗi thiết bị có khóa riêng, thiết bị mới phải xác nhận "
                                "qua Email. Bạn có thể:")
    q = doc.para_after(q, bullet, "Bấm **Windows Desktop** hoặc **Android Mobile** để mở trang phát hành trên GitHub, tải tệp nén cho thiết bị của bạn "
                                  "rồi cài đặt theo mục 2.8. Nút phù hợp với thiết bị đang dùng được làm nổi bật. iPhone, iPad, máy Mac, Linux chưa có "
                                  "ứng dụng.")
    q = doc.para_after(q, bullet, "Bấm **Tiếp tục bằng Web App** để vào trang Đăng nhập như bình thường.")
    blank = blank_para(doc)
    q = doc.para_after(q, para, "Thông báo hiện một lần trong mỗi thẻ (tab) trình duyệt. Muốn xem lại, bấm **Tải ứng dụng Windows / Android** ở cuối "
                                "trang Đăng nhập. Ứng dụng Windows, Android không hiện thông báo này.")
    cap = insert_picture(doc, q, SHOT_NOTICE, "Bản Web: thông báo Nên dùng ứng dụng AI Workspace trước trang Đăng nhập (mở trên máy Windows, dữ liệu mẫu)",
                         9.5, "Trang Đăng nhập của bản Web: khung AI Workspace của Tổ chức")
    b = copy.deepcopy(blank)
    cap.addnext(b)
    n = doc.note_after(b, "Lưu ý", f"Ứng dụng {VER} chỉ bỏ bước kiểm tra người thật khi AI Gateway của Tổ chức đã cập nhật (mục 2.5). Trang phát hành nằm "
                                   "trên Internet: nếu mạng của bạn không vào được, hãy nhận bộ cài từ Quản trị viên.", blank_template=blank)
    doc.para_after(n, para, "**Các bước đăng nhập:**")

    # ---- 2.8
    p = doc.find("Ngoài trình duyệt web, bạn có thể dùng AI Workspace như một ứng dụng", "p")
    doc.replace_in(p, "Ứng dụng chưa có trên Microsoft Store, Google Play.",
                   "Thông báo trước trang Đăng nhập của bản Web cũng dẫn tới trang này (mục 2.4). Ứng dụng chưa có trên Microsoft Store, Google Play.")

    # ---- 12.1
    t = doc.table("Vì sao hôm nay phải quét mặt, hôm qua thì không?")
    r = doc.find_row(t, "Vì sao hôm nay phải quét mặt")
    for row in [["**Vì sao bản Web hiện thông báo Nên dùng ứng dụng AI Workspace?**",
                 "Bản Web luôn phải kiểm tra người thật khi quét FaceID; ứng dụng Windows, Android an toàn và tiện lợi hơn. Đây là khuyến nghị: bấm "
                 "**Tiếp tục bằng Web App** để đăng nhập như bình thường (mục 2.4)."],
                ["**Bấm nút tải ứng dụng nhưng không mở được trang GitHub?**",
                 "Mạng của bạn có thể không vào được Internet hoặc chặn GitHub. Dùng mạng khác để tải, hoặc nhận bộ cài từ Quản trị viên."]]:
        r = doc.add_row(t, row, like=r, after=r)

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
