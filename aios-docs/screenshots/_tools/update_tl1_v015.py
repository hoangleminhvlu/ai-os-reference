# -*- coding: utf-8 -*-
"""TL1 bản 0.15 (21/11/2026): Ảnh FaceID trong hồ sơ (iThunder-AI-OS PR #7, ai-workspace PR #8).
- Trang bìa; mục "Những điểm mới trong bản 0.15".
- 2.5: hộp Cảnh báo chỉ tới mục 9.1. 9.1: hàng Ảnh FaceID, các bước xem / xóa ảnh, hộp Lưu ý, ảnh màn hình mới.
- 12.2: hai thông báo lỗi mới. C.2: tình huống ảnh khuôn mặt chất lượng thấp, cách xử lý khi không nhận được khuôn mặt.
- Ảnh màn hình mới chèn vào chương 9: đánh lại số các ảnh sau và chỗ dẫn chiếu.
Đầu vào: TL1 bản 0.14 ĐÃ điền mục lục. Đầu ra: bản 0.15 chưa điền mục lục ở <thư mục ra>/<tên>.src.docx, sau đó chép vào
_tools/toc_work/ và chạy build_toc.py TL1.
    python update_tl1_v015.py <TL1_v014.docx> <thư mục ra>"""
import copy
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docedit import Doc  # noqa: E402
from doc_versioning import blank_para, insert_picture, new_section, renumber_shots, unfill_toc  # noqa: E402

T1 = "iThunder_AI_OS_TL1_Huong_dan_su_dung_AI_Workspace.docx"
DATE = "21/11/2026"
SHOT = os.path.join(HERE, "..", "tl1_12_faceid_photos.png")
SHOT_TITLE = ("Hồ sơ cá nhân, mục Ảnh FaceID trong hồ sơ: đã chọn các ảnh chưa đạt để xóa (ảnh khuôn mặt được thay bằng "
              "hình minh họa)")
STEPS = [
    "Mở **Hồ sơ cá nhân**, kéo xuống mục **Ảnh FaceID trong hồ sơ**. Mỗi ảnh có nhãn **Tốt**, **Chất lượng thấp** hoặc "
    "**Không dùng được**, kèm lý do: ảnh mờ, quá tối, quá chói, khuôn mặt nhỏ, không thấy khuôn mặt, hoặc khác hẳn các ảnh "
    "còn lại (có thể là người khác).",
    "Bấm **Chọn ảnh chưa đạt** để đánh dấu mọi ảnh không ở mức Tốt, hoặc tự đánh dấu từng ảnh.",
    "Bấm **Xóa ảnh đã chọn**, xác nhận và nhập lại PassCode.",
    "Hệ thống tính lại mẫu khuôn mặt từ các ảnh còn lại. Ảnh đã xóa không còn được dùng ở lần quét FaceID sau.",
]


def tl1(src, dst):
    doc = Doc(src)
    unfill_toc(doc)
    blank = blank_para(doc)

    t = doc.table("Phiên bản tài liệu")
    doc.set_cell(t, doc.find_row(t, "Phiên bản tài liệu"), 1,
                 f"0.15, bản dự thảo ngày {DATE} (cập nhật bản 0.14: xem và xóa Ảnh FaceID trong hồ sơ)")
    doc.set_cell(t, doc.find_row(t, "Phiên bản phần mềm"), 1,
                 f"AI Workspace 6.11.16 (Web, ứng dụng Windows và Android) và AI Platform 6.11.18, nhánh main ngày {DATE}")
    new_section(doc, "Những điểm mới trong bản 0.14", "Những điểm mới trong bản 0.15",
                "So với bản 0.14, trang Hồ sơ cá nhân và trang Quản trị Định danh có thêm mục Ảnh FaceID trong hồ sơ:",
                "So với bản 0.13, giao diện AI Workspace không thay đổi",
                "Đối chiếu lại với AI Workspace 6.11.16 và AI Platform 6.11.1",
                [("Xem các ảnh khuôn mặt đã lưu kèm đánh giá chất lượng; xóa ảnh mờ, tối, quá nhỏ hoặc không phải bạn", "9.1"),
                 ("Quản trị viên xem, xóa ảnh khuôn mặt của người dùng", "C.2"),
                 ("Thông báo lỗi khi xóa ảnh", "12.2")])

    # ---- 2.5: hộp Cảnh báo
    box = doc.find("Cảnh báo. Khuôn mặt quét đầu tiên", "tbl")
    doc.set_note(box, "Khuôn mặt quét đầu tiên sẽ được lưu làm mẫu của tài khoản. Lần đầu đăng nhập, chính bạn phải là người "
                      "quét. Không đưa PassCode cho người khác đăng nhập hộ. Nếu nghi ngờ sai sót, hãy báo Quản trị viên để xóa "
                      "dữ liệu khuôn mặt và cấp lại PassCode. Ảnh khuôn mặt đã lưu xem và xóa được ở Hồ sơ cá nhân (mục 9.1).")

    # ---- 9.1
    t = doc.table("Bảo mật đăng nhập")
    doc.add_row(t, ["**Ảnh FaceID trong hồ sơ**", "Xem các ảnh khuôn mặt hệ thống đã lưu để so khớp khi quét FaceID; xóa ảnh không "
                    "đạt (xem bên dưới). Cần nhập lại PassCode."])
    h92 = doc.heading("9.2. Tích hợp MCP (không bắt buộc)")
    anchor = h92.getprevious()
    p = doc.para_after(anchor, "MCP là chuẩn kết nối giúp AI làm việc",
                       "**Xem và xóa ảnh FaceID.** Mỗi lần quét FaceID, hệ thống so khuôn mặt của bạn với các ảnh đã lưu trong "
                       "hồ sơ. Ảnh mờ, tối, quá nhỏ hoặc không phải bạn làm việc nhận diện kém chính xác:")
    step_tpl = doc.find("Bước 1Khi trình duyệt hỏi", "p")
    last = p
    for n, text in enumerate(STEPS, start=1):
        last = doc.step_after(last, step_tpl, n, text)
    gap = copy.deepcopy(blank)
    last.addnext(gap)
    last = doc.note_after(gap, "Lưu ý", "Bạn phải giữ lại ít nhất một ảnh dùng được (nhãn Tốt hoặc Chất lượng thấp). Muốn đăng "
                                        "ký lại khuôn mặt từ đầu, hãy nhờ Quản trị viên Xóa FaceID.",
                          blank_template=copy.deepcopy(blank))
    insert_picture(doc, last, SHOT, SHOT_TITLE, 11.0, "Ảnh màn hình 10. ")

    # ---- 12.2
    t = doc.table("Mã PassCode không chính xác hoặc tài khoản chưa được phê duyệt")
    doc.add_row(t, ["Phải giữ lại ít nhất một ảnh dùng được cho FaceID...", "Bạn đang chọn xóa mọi ảnh còn dùng được để xác thực.",
                    "Bỏ chọn ít nhất một ảnh Tốt hoặc Chất lượng thấp. Muốn đăng ký lại từ đầu, nhờ Quản trị viên Xóa FaceID."])
    doc.add_row(t, ["Ảnh FaceID vừa thay đổi...", "Một lần đăng nhập hoặc thao tác khác vừa thay đổi ảnh trong hồ sơ.",
                    "Bấm nút tải lại ở mục Ảnh FaceID trong hồ sơ rồi chọn lại."])

    # ---- C.2
    t = doc.table("Người dùng nghi lộ PassCode")
    r = doc.find_row(t, "Khuôn mặt không nhận được")
    doc.set_cell(t, r, 1, "Trước hết mở hồ sơ chi tiết của người dùng, mục **Ảnh FaceID trong hồ sơ**, xóa ảnh không đạt. Nếu "
                          "vẫn không được hoặc đăng ký nhầm người: Xóa FaceID và Cấp lại PassCode cùng lúc. Khuôn mặt quét đầu tiên "
                          "sau đó sẽ thành mẫu mới.")
    doc.add_row(t, ["Hồ sơ có ảnh khuôn mặt mờ, tối hoặc của người khác",
                    "Hồ sơ chi tiết, mục Ảnh FaceID trong hồ sơ: Chọn ảnh chưa đạt, Xóa ảnh đã chọn (nhập lại PassCode). Xóa "
                    "hết ảnh tương đương Xóa FaceID. Mỗi lần xem ảnh của người khác đều được ghi nhật ký."], like=r, after=r)

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
