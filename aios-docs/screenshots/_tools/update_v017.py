# -*- coding: utf-8 -*-
"""TL2 bản 0.17 (21/11/2026): Ảnh FaceID trong hồ sơ định danh (iThunder-AI-OS PR #7, ai-workspace PR #8).
- Trang bìa; mục "Những điểm mới trong bản 0.17".
- 6.7 (mới): đánh giá chất lượng, xóa ảnh và tính lại mẫu khuôn mặt, quyền, ảnh màn hình; 6.4 step-up; 6.6 nhật ký; 8.3.
- 10.1, 10.2; 11.1 Bảng 3; Phụ lục A.3 (API), G (căn cứ mã nguồn). Đánh lại số ảnh màn hình.
Đầu vào: TL2 bản 0.16 ĐÃ điền mục lục. Đầu ra: bản 0.17 chưa điền mục lục ở <thư mục ra>/<tên>.src.docx, sau đó chép vào
_tools/toc_work/ và chạy build_toc.py TL2.
    python update_v017.py <TL2_v016.docx> <thư mục ra>"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docedit import Doc  # noqa: E402
from doc_versioning import clone_heading, insert_picture, new_section, renumber_shots, unfill_toc  # noqa: E402

T2 = "iThunder_AI_OS_TL2_Cai_dat_va_Van_hanh_AI_Platform.docx"
DATE = "21/11/2026"
SHOT = os.path.join(HERE, "..", "tl2_06_faceid_photos.png")
SHOT_TITLE = ("Quản trị Định danh, ngăn chi tiết người dùng, mục Ảnh FaceID trong hồ sơ: ảnh kèm đánh giá chất lượng (ảnh "
              "khuôn mặt được thay bằng hình minh họa)")
RULES = [
    "**Đánh giá** (identity/face_photos.py, dlib): đúng một khuôn mặt, bề rộng mặt từ 80 px, độ nét (phương sai Laplacian trên "
    "vùng mặt chuẩn hóa), độ sáng, độ khớp với các ảnh còn lại (dưới 0,5 là có thể người khác). Kết quả: Tốt, Chất lượng thấp, "
    "Không dùng được. API chỉ trả ảnh thu nhỏ 240 px, không trả ảnh gốc hay mã đặc trưng.",
    "**Xóa**: ảnh chỉ định bằng mã SHA-256 (16 ký tự). Mẫu khuôn mặt (initials) được tính lại từ chính các ảnh còn lại, như "
    "deletePicture của AI-LMS dựng lại jpegPhoto, nên ảnh đã xóa không còn dùng để so khớp. Ảnh trên OpenLDAP vừa đổi kể từ "
    "lúc đọc (có lần đăng nhập thêm ảnh) thì từ chối ghi (409).",
    "**Người dùng tự xóa** (Hồ sơ của tôi): cần step-up, phải giữ ít nhất một ảnh dùng được, để tài khoản không trở về trạng "
    "thái chưa đăng ký (khi đó khuôn mặt bất kỳ đi kèm PassCode thành FaceID mới).",
    "**Quản trị viên** (quyền identity.security.reset, theo phạm vi đơn vị): xem ảnh của người dùng (ghi nhật ký "
    "user.faceid.photos.view), xóa cần step-up; xóa hết ảnh tương đương Xóa FaceID, xóa cả mốc FaceID của Tổ chức.",
]


def tl2(src, dst):
    doc = Doc(src)
    unfill_toc(doc)
    t = doc.table("Phiên bản tài liệu")
    doc.set_cell(t, doc.find_row(t, "Phiên bản tài liệu"), 1,
                 f"0.17, bản dự thảo ngày {DATE} (cập nhật bản 0.16: Ảnh FaceID trong hồ sơ định danh)")
    doc.set_cell(t, doc.find_row(t, "Phiên bản phần mềm"), 1,
                 f"AI Workspace 6.11.16 và AI Platform 6.11.18 (nhánh main ngày {DATE}, sau khi hợp nhất PR #6, PR #7 của "
                 "iThunder-AI-OS và PR #7, PR #8 của ai-workspace: Kết nối Kho Đám mây, Ảnh FaceID trong hồ sơ)")
    new_section(doc, "Những điểm mới trong bản 0.16", "Những điểm mới trong bản 0.17",
                "Bổ sung chức năng xem, đánh giá chất lượng và xóa ảnh FaceID trong hồ sơ định danh (iThunder-AI-OS PR #7, "
                "ai-workspace PR #8). Các thay đổi và nơi mô tả:",
                "Bổ sung chức năng Kết nối Kho Dữ liệu",
                "Kết nối Kho Dữ liệu với kho đám mây: dịch vụ hỗ trợ",
                [("Ảnh FaceID trong hồ sơ: đánh giá chất lượng, xóa ảnh, tính lại mẫu khuôn mặt, quyền", "6.7"),
                 ("Step-up, nhật ký, cơ chế an toàn", "6.4, 6.6, 10.1"),
                 ("API /api/identity/.../faceid/photos; kết quả kiểm thử", "A.3, 11.1")])

    # ---- 6.7 mới, trước chương 7
    h7 = doc.heading("7. Kho Dữ liệu và Kho Tri thức")
    h = clone_heading(doc, "6.6. Nhật ký", h7.getprevious(), "6.7. Ảnh FaceID trong hồ sơ")
    intro = doc.para_after(h, "Thao tác quản trị được ghi vào bảng identity_audit",
                           "[[ĐÃ CÓ]] Ảnh FaceID (jpegPhoto, nhiều giá trị) là nguồn của mẫu khuôn mặt (initials) dùng khi xác "
                           "thực. Người dùng (tab **Hồ sơ của tôi**) và Quản trị viên (ngăn chi tiết người dùng của trang Quản "
                           "trị Định danh) xem được mọi ảnh kèm đánh giá chất lượng, rồi xóa ảnh không đạt (nút **Chọn ảnh "
                           "chưa đạt**, **Xóa ảnh đã chọn**).")
    bullet_tpl = doc.find("Khóa là DN. gidNumber", "p")
    last = intro
    for rule in RULES:
        last = doc.para_after(last, bullet_tpl, rule)
    insert_picture(doc, last, SHOT, SHOT_TITLE, 16.0, "Ảnh màn hình 2. ")

    # ---- 6.4 step-up, 6.6 nhật ký, 8.3
    p = doc.find("Các thao tác nhạy cảm trả HTTP 428", "p")
    doc.replace_in(p, "xóa FaceID", "xóa FaceID, xóa ảnh FaceID (cả khi người dùng tự xóa)")
    p = doc.find("Thay đổi danh mục MCP Server bên ngoài cũng được ghi", "p")
    doc.para_after(p, p, "Xem ảnh FaceID của người khác ghi user.faceid.photos.view; xóa ảnh ghi user.faceid.photos.delete "
                         "hoặc self.faceid.photos.delete, kèm số ảnh trước và sau.")
    p = doc.find("Dùng face_recognition (dlib) và OpenCV.", "p")
    doc.para_after(p, p, "Ảnh mẫu của từng người dùng xem và xóa được ở mục 6.7; sau khi xóa, mẫu khuôn mặt được tính lại từ "
                         "các ảnh còn lại.")

    # ---- 10.1, 10.2
    p = doc.find("Kết nối Kho đám mây: thông tin đăng nhập mã hóa", "p")
    doc.para_after(p, p, "Ảnh FaceID: xem ảnh của người khác cần quyền identity.security.reset và được ghi nhật ký, API chỉ trả "
                         "ảnh thu nhỏ; xóa ảnh cần step-up; người dùng không tự xóa được ảnh dùng được cuối cùng (mục 6.7).")
    t = doc.table("Ưu tiênVấn đềKhuyến nghị")
    r = doc.find_row(t, "Ảnh đầu tiên tự thành mẫu FaceID")
    doc.set_cell(t, r, 2, "Đăng ký khuôn mặt có xác minh; thêm liveness. Đã có: xem, xóa ảnh mẫu không đạt (mục 6.7)", code=False)

    # ---- 11.1
    t = doc.table("Bộ kiểm thửKết quả ngày")
    r = doc.find_row(t, "tests/identity/run_all.py")
    doc.set_cell(t, r, 1, "Đạt (17 tệp, gồm test_14_rag_chat.py so sánh trực tiếp lời nhắc của AI Chat/Work và Trao đổi trực "
                          "tuyến, test_17_storage_link_hooks.py, test_18_faceid_photos.py ngày 21/11/2026)", code=False)

    # ---- Phụ lục A.3, G
    t = doc.table("GET /me/permissions, /me/profile")
    for i, values in enumerate([("GET /me/faceid/photos; POST /me/faceid/photos/delete", "Đăng nhập", "Khi xóa"),
                                ("GET /users/{uid}/faceid/photos; POST /users/{uid}/faceid/photos/delete", "security.reset",
                                 "Khi xóa")]):
        doc.add_row(t, list(values), like=1, after=1 + i)
    t = doc.table("Căn cứ chính của nội dung hiện trạng")
    r = doc.find_row(t, "Căn cứ chính của nội dung hiện trạng")
    doc.set_cell(t, r, 0, f"Mã nguồn ai-platform (AI Platform 6.11.18), ai-work (AI Workspace 6.11.16), nhánh main ngày {DATE} "
                          "sau khi hợp nhất Kết nối Kho Đám mây (iThunder-AI-OS PR #6, ai-workspace PR #7) và Ảnh FaceID trong "
                          "hồ sơ (iThunder-AI-OS PR #7, ai-workspace PR #8)", code=False)

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
