# -*- coding: utf-8 -*-
"""TL2 bản 0.18 (21/11/2026): FaceID nâng độ chính xác và kiểm tra người thật (iThunder-AI-OS PR #5, ai-workspace PR #6).
- Trang bìa; mục "Những điểm mới trong bản 0.18".
- 2.4 luồng đăng nhập; 8.3 viết lại (chất lượng khung hình, so khớp, liveness, bảng biến cấu hình, hộp Lưu ý giới hạn).
- 10.1, 10.2; 11.1 Bảng 3; 11.3 sự cố; Phụ lục A.1, C (biến, giới hạn), G.
Đầu vào: TL2 bản 0.17 ĐÃ điền mục lục. Đầu ra: bản 0.18 chưa điền mục lục ở <thư mục ra>/<tên>.src.docx, sau đó chép vào
_tools/toc_work/ và chạy build_toc.py TL2.
    python update_v018.py <TL2_v017.docx> <thư mục ra>"""
import copy
import os
import sys

from docx.table import Table

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docedit import Doc  # noqa: E402
from doc_versioning import blank_para, clone_table, new_section, unfill_toc  # noqa: E402

T2 = "iThunder_AI_OS_TL2_Cai_dat_va_Van_hanh_AI_Platform.docx"
DATE = "21/11/2026"
POINTS = [
    "**Khung hình**: AI Workspace cắt vùng vuông giữa ảnh camera (không kéo giãn 4:3 thành 1:1), gửi 320×320 khoảng 3 khung "
    "mỗi giây qua /ws/auth, chỉ gửi khung mới khi khung trước đã có kết quả. AI Platform xử lý ảnh trong bộ nhớ, ở luồng phụ.",
    "**Chất lượng** (identity/faceid.py): đúng một khuôn mặt, đủ lớn, ở giữa khung, đủ sáng, không mờ. Khung chưa đạt trả "
    "status HINT kèm lời nhắc, không tính là lần thất bại.",
    "**So khớp**: chỉ mã hóa khung nhìn thẳng (FACEID_NUM_JITTERS lần lấy mẫu), so với mọi mẫu của người dùng, lấy khoảng cách "
    "nhỏ nhất. Ngưỡng là iDAccountPrimaryType, mặc định 0,55, không thấp hơn 0,5. Người dùng bản cũ chỉ có ảnh: ảnh được mã "
    "hóa một lần và ghi vào initials. Sau lần đăng nhập khớp rõ ràng mà đủ khác các mẫu sẵn có, mẫu mới (kèm ảnh) được bổ sung, "
    "tối đa FACEID_MAX_TEMPLATES.",
    "**Kiểm tra người thật**: khung nhìn thẳng đầu tiên khớp mẫu, rồi làm FACEID_LIVENESS_STEPS thử thách ngẫu nhiên (quay trái, "
    "quay phải, nhắm mắt rồi mở; status LIVENESS kèm hướng dẫn), cuối cùng nhìn thẳng và khớp lại, cùng người với khung đầu. "
    "AI Platform phát và tự kiểm tra thử thách; máy trạm chỉ hiển thị.",
    "**Chống ảnh phẳng**: khi xoay, nghiêng một mặt phẳng (ảnh in, màn hình), vị trí đầu mũi giữa hai bên hàm và giữa hai khóe "
    "mắt gần như không đổi; đầu người thật có mũi nhô ra nên lệch hẳn (thị sai). Theo mô phỏng, ảnh phẳng xoay tới 65° sát camera "
    "15 cm chỉ lệch 0,05, dưới ngưỡng 0,18. Khung mặt nhảy vị trí, đổi cỡ đột ngột (tráo ảnh) hay quá thời hạn thì làm lại với "
    "thử thách mới.",
    "**Đăng ký lần đầu** cũng phải qua liveness; lưu ảnh và mẫu của khung đầu, khung cuối. Mỗi FACEID_MAX_MISMATCHES (10) khung "
    "không khớp trên một kết nối tính một lần sai vào giới hạn đăng nhập theo IP. Nhật ký XML ghi thêm Liveness, Duration.",
]
CONFIG = [
    ("FACEID_LIVENESS", "required (mặc định): máy trạm không gửi liveness=1 nhận CLIENT_UPDATE_REQUIRED; optional: chỉ với "
                        "máy trạm mới (giai đoạn chuyển tiếp, máy trạm cũ chưa chống được ảnh tĩnh); off: tắt"),
    ("FACEID_LIVENESS_STEPS, FACEID_LIVENESS_TIMEOUT", "Số thử thách (2, tối đa 3); giây cho cả lượt (40)"),
    ("FACEID_DEFAULT_ACCURACY, FACEID_MIN_ACCURACY", "Ngưỡng khi người dùng chưa có ngưỡng riêng (0,55); ngưỡng thấp nhất (0,5)"),
    ("FACEID_NUM_JITTERS, FACEID_MAX_TEMPLATES", "Số lần lấy mẫu khi mã hóa (2); số mẫu tối đa mỗi người dùng (5)"),
    ("FACEID_YAW_DELTA, FACEID_PARALLAX_DELTA", "Độ quay đầu cần đạt (0,12); độ lệch đầu mũi so với khóe mắt cần đạt (0,18, "
                                                "không nên dưới 0,12)"),
]


def tl2(src, dst):
    doc = Doc(src)
    unfill_toc(doc)
    blank = blank_para(doc)
    t = doc.table("Phiên bản tài liệu")
    doc.set_cell(t, doc.find_row(t, "Phiên bản tài liệu"), 1,
                 f"0.18, bản dự thảo ngày {DATE} (cập nhật bản 0.17: FaceID có kiểm tra người thật)")
    doc.set_cell(t, doc.find_row(t, "Phiên bản phần mềm"), 1,
                 f"AI Workspace 6.11.16 và AI Platform 6.11.18 (nhánh main ngày {DATE}, sau khi hợp nhất PR #5, #6, #7 của "
                 "iThunder-AI-OS và PR #6, #7, #8 của ai-workspace: FaceID liveness, Kết nối Kho Đám mây, Ảnh FaceID trong hồ sơ)")
    new_section(doc, "Những điểm mới trong bản 0.17", "Những điểm mới trong bản 0.18",
                "Bổ sung kiểm tra người thật (liveness) và nâng độ chính xác khi xác thực FaceID (iThunder-AI-OS PR #5, "
                "ai-workspace PR #6). Các thay đổi và nơi mô tả:",
                "Bổ sung chức năng xem, đánh giá chất lượng",
                "Ảnh FaceID trong hồ sơ: đánh giá chất lượng",
                [("Luồng đăng nhập: khung hình vuông 3 khung mỗi giây, thử thách ngẫu nhiên", "2.4"),
                 ("Chất lượng khung hình, so khớp nhiều mẫu, kiểm tra người thật, biến cấu hình, giới hạn", "8.3, C"),
                 ("Cơ chế an toàn, điểm cần tăng cường, xử lý sự cố, kiểm thử", "10.1, 10.2, 11.1, 11.3"),
                 ("Thông điệp của WS /ws/auth", "A.1")])

    # ---- 2.4
    p = doc.find("Bước 4Nếu không, Backend trả REQUIRE_FACEID", "p")
    doc.step_after(p, p, 4, "Nếu không, Backend trả REQUIRE_FACEID. Workspace mở WebSocket /ws/auth và gửi khung hình vuông "
                            "320×320 khoảng 3 khung mỗi giây (chỉ khi khung trước đã có kết quả) qua hành động verify_faceid, "
                            "kèm liveness=1.")
    p.getparent().remove(p)
    p = doc.find("Bước 5Backend so sánh khuôn mặt bằng face_recognition", "p")
    doc.step_after(p, p, 5, "Backend kiểm tra chất lượng khung, so khớp khung nhìn thẳng với mọi mẫu (ngưỡng iDAccountPrimaryType, "
                            "mặc định 0,55; tối thiểu 0,5), phát hai thử thách ngẫu nhiên để kiểm tra người thật, rồi so khớp lại "
                            "khung nhìn thẳng cuối (mục 8.3). Tài khoản chưa có mẫu cũng phải qua các bước này; ảnh khung đầu và "
                            "khung cuối được lưu làm mẫu.")
    p.getparent().remove(p)

    # ---- 8.3
    p = doc.find("Dùng face_recognition (dlib) và OpenCV.", "p")
    doc.set_para(p, "Dùng face_recognition (dlib) và OpenCV. Ảnh xác thực lưu ở FaceID/AccessStorage/<năm>/<tháng>/<ngày>/ "
                    "kèm tệp XML, ảnh không khớp lưu ở BadImages/. Ảnh và đặc trưng khuôn mặt là dữ liệu sinh trắc học: cần có "
                    "chính sách lưu giữ, truy cập, xóa và căn cứ pháp lý.")
    intro = doc.para_after(p, p, "[[ĐÃ CÓ]] Từ ngày 21/11/2026 (iThunder-AI-OS PR #5, ai-workspace PR #6), mỗi lượt quét FaceID "
                                 "là một phiên gắn với kết nối /ws/auth và PassCode:")
    bullet_tpl = doc.find("Khóa là DN. gidNumber", "p")
    last = intro
    for text in POINTS:
        last = doc.para_after(last, bullet_tpl, text)
    gap = copy.deepcopy(blank)
    last.addnext(gap)
    tbl_el, tbl = clone_table(doc, "Sự kiệnHệ thống làm gì", gap, 2)
    doc.replace_in(tbl.cell(0, 0)._tc, "Sự kiện", "Biến (.env)")
    doc.replace_in(tbl.cell(0, 1)._tc, "Hệ thống làm gì", "Ý nghĩa, mặc định")
    doc.set_cell(tbl, 1, 0, CONFIG[0][0], code=True)
    doc.set_cell(tbl, 1, 1, CONFIG[0][1], code=False)
    for name, desc in CONFIG[1:]:
        idx = doc.add_row(tbl, [name, desc])
        doc.set_cell(tbl, idx, 0, name, code=True, note="-")
        doc.set_cell(tbl, idx, 1, desc, code=False, note="-")
    # Cột tên biến rộng hơn (bảng mẫu có cột đầu hẹp): 42% / 58% tổng bề rộng
    from docx.oxml.ns import qn
    cols = tbl_el.find(qn("w:tblGrid")).findall(qn("w:gridCol"))
    total = sum(int(c.get(qn("w:w"))) for c in cols)
    widths = [int(total * 0.42), total - int(total * 0.42)]
    for c, w in zip(cols, widths):
        c.set(qn("w:w"), str(w))
    for tr in tbl_el.findall(qn("w:tr")):
        for tc, w in zip(tr.findall(qn("w:tc")), widths):
            tcw = tc.find(qn("w:tcPr")).find(qn("w:tcW"))
            if tcw is not None:
                tcw.set(qn("w:w"), str(w))
                tcw.set(qn("w:type"), "dxa")
    gap2 = copy.deepcopy(blank)
    tbl_el.addnext(gap2)
    doc.note_after(gap2, "Lưu ý", "Liveness chủ động chống ảnh tĩnh 2D; chưa chống được mặt nạ 3D, video giả mạo tạo theo thời "
                                  "gian thực hay camera ảo. Ngưỡng thị sai mới được kiểm bằng mô phỏng hình học 3D: nếu người dùng "
                                  "thật khó qua bước quay đầu, giảm FACEID_PARALLAX_DELTA hoặc FACEID_YAW_DELTA. Ứng dụng Windows, "
                                  "Android bản 6.11.16 phải build lại; trong lúc chờ có thể đặt FACEID_LIVENESS=optional.",
                   blank_template=copy.deepcopy(blank))

    # ---- 10.1, 10.2
    p = doc.find("Ảnh FaceID: xem ảnh của người khác cần quyền identity.security.reset", "p")
    doc.para_after(p, p, "FaceID có kiểm tra người thật do AI Platform phát thử thách ngẫu nhiên và tự kiểm tra; kiểm tra chất "
                         "lượng khung, chống tráo ảnh giữa chừng; khung không khớp lặp lại được tính vào giới hạn đăng nhập "
                         "theo IP (mục 8.3).")
    t = doc.table("Ưu tiênVấn đềKhuyến nghị")
    r = doc.find_row(t, "Ảnh đầu tiên tự thành mẫu FaceID")
    doc.set_cell(t, r, 1, "Liveness chủ động (mục 8.3) chống ảnh tĩnh 2D; chưa chống mặt nạ 3D, video giả mạo thời gian thực, "
                          "camera ảo. Ngưỡng chưa kiểm trên nhiều người dùng thật", code=False)
    doc.set_cell(t, r, 2, "Thử với người dùng thật, hiệu chỉnh ngưỡng; thêm mô hình chống giả mạo thụ động (onnxruntime đã có); "
                          "đăng ký khuôn mặt có xác minh. Ảnh mẫu xem, xóa được (mục 6.7)", code=False)

    # ---- 11.1, 11.3
    t = doc.table("Bộ kiểm thửKết quả ngày")
    r = doc.find_row(t, "tests/identity/run_all.py")
    doc.set_cell(t, r, 1, "Đạt (18 tệp, gồm test_14_rag_chat.py so sánh trực tiếp lời nhắc của AI Chat/Work và Trao đổi trực "
                          "tuyến; test_16_faceid_liveness.py mô phỏng hình học đầu người thật và ảnh phẳng; test_17, test_18 "
                          "ngày 21/11/2026)", code=False)
    t = Table(doc.heading("11.3. Xử lý sự cố thường gặp").getnext(), doc.d)
    for values in [
        ("Người dùng thật không qua bước quay mặt", "Ngưỡng thị sai cao so với camera, khoảng cách, góc nhìn",
         "Giảm FACEID_PARALLAX_DELTA (không dưới 0,12) hoặc FACEID_YAW_DELTA trong .env, khởi động lại"),
        ("Máy trạm báo AI Gateway yêu cầu kiểm tra người thật", "Bản Web cũ trong bộ đệm, ứng dụng 6.11.16 chưa build lại",
         "Triển khai bản Web mới (no-store cho index.html); build lại ứng dụng; tạm thời FACEID_LIVENESS=optional"),
    ]:
        doc.add_row(t, list(values))

    # ---- Phụ lục A.1, C, G
    t = doc.table("POST /api/auth/loginĐăng nhập bằng PassCode")
    r = doc.find_row(t, "WS /ws/auth")
    doc.set_cell(t, r, 1, "Xác thực FaceID theo khung ảnh (verify_faceid, liveness=1). Phản hồi: HINT (khung chưa đạt), LIVENESS "
                          "(bước kế tiếp), FAIL (FACE_MISMATCH, LIVENESS_TIMEOUT, FACE_SWITCHED), SUCCESS; máy trạm cũ nhận "
                          "ERROR CLIENT_UPDATE_REQUIRED", code=False)
    t = doc.table("NhómBiến trong .env")
    doc.add_row(t, ["FaceID", "FACEID_LIVENESS, FACEID_LIVENESS_STEPS, FACEID_LIVENESS_TIMEOUT, FACEID_DEFAULT_ACCURACY, "
                    "FACEID_MIN_ACCURACY, FACEID_NUM_JITTERS, FACEID_MAX_TEMPLATES, FACEID_YAW_DELTA, FACEID_PARALLAX_DELTA "
                    "(mục 8.3)"])
    t = doc.table("Giới hạn trong config.py")
    doc.add_row(t, ["FaceID", "Khung 320×320, khoảng 3 khung mỗi giây; 2 thử thách, 40 giây mỗi lượt (máy trạm dừng sau 90 giây); "
                    "tối đa 5 mẫu mỗi người dùng; 10 khung không khớp tính một lần sai"])
    t = doc.table("Căn cứ chính của nội dung hiện trạng")
    r = doc.find_row(t, "Căn cứ chính của nội dung hiện trạng")
    doc.set_cell(t, r, 0, f"Mã nguồn ai-platform (AI Platform 6.11.18), ai-work (AI Workspace 6.11.16), nhánh main ngày {DATE} "
                          "sau khi hợp nhất FaceID liveness (iThunder-AI-OS PR #5, ai-workspace PR #6), Kết nối Kho Đám mây "
                          "(PR #6, PR #7) và Ảnh FaceID trong hồ sơ (PR #7, PR #8)", code=False)
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
