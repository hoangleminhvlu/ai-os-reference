# -*- coding: utf-8 -*-
"""TL1 bản 0.20 (23/11/2026): quy trình FaceID mới (iThunder-AI-OS PR #9, ai-workspace PR #16).
- Trang bìa; mục "Những điểm mới trong bản 0.20".
- 2.4, 2.5: Web kiểm tra người thật; ứng dụng từ bản 6.11.23 không cần, thiết bị mới nhập mã gửi qua Email; PassTime dùng chung.
- 2.8 (hộp Lưu ý), 9.1 (Thiết bị FaceID), 12.1, 12.2, C.2, Phụ lục thuật ngữ. Hai ảnh màn hình mới từ ứng dụng Windows 6.11.23 thật.
Đầu vào: TL1 bản 0.19 ĐÃ điền mục lục. Đầu ra: bản 0.20 chưa điền mục lục ở <thư mục ra>/<tên>.src.docx, sau đó chép vào
_tools/toc_work/ và chạy build_toc.py TL1.
    python update_tl1_v020.py <TL1_v019.docx> <thư mục ra>"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docedit import Doc  # noqa: E402
from doc_versioning import blank_para, insert_picture, new_section, renumber_shots, unfill_toc  # noqa: E402

T1 = "iThunder_AI_OS_TL1_Huong_dan_su_dung_AI_Workspace.docx"
DATE = "23/11/2026"
APP = "6.11.23"
SHOT_CONFIRM = os.path.join(HERE, "..", "tl1_13_device_confirm.png")
SHOT_DEVICES = os.path.join(HERE, "..", "tl1_14_faceid_devices.png")


def step(doc, needle, n, markup):
    p = doc.find(needle, "p")
    new = doc.step_after(p, p, n, markup)
    p.getparent().remove(p)
    return new


def tl1(src, dst):
    doc = Doc(src)
    unfill_toc(doc)
    t = doc.table("Phiên bản tài liệu")
    doc.set_cell(t, doc.find_row(t, "Phiên bản tài liệu"), 1,
                 f"0.20, bản dự thảo ngày {DATE} (cập nhật bản 0.19: quy trình FaceID mới, xác nhận thiết bị mới qua Email)")
    doc.set_cell(t, doc.find_row(t, "Phiên bản phần mềm"), 1,
                 f"AI Workspace Web nhánh main ngày {DATE}, ứng dụng Windows và Android 6.11.21 (trang phát hành) và {APP} (bản "
                 "thử, chưa đăng trang phát hành), AI Platform 6.11.18")
    new_section(doc, "Những điểm mới trong bản 0.19", "Những điểm mới trong bản 0.20",
                "So với bản 0.19, cách quét khuôn mặt khác nhau giữa bản Web và ứng dụng:",
                "So với bản 0.18, ứng dụng Android",
                "Tải ứng dụng Android 6.11.21 kèm tài liệu này",
                [(f"Bản Web vẫn kiểm tra người thật; ứng dụng Windows, Android từ bản {APP} chỉ cần nhìn thẳng", "2.5"),
                 ("Thiết bị mới phải nhập mã xác nhận 8 số gửi qua Email, một lần cho mỗi thiết bị", "2.5"),
                 ("Xem, thu hồi, khôi phục Thiết bị FaceID; Quản trị viên xử lý thiết bị bị mất", "9.1, C.2"),
                 ("Một ngưỡng khớp khuôn mặt chung cho mọi người; PassTime dùng chung cho Web và ứng dụng", "2.5"),
                 ("Thông báo mới khi xác nhận thiết bị", "12.2")])

    # ---- 2.4
    step(doc, "Nếu thời hạn miễn quét (PassTime) còn hiệu lực, bạn vào thẳng AI Workspace.", 3,
         "Nếu thời hạn miễn quét (PassTime) còn hiệu lực, bạn vào thẳng AI Workspace, dù lần quét gần nhất là trên bản Web hay trên "
         "ứng dụng. Nếu không, màn hình quét khuôn mặt sẽ hiện ra.")

    # ---- 2.5
    h = doc.heading("2.5. Quét khuôn mặt (FaceID)")
    doc.para_after(h, "PassTime là khoảng thời gian được miễn quét khuôn mặt",
                   "Trên bản Web, hệ thống kiểm tra người thật bằng các thử thách (Bước 3). Trên ứng dụng Windows, Android từ bản "
                   f"{APP} [[ĐANG PHÁT TRIỂN]], bạn chỉ cần nhìn thẳng vào camera (bỏ qua Bước 3 và Bước 4); lần đầu trên mỗi thiết bị "
                   "phải xác nhận qua Email như hướng dẫn ở cuối mục này. Ngưỡng khớp khuôn mặt là như nhau cho mọi người.")
    step(doc, "Làm theo hướng dẫn kiểm tra người thật hiện phía trên khung tròn", 3,
         "Bản Web: làm theo hướng dẫn kiểm tra người thật hiện phía trên khung tròn, gồm hai thử thách chọn ngẫu nhiên trong ba thử "
         "thách từ từ quay mặt sang trái, sang phải của bạn, nhắm mắt khoảng một giây rồi mở. Ảnh trong khung là ảnh gương, nên quay "
         "sang trái của bạn cũng là sang trái trên màn hình. Thanh tiến độ cho biết đã xong mấy bước.")
    step(doc, "Cuối cùng nhìn thẳng lại vào camera. Khi khớp", 4,
         "Bản Web: cuối cùng nhìn thẳng lại vào camera. Khi khớp, bạn tự động vào AI Workspace.")
    box = doc.find("Ảnh in, ảnh trên điện thoại hay màn hình không qua được bước kiểm tra người thật", "tbl")
    doc.set_note(box, "Trên bản Web, ảnh in, ảnh trên điện thoại hay màn hình không qua được bước kiểm tra người thật; lần đầu đăng ký "
                      "khuôn mặt cũng phải làm đủ các bước. Ứng dụng 6.11.21 trở về trước cũng phải làm các bước này. Nếu thấy AI "
                      "Gateway yêu cầu kiểm tra người thật, hãy tải lại trang (bản Web) hoặc cài bản ứng dụng mới.")
    p = doc.find("PassTime là khoảng thời gian được miễn quét khuôn mặt", "p", nth=0)
    doc.set_para(p, "PassTime là khoảng thời gian được miễn quét khuôn mặt sau lần xác thực gần nhất, mặc định 5 giờ và do Quản trị "
                    "viên thiết lập; dùng chung cho bản Web và ứng dụng. Ngoài ra, mỗi phiên đăng nhập kéo dài tối đa 8 giờ; khi hết "
                    "hạn, hệ thống sẽ yêu cầu bạn đăng nhập lại.")
    blank = blank_para(doc)
    import copy
    intro = doc.para_after(p, p, f"**Xác nhận thiết bị mới trên ứng dụng** (từ bản {APP}) [[ĐANG PHÁT TRIỂN]]. Mỗi máy tính, điện thoại "
                                 "cài ứng dụng có một khóa riêng. Lần đầu bạn đăng nhập trên thiết bị, sau khi khuôn mặt khớp:")
    s_tpl = doc.find("Nếu chưa khớp sau nhiều lần, hoặc quá 90 giây", "p")
    s1 = doc.step_after(intro, s_tpl, 1, "Hệ thống gửi mã xác nhận 8 số tới các Email đăng ký của bạn (màn hình chỉ hiện Email đã che bớt). "
                                         "Thư ghi tên thiết bị, địa chỉ mạng và thời gian. Kiểm tra cả thư mục Spam.")
    s2 = doc.step_after(s1, s_tpl, 2, "Nhập mã vào ô trên chính thiết bị đó rồi bấm **Đăng ký thiết bị**. Mã có hiệu lực 10 phút, nhập sai tối "
                                      "đa 5 lần. Bạn có thể mở ứng dụng thư rồi quay lại nhập mã. Chưa nhận được thư thì sau 60 giây bấm "
                                      "**Gửi lại mã**.")
    s3 = doc.step_after(s2, s_tpl, 3, "Mã đúng thì thiết bị được đăng ký; nhìn thẳng vào camera lần nữa để vào AI Workspace. Những lần sau "
                                      "trên thiết bị này không cần nhập mã.")
    cap = insert_picture(doc, s3, SHOT_CONFIRM, f"Ứng dụng Windows {APP} (bản thử): nhập mã xác nhận thiết bị mới gửi qua Email (dữ liệu mẫu)",
                         9.0, "Ảnh màn hình 2. ")
    b1 = copy.deepcopy(blank)
    cap.addnext(b1)
    doc.note_after(b1, "Cảnh báo", "Không đưa mã xác nhận cho bất kỳ ai, kể cả người tự xưng là Quản trị viên. Nhận được thư xác nhận mà bạn "
                                   "không đăng nhập: báo ngay Quản trị viên để cấp lại PassCode. Hãy bảo vệ hộp thư của bạn (bật xác thực "
                                   "hai lớp), vì PassCode và mã xác nhận đều gửi qua Email.", blank_template=blank)

    # ---- 2.8
    box = doc.find("Bản thử nghiệm đã được kiểm tra cài đặt, gỡ cài đặt", "tbl")
    doc.set_note(box, "Bản thử nghiệm đã được kiểm tra cài đặt, gỡ cài đặt, cài bản mới đè lên bản cũ, mở ứng dụng, kết nối AI Gateway và "
                      f"hiển thị màn hình làm việc. Bản thử {APP} đã được kiểm tra đăng nhập và xác nhận thiết bị mới với dữ liệu mẫu; "
                      "quét khuôn mặt bằng camera thật, gọi thoại và video trên ứng dụng đang được kiểm thử thêm. Nếu gặp lỗi, hãy báo "
                      "Quản trị viên kèm ảnh chụp màn hình và phiên bản ứng dụng.")

    # ---- 9.1
    t = doc.table("Xem các ảnh khuôn mặt hệ thống đã lưu để so khớp khi quét FaceID")
    r = doc.find_row(t, "Xem các ảnh khuôn mặt hệ thống đã lưu")
    doc.add_row(t, ["Thiết bị FaceID", "Xem các máy tính, điện thoại đã đăng ký; thu hồi thiết bị bị mất hoặc lạ, khôi phục khi thu hồi nhầm "
                                       "(xem bên dưới). Cần nhập lại PassCode."], like=r, after=r)
    cap10 = doc.find("Hồ sơ cá nhân, mục Ảnh FaceID trong hồ sơ: đã chọn các ảnh chưa đạt", "p")
    tpl = doc.find("Xem và xóa ảnh FaceID. Mỗi lần quét FaceID", "p")
    st = doc.find("Hệ thống tính lại mẫu khuôn mặt từ các ảnh còn lại", "p")
    q = doc.para_after(cap10, tpl, "**Xem và thu hồi thiết bị FaceID.** Danh sách gồm các ứng dụng Windows, Android đã xác nhận qua Email "
                                   "(mục 2.5), kèm thời điểm đăng ký, lần dùng gần nhất và địa chỉ mạng:")
    q = doc.step_after(q, st, 1, "Mở **Hồ sơ cá nhân**, kéo xuống mục **Thiết bị FaceID**.")
    q = doc.step_after(q, st, 2, "Ở thiết bị bị mất hoặc bạn không nhận ra, bấm **Thu hồi**, nhập lại PassCode. Thiết bị đó không quét FaceID "
                                 "đăng nhập được nữa.")
    q = doc.step_after(q, st, 3, "Thu hồi nhầm: bấm **Khôi phục**.")
    cap = insert_picture(doc, q, SHOT_DEVICES, f"Hồ sơ cá nhân, mục Thiết bị FaceID, trên ứng dụng Windows {APP} (bản thử, dữ liệu mẫu)",
                         15.0, "Ảnh màn hình 2. ")
    b2 = copy.deepcopy(blank)
    cap.addnext(b2)
    doc.note_after(b2, "Lưu ý", "Thu hồi không ngắt phiên đang mở và thời hạn PassTime còn lại. Mất thiết bị hoặc nghi lộ PassCode: nhờ Quản "
                                "trị viên cấp lại PassCode ngay.", blank_template=blank)

    # ---- 12.2
    t = doc.table("AI Gateway yêu cầu kiểm tra người thật khi quét FaceID...")
    r = doc.find_row(t, "AI Gateway yêu cầu kiểm tra người thật khi quét FaceID")
    doc.set_cell(t, r, 2, "Tải lại trang (bản Web) hoặc cài bản ứng dụng mới; trong lúc chờ, làm theo các bước kiểm tra người thật.")
    r = doc.find_row(t, "Khuôn mặt quét được không khớp với Hồ sơ")
    rows = [["Thiết bị mới: đã gửi mã xác nhận 8 số...", "Bạn đăng nhập lần đầu trên ứng dụng ở thiết bị này.",
             "Mở Email đăng ký, nhập mã trên chính thiết bị (mục 2.5)."],
            ["Mã xác nhận không đúng. Còn ... lần nhập.", "Nhập sai mã, hoặc dùng mã của thư cũ.", "Nhập mã trong thư mới nhất."],
            ["Mã xác nhận đã hết hạn hoặc nhập sai quá số lần...", "Quá 10 phút hoặc sai 5 lần.", "Bấm Quét lại khuôn mặt để nhận mã mới."],
            ["Đã có quá nhiều yêu cầu gửi Email...", "Đã yêu cầu nhiều mã trong thời gian ngắn.", "Kiểm tra hộp thư (cả Spam) hoặc chờ 15 phút."],
            ["Tài khoản chưa có Email đăng ký nên không xác nhận được thiết bị mới...", "Hồ sơ chưa có Email.", "Nhờ Quản trị viên thêm Email."],
            ["Thiết bị này đã bị thu hồi quyền đăng nhập FaceID...", "Bạn hoặc Quản trị viên đã thu hồi thiết bị.",
             "Nếu thu hồi nhầm: đăng nhập bản Web, vào Hồ sơ cá nhân, mục Thiết bị FaceID, bấm Khôi phục; hoặc liên hệ Quản trị viên."]]
    for x in rows:
        r = doc.add_row(t, x, like=r, after=r)

    # ---- C.2
    t = doc.table("Khuôn mặt không nhận được, hoặc đăng ký nhầm người")
    r = doc.find_row(t, "Hồ sơ có ảnh khuôn mặt mờ, tối hoặc của người khác")
    doc.add_row(t, ["Người dùng mất máy tính, điện thoại đã đăng ký", "Hồ sơ chi tiết, mục Thiết bị FaceID: Thu hồi (nhập lại PassCode), rồi "
                    "Cấp lại PassCode. Đăng ký, thu hồi, khôi phục thiết bị đều được ghi nhật ký."], like=r, after=r)

    # ---- Thuật ngữ
    t = doc.table("Bước của FaceID yêu cầu quay mặt, nhắm mắt")
    r = doc.find_row(t, "Kiểm tra người thật (liveness)")
    doc.set_cell(t, r, 1, "Bước của FaceID trên bản Web yêu cầu quay mặt, nhắm mắt theo hướng dẫn ngẫu nhiên, để ảnh chụp không dùng để "
                          "đăng nhập được.")
    r = doc.find_row(t, "Thời gian được miễn quét khuôn mặt sau lần xác thực gần nhất")
    doc.set_cell(t, r, 1, "Thời gian được miễn quét khuôn mặt sau lần xác thực gần nhất, dùng chung cho bản Web và ứng dụng.")
    r = max(i for i, row in enumerate(t.rows))
    doc.add_row(t, ["Thiết bị FaceID", "Máy tính, điện thoại cài ứng dụng đã đăng ký khóa riêng sau khi xác nhận qua Email."], like=r, after=r)

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
