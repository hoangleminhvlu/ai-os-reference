# -*- coding: utf-8 -*-
"""TL2 bản 0.21 (23/11/2026): quy trình FaceID mới (iThunder-AI-OS PR #9, ai-workspace PR #16).
- Trang bìa; mục "Những điểm mới trong bản 0.21".
- 2.4, 4.5, 4.6, 6.6, 8.3 (ngưỡng thống nhất, khóa thiết bị, xác nhận qua Email), 10.1, 10.2, 10.4, 11.1, 11.3, A.1, A.3, B, C, G.
Đầu vào: TL2 bản 0.20 ĐÃ điền mục lục. Đầu ra: bản 0.21 chưa điền mục lục ở <thư mục ra>/<tên>.src.docx, sau đó chép vào
_tools/toc_work/ và chạy build_toc.py TL2.
    python update_v021.py <TL2_v020.docx> <thư mục ra>"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docedit import Doc  # noqa: E402
from doc_versioning import new_section, unfill_toc  # noqa: E402

T2 = "iThunder_AI_OS_TL2_Cai_dat_va_Van_hanh_AI_Platform.docx"
DATE = "23/11/2026"
VER = "6.11.23"


def step(doc, needle, n, markup):
    p = doc.find(needle, "p")
    new = doc.step_after(p, p, n, markup)
    p.getparent().remove(p)
    return new


def tl2(src, dst):
    doc = Doc(src)
    unfill_toc(doc)
    t = doc.table("Phiên bản tài liệu")
    doc.set_cell(t, doc.find_row(t, "Phiên bản tài liệu"), 1,
                 f"0.21, bản dự thảo ngày {DATE} (cập nhật bản 0.20: quy trình FaceID mới, khóa thiết bị, xác nhận qua Email)")
    doc.set_cell(t, doc.find_row(t, "Phiên bản phần mềm"), 1,
                 f"AI Workspace và AI Platform 6.11.18 (nhánh main ngày {DATE}, sau khi hợp nhất PR #5 đến #9 của iThunder-AI-OS và "
                 f"PR #6 đến #16 của ai-workspace); ứng dụng 6.11.21 trên trang phát hành, {VER} là bản thử")
    new_section(doc, "Những điểm mới trong bản 0.20", "Những điểm mới trong bản 0.21",
                "Bổ sung quy trình FaceID mới (iThunder-AI-OS PR #9, ai-workspace PR #16). Các thay đổi và nơi mô tả:",
                "Bổ sung bộ phát hành dùng thử Windows và Android",
                "Ứng dụng máy trạm: mặc định AI-TECH Center, ô Chọn Tổ chức",
                [("Ngưỡng độ khớp thống nhất 0,6 (FACEID_ACCURACY); bỏ ngưỡng riêng từng người dùng", "2.4, 8.3, C"),
                 ("Web bắt buộc kiểm tra người thật; ứng dụng có khóa thiết bị thì không", "2.4, 8.3"),
                 ("Khóa thiết bị, xác nhận thiết bị mới bằng mã gửi qua Email, thu hồi thiết bị", "8.3, 6.6, A.1, A.3, B"),
                 ("Triển khai, bảo mật, xử lý sự cố, kết quả kiểm thử", "4.5, 4.6, 10, 11")])

    # ---- 2.4
    step(doc, "Nếu lần FaceID gần nhất của Tổ chức (pager) còn trong PassTime", 3,
         "Nếu lần FaceID gần nhất của Tổ chức (pager) còn trong PassTime (employeeType), Backend cấp JWT đầy đủ, dù lần quét đó trên "
         "Web hay ứng dụng. Mặc định PassTime là 18.000.000 ms, tức 5 giờ.")
    step(doc, "Nếu không, Backend trả REQUIRE_FACEID. Workspace mở WebSocket", 4,
         "Nếu không, Backend trả REQUIRE_FACEID. Workspace mở WebSocket /ws/auth; ứng dụng máy trạm chứng minh khóa thiết bị trước "
         "(device_challenge, device_auth). Sau đó Workspace gửi khung hình vuông 320×320 khoảng 3 khung mỗi giây (chỉ khi khung trước "
         "đã có kết quả) qua hành động verify_faceid, kèm liveness=1 và client.")
    step(doc, "Backend kiểm tra chất lượng khung, so khớp khung nhìn thẳng với mọi mẫu", 5,
         "Backend kiểm tra chất lượng khung, so khớp khung nhìn thẳng với mọi mẫu (ngưỡng thống nhất FACEID_ACCURACY 0,6). Web: phát hai "
         "thử thách ngẫu nhiên để kiểm tra người thật, rồi so khớp lại khung nhìn thẳng cuối. Ứng dụng có khóa thiết bị đã đăng ký: "
         "một khung khớp là đủ; thiết bị chưa đăng ký: gửi mã xác nhận qua Email (mục 8.3). Tài khoản chưa có mẫu: ảnh vừa quét được lưu "
         "làm mẫu.")

    # ---- 4.5, 4.6
    t = doc.table("Cài bộ cài Windows hoặc APK 6.11.21 từ trang phát hành")
    r = doc.find_row(t, "Cài bộ cài Windows hoặc APK 6.11.21")
    doc.set_cell(t, r, 2, f"Trang Đăng nhập hiện đúng Tổ chức; đăng nhập được. Với ứng dụng từ bản {VER}: thiết bị mới nhận được mã qua "
                          "Email và đăng ký được")
    box = doc.find("Đã kiểm tra ngày 21/11/2026 với bộ phát hành 6.11.21", "tbl")
    doc.replace_in(box, "Chưa kiểm tra đăng nhập bằng PassCode, FaceID có kiểm tra người thật, gọi thoại / video trên ứng dụng và trên "
                        "điện thoại thật; chưa build iOS.",
                   f"Ngày 23/11/2026, bản thử {VER} (chưa đăng trang phát hành): ứng dụng Windows 11 và APK trên máy ảo Android 16 (cài đè "
                   "6.11.21, versionCode 61123) đăng ký thiết bị mới bằng mã Email và đăng nhập được, với AI Platform dùng LDAP, Email, "
                   "khuôn mặt giả lập. Chưa kiểm tra camera và Email thật, gọi thoại / video trên ứng dụng, điện thoại thật; chưa build iOS.")

    # ---- 6.5 step-up, 6.6 nhật ký
    p = doc.find("Các thao tác nhạy cảm trả HTTP 428 STEP_UP_REQUIRED", "p")
    doc.replace_in(p, "xóa ảnh FaceID (cả khi người dùng tự xóa)", "xóa ảnh FaceID và thu hồi, khôi phục thiết bị FaceID (cả khi người dùng tự làm)")
    p = doc.find("Xem ảnh FaceID của người khác ghi user.faceid.photos.view", "p")
    doc.para_after(p, p, "Thiết bị FaceID ghi self.faceid.device.request (gửi mã), self.faceid.device.register (đăng ký, cả khi mã hết hạn "
                         "hoặc sai quá số lần), self.faceid.device.revoke / restore và user.faceid.device.revoke / restore. Mã xác nhận không "
                         "bao giờ nằm trong nhật ký.")

    # ---- 8.3
    p = doc.find("Từ ngày 21/11/2026 (iThunder-AI-OS PR #5, ai-workspace PR #6), mỗi lượt quét FaceID", "p")
    doc.set_para(p, "[[ĐÃ CÓ]] Mỗi lượt quét FaceID là một phiên gắn với kết nối /ws/auth và PassCode (iThunder-AI-OS PR #5, #9; "
                    "ai-workspace PR #6, #16). AI Platform tự phân loại máy trạm: chỉ coi là ứng dụng khi Origin của vỏ ứng dụng "
                    "(tauri.localhost, capacitor://localhost, http://localhost không cổng) khớp với trường client (desktop / mobile); mọi "
                    "trường hợp khác là Web. Các bước:")
    p = doc.find("So khớp: chỉ mã hóa khung nhìn thẳng", "p")
    doc.set_para(p, "**So khớp**: chỉ mã hóa khung nhìn thẳng (FACEID_NUM_JITTERS lần lấy mẫu), so với mọi mẫu của người dùng, lấy khoảng cách "
                    "nhỏ nhất. Ngưỡng độ khớp thống nhất FACEID_ACCURACY = 0,6 cho mọi người dùng, Web và ứng dụng; iDAccountPrimaryType "
                    "không còn dùng làm ngưỡng. Người dùng bản cũ chỉ có ảnh: ảnh được mã hóa một lần và ghi vào initials. Sau lần đăng "
                    "nhập có kiểm tra người thật, khớp rõ ràng mà đủ khác các mẫu sẵn có, mẫu mới (kèm ảnh) được bổ sung, tối đa "
                    "FACEID_MAX_TEMPLATES.")
    p = doc.find("Kiểm tra người thật: khung nhìn thẳng đầu tiên khớp mẫu", "p")
    doc.set_para(p, "**Kiểm tra người thật**: bắt buộc trên Web và trên ứng dụng không chứng minh được khóa thiết bị (bản 6.11.21 trở về trước, "
                    "hoặc khi không tra được sổ đăng ký khóa). Khung nhìn thẳng đầu tiên khớp mẫu, rồi làm FACEID_LIVENESS_STEPS thử thách "
                    "ngẫu nhiên (quay trái, quay phải, nhắm mắt rồi mở; status LIVENESS kèm hướng dẫn), cuối cùng nhìn thẳng và khớp lại, "
                    "cùng người với khung đầu. AI Platform phát và tự kiểm tra thử thách; máy trạm chỉ hiển thị.")
    p = doc.find("Đăng ký lần đầu cũng phải qua liveness; lưu ảnh", "p")
    doc.set_para(p, "**Đăng ký khuôn mặt lần đầu**: trên Web cũng phải qua kiểm tra người thật, lưu ảnh và mẫu của khung đầu, khung cuối; trên ứng "
                    "dụng lưu một khung nhìn thẳng sau khi thiết bị đã xác nhận qua Email. Mỗi FACEID_MAX_MISMATCHES (10) khung không khớp "
                    "trên một kết nối tính một lần sai vào giới hạn đăng nhập theo IP. Nhật ký XML ghi Liveness, Duration, Client, Device.")
    b = doc.para_after(p, p, "**Khóa thiết bị** (identity/devices.py, bảng iAIP_faceid_devices): ứng dụng sinh cặp khóa ECDSA P-256 cho từng AI "
                             "Gateway, khóa riêng không xuất ra được (WebCrypto, IndexedDB). Trên /ws/auth, AI Platform phát số ngẫu nhiên dùng "
                             "một lần (60 giây), ứng dụng ký kèm Mã Tổ chức. Khóa đã đăng ký và còn hiệu lực: quét một khung khớp là vào. Khóa "
                             "đã thu hồi: từ chối (DEVICE_REVOKED).")
    b = doc.para_after(b, p, "**Xác nhận thiết bị mới qua Email**: khuôn mặt khớp trên thiết bị chưa đăng ký thì chưa đăng nhập; AI Platform gửi mã 8 "
                             "số tới mọi Email đăng ký và trả DEVICE_CONFIRM (Email che bớt). Người dùng nhập mã trên chính thiết bị "
                             "(confirm_device), khóa được đăng ký rồi ứng dụng quét lại để vào. Mã hiệu lực 10 phút, 5 lần nhập, dùng một lần, gắn "
                             "người dùng và thiết bị, giữ dạng băm trong bộ nhớ (khởi động lại thì quét lại để nhận mã mới); gửi lại sau 60 giây, "
                             "giới hạn theo EMAIL_REQUEST_MAX_PER_ADDRESS cho mỗi người dùng. Mỗi lần nhập sai tính vào giới hạn đăng nhập theo "
                             "IP. Dùng mã thay cho liên kết để bộ quét thư không tự xác nhận.")
    doc.para_after(b, p, "**Quản lý thiết bị**: người dùng (Hồ sơ của tôi) và Quản trị viên (identity.security.reset) xem, thu hồi, khôi phục, cần "
                         "step-up. Thu hồi không ngắt phiên đang mở và PassTime còn hiệu lực: khi mất thiết bị, cấp lại cả PassCode. Thu hồi "
                         "tài khoản thì xóa các khóa thiết bị của tài khoản.")
    t = doc.table("required (mặc định): máy trạm không gửi liveness=1")
    r = doc.find_row(t, "required (mặc định): máy trạm không gửi liveness=1")
    doc.set_cell(t, r, 1, "Áp dụng cho Web và ứng dụng chưa có khóa thiết bị. required (mặc định): máy trạm không gửi liveness=1 nhận "
                          "CLIENT_UPDATE_REQUIRED; optional: chỉ với máy trạm mới (giai đoạn chuyển tiếp, máy trạm cũ chưa chống được ảnh "
                          "tĩnh); off: tắt")
    r = doc.find_row(t, "FACEID_DEFAULT_ACCURACY, FACEID_MIN_ACCURACY")
    doc.set_cell(t, r, 0, "FACEID_ACCURACY")
    doc.set_cell(t, r, 1, "Ngưỡng độ khớp thống nhất (0,6). Thay FACEID_DEFAULT_ACCURACY, FACEID_MIN_ACCURACY: xóa hai biến cũ khỏi .env")
    doc.add_row(t, ["FACEID_APP_LIVENESS", "off (mặc định): ứng dụng có khóa thiết bị không cần kiểm tra người thật; required: kiểm tra như Web"],
                like=r, after=r)
    box = doc.find("Liveness chủ động chống ảnh tĩnh 2D; chưa chống được mặt nạ 3D", "tbl")
    doc.set_note(box, "Liveness chủ động chống ảnh tĩnh 2D; chưa chống được mặt nạ 3D, video giả mạo tạo theo thời gian thực hay camera ảo. Ngưỡng "
                      "thị sai mới được kiểm bằng mô phỏng hình học 3D, ngưỡng 0,6 chưa đo trên nhiều người dùng thật: nếu người dùng thật khó "
                      "qua bước quay đầu, giảm FACEID_PARALLAX_DELTA hoặc FACEID_YAW_DELTA; nếu bị từ chối nhầm, xóa ảnh mẫu kém (mục 6.7) trước "
                      "khi nghĩ tới hạ ngưỡng. Ứng dụng 6.11.21 vẫn dùng được nhưng phải kiểm tra người thật.")
    # Biện pháp vận hành cho ứng dụng không liveness
    p = doc.find("Ảnh mẫu của từng người dùng xem và xóa được ở mục 6.7", "p")
    doc.para_after(p, p, "**Vận hành**: trên ứng dụng, bảo vệ dựa vào PassCode, khớp khuôn mặt, khóa thiết bị và hộp thư của người dùng. Vì vậy mọi tài "
                         "khoản cần có Email đăng ký; hộp thư nên bật xác thực hai lớp; đặt PassTime ngắn cho vai trò quản trị, vì trong thời gian "
                         "này đăng nhập không qua FaceID; Tổ chức cần mức cao hơn thì đặt FACEID_APP_LIVENESS=required.")

    # ---- 10.1, 10.2, 10.4
    p = doc.find("FaceID có kiểm tra người thật do AI Platform phát thử thách ngẫu nhiên", "p")
    doc.set_para(p, "FaceID trên Web có kiểm tra người thật do AI Platform phát thử thách ngẫu nhiên và tự kiểm tra; ứng dụng máy trạm dùng khóa "
                    "thiết bị ký số, thiết bị mới xác nhận bằng mã gửi qua Email, thu hồi được; kiểm tra chất lượng khung, chống tráo ảnh giữa "
                    "chừng; khung không khớp lặp lại và mã nhập sai được tính vào giới hạn đăng nhập theo IP (mục 8.3).")
    t = doc.table("Liveness chủ động (mục 8.3) chống ảnh tĩnh 2D")
    r = doc.find_row(t, "Liveness chủ động (mục 8.3) chống ảnh tĩnh 2D")
    doc.set_cell(t, r, 1, "Liveness chủ động (mục 8.3) trên Web chống ảnh tĩnh 2D; chưa chống mặt nạ 3D, video giả mạo thời gian thực, camera ảo. "
                          "Ứng dụng máy trạm dựa vào khóa thiết bị và xác nhận qua Email. Ngưỡng chưa kiểm trên nhiều người dùng thật")
    doc.set_cell(t, r, 2, "Thử với người dùng thật, hiệu chỉnh ngưỡng; thêm mô hình chống giả mạo thụ động (onnxruntime đã có) cho cả ứng dụng; kênh "
                          "xác nhận tách khỏi Email nhận PassCode; đăng ký khuôn mặt có xác minh. Ảnh mẫu xem, xóa được (mục 6.7)")
    t = doc.table("Lộ JWT hoặc mã truy cập")
    r = doc.find_row(t, "Lộ PassCode")
    doc.set_cell(t, r, 1, "Cấp lại PassCode; nếu nghi có đăng ký khuôn mặt sai thì xóa FaceID; thu hồi thiết bị FaceID lạ")
    doc.add_row(t, ["Mất thiết bị hoặc nghi lộ hộp thư của người dùng", "Thu hồi thiết bị FaceID, cấp lại PassCode (sau khi hộp thư đã an toàn); rà soát "
                    "self.faceid.device.* trong nhật ký"], like=r, after=r)

    # ---- 11.1, 11.3
    t = doc.table("Kết quả ngày 20/11/2026")
    r = doc.find_row(t, "tests/identity/run_all.py")
    doc.set_cell(t, r, 1, "Đạt (20 tệp, ngày 23/11/2026, gồm test_14_rag_chat.py so sánh trực tiếp lời nhắc của AI Chat/Work và Trao đổi trực tuyến; "
                          "test_16_faceid_liveness.py: hình học người thật và ảnh phẳng, ngưỡng 0,6, khóa thiết bị, mã xác nhận Email, PassTime dùng "
                          "chung; test_20_faceid_devices.py: API thiết bị)")
    r = doc.find_row(t, "tests/ai_usage/test_ai_usage_db.py")
    doc.add_row(t, ["tests/faceid_devices/test_devices_db.py", "Đạt (23/11/2026)", "PostgreSQL thật, mã Tổ chức tạm, dọn sạch: đăng ký, thu hồi, khôi phục, "
                    "tách theo Tổ chức"], like=r, after=r)
    t = doc.table("Người dùng thật không qua bước quay mặt")
    r = doc.find_row(t, "Máy trạm báo AI Gateway yêu cầu kiểm tra người thật")
    doc.set_cell(t, r, 2, "Triển khai bản Web mới (no-store cho index.html); người dùng cài ứng dụng mới từ trang phát hành; tạm thời "
                          "FACEID_LIVENESS=optional", code=False)
    for x in [["Ứng dụng mới vẫn đòi kiểm tra người thật", "nginx không chuyển tiếp Origin, CLIENT_APP_ORIGINS tắt, không tra được iAIP_faceid_devices, hoặc "
               "FACEID_APP_LIVENESS=required", "Kiểm tra proxy_set_header và nhật ký [FaceID Device]; quyền tạo bảng của tài khoản CSDL"],
              ["Không nhận được mã xác nhận thiết bị", "Tài khoản chưa có Email, SMTP lỗi, thư vào Spam, quá giới hạn gửi", "Kiểm tra Email trong hồ sơ, "
               "nhật ký [FaceID Device] và [iAIP Send Email]; chờ LOGIN_LOCK_MINUTES"],
              ["Nhiều người dùng bị từ chối khuôn mặt sau khi nâng cấp", "Ngưỡng cũ 0,55 nay là 0,6; mẫu cũ kém chất lượng", "Xóa ảnh kém (mục 6.7) hoặc Xóa "
               "FaceID để đăng ký lại; chỉ hạ FACEID_ACCURACY khi đã đo"]]:
        r = doc.add_row(t, x, like=r, after=r)

    # ---- A.1, A.3, B, C
    t = doc.table("Xác thực FaceID theo khung ảnh (verify_faceid, liveness=1)")
    r = doc.find_row(t, "Xác thực FaceID theo khung ảnh (verify_faceid, liveness=1)")
    doc.set_cell(t, r, 1, "Xác thực FaceID theo khung ảnh (verify_faceid, liveness=1, client). Phản hồi: HINT (khung chưa đạt), LIVENESS (bước kế "
                          "tiếp), FAIL (FACE_MISMATCH, LIVENESS_TIMEOUT, FACE_SWITCHED), SUCCESS; máy trạm cũ nhận ERROR CLIENT_UPDATE_REQUIRED. Khóa "
                          "thiết bị: device_challenge → DEVICE_CHALLENGE; device_auth → DEVICE_OK / DEVICE_REJECTED. Thiết bị mới: DEVICE_CONFIRM; "
                          "confirm_device → DEVICE_CONFIRMED, FAIL (DEVICE_CODE_WRONG, DEVICE_CODE_EXPIRED); resend_device_code; thiết bị bị thu hồi: "
                          "ERROR DEVICE_REVOKED")
    t = doc.table("GET /users/{uid}/faceid/photos; POST")
    r = doc.find_row(t, "GET /users/{uid}/faceid/photos; POST")
    r = doc.add_row(t, ["GET /me/faceid/devices; POST /me/faceid/devices/{id}", "Đăng nhập", "Khi thu hồi, khôi phục"], like=r, after=r)
    doc.add_row(t, ["GET /users/{uid}/faceid/devices; POST /users/{uid}/faceid/devices/{id}", "security.reset", "Khi thu hồi, khôi phục"], like=r, after=r)
    t = doc.table("iAIP_chat_threads, iAIP_chat_messages")
    r = doc.find_row(t, "iAIP_ai_usage")
    doc.add_row(t, ["iAIP_faceid_devices", "Khóa công khai của thiết bị FaceID theo Tổ chức, người dùng; thời điểm đăng ký, dùng gần nhất, thu hồi",
                    "mã nguồn"], like=r, after=r)
    t = doc.table("FACEID_LIVENESS, FACEID_LIVENESS_STEPS, FACEID_LIVENESS_TIMEOUT, FACEID_DEFAULT_ACCURACY")
    r = doc.find_row(t, "FACEID_LIVENESS, FACEID_LIVENESS_STEPS")
    doc.set_cell(t, r, 1, "FACEID_LIVENESS, FACEID_LIVENESS_STEPS, FACEID_LIVENESS_TIMEOUT, FACEID_ACCURACY, FACEID_APP_LIVENESS, FACEID_NUM_JITTERS, "
                          "FACEID_MAX_TEMPLATES, FACEID_YAW_DELTA, FACEID_PARALLAX_DELTA (mục 8.3)")
    t = doc.table("Khung 320×320, khoảng 3 khung mỗi giây")
    r = doc.find_row(t, "Khung 320×320, khoảng 3 khung mỗi giây")
    doc.set_cell(t, r, 1, "Khung 320×320, khoảng 3 khung mỗi giây; Web: 2 thử thách, 40 giây mỗi lượt (máy trạm dừng sau 90 giây); tối đa 5 mẫu mỗi "
                          "người dùng; 10 khung không khớp tính một lần sai; mã xác nhận thiết bị 8 số, 10 phút, 5 lần nhập, gửi lại sau 60 giây")

    # ---- Phụ lục G
    t = doc.table("Căn cứ chính của nội dung hiện trạng")
    r = doc.find_row(t, "Căn cứ chính của nội dung hiện trạng")
    doc.set_cell(t, r, 0, f"Mã nguồn ai-platform (AI Platform 6.11.18), ai-work (AI Workspace, ứng dụng 6.11.21 và bản thử {VER}), nhánh main ngày {DATE} sau "
                          "khi hợp nhất FaceID liveness (iThunder-AI-OS PR #5, ai-workspace PR #6), Kết nối Kho Đám mây (PR #6, PR #7), Ảnh FaceID trong "
                          "hồ sơ (PR #7, PR #8), ẩn địa chỉ AI Gateway và phóng to (ai-workspace PR #9), AI Workspace 6.11.21 (PR #10), thống kê AI "
                          "Tokens (PR #8, PR #11), bộ phát hành Windows và trang Đăng nhập chọn Tổ chức (ai-workspace PR #12 đến #14), bộ phát hành "
                          "Android (PR #15), khóa thiết bị FaceID và xác nhận qua Email (iThunder-AI-OS PR #9, ai-workspace PR #16)", code=False)
    doc.save(dst)
    return doc


if __name__ == "__main__":
    src, out = sys.argv[1:3]
    os.makedirs(out, exist_ok=True)
    d = tl2(src, os.path.join(out, T2[:-5] + ".src.docx"))
    print("tl2", len(d.changes), "thay đổi")
