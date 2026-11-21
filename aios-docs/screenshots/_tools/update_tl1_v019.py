# -*- coding: utf-8 -*-
"""TL1 bản 0.19 (21/11/2026): bản Android 6.11.21 trên GitHub Releases (ai-workspace PR #15).
- Trang bìa; mục "Những điểm mới trong bản 0.19".
- 2.8: đoạn mở đầu, bảng thiết bị, bộ phát hành (hai tệp ZIP Windows, Android), kiểm tra tệp, cài trên Android (bước 1, 4),
  Ảnh màn hình 4 (trang Đăng nhập của ứng dụng Android 6.11.21 thật, ghép với màn hình làm việc cũ), cập nhật.
Đầu vào: TL1 bản 0.18 ĐÃ điền mục lục. Đầu ra: bản 0.19 chưa điền mục lục ở <thư mục ra>/<tên>.src.docx, sau đó chép vào
_tools/toc_work/ và chạy build_toc.py TL1.
    python update_tl1_v019.py <TL1_v018.docx> <thư mục ra>"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docedit import Doc  # noqa: E402
from doc_versioning import new_section, set_caption, unfill_toc  # noqa: E402
from update_tl1_v016 import replace_picture_fit  # noqa: E402

T1 = "iThunder_AI_OS_TL1_Huong_dan_su_dung_AI_Workspace.docx"
DATE = "21/11/2026"
VER = "6.11.21"
SHOT = os.path.join(HERE, "..", "tl1_10_android_login.png")
RELEASE = "https://github.com/hoangleminhvlu/ai-workspace/releases/latest"
ZIP_W = f"iThunder-AI-Workspace-{VER}-Windows.zip"
ZIP_A = f"iThunder-AI-Workspace-{VER}-Android.zip"
EXE = f"iThunder-AI-Workspace-{VER}-Windows-x64-setup.exe"
APK = f"iThunder-AI-Workspace-{VER}-Android.apk"
CERT = "394558b8f8f5026123e861d7e4b5b16fac1240131e1833720c192055d8700328"


def tl1(src, dst):
    doc = Doc(src)
    unfill_toc(doc)
    t = doc.table("Phiên bản tài liệu")
    doc.set_cell(t, doc.find_row(t, "Phiên bản tài liệu"), 1,
                 f"0.19, bản dự thảo ngày {DATE} (cập nhật bản 0.18: bộ phát hành Android {VER})")
    doc.set_cell(t, doc.find_row(t, "Phiên bản phần mềm"), 1,
                 f"AI Workspace {VER} (Web, ứng dụng Windows và Android) và AI Platform 6.11.18, nhánh main ngày {DATE}")
    new_section(doc, "Những điểm mới trong bản 0.18", "Những điểm mới trong bản 0.19",
                f"So với bản 0.18, ứng dụng Android {VER} được phát hành cùng bản Windows trên Internet:",
                "So với bản 0.17, trang Đăng nhập",
                "Ứng dụng tự vào AI Workspace của AI-TECH Center",
                [(f"Tải ứng dụng Android {VER} kèm tài liệu này từ trang phát hành; cài đè lên bản 6.11.16", "2.8"),
                 ("Trang Đăng nhập trên Android chọn Tổ chức như trên Windows (Ảnh màn hình 4)", "2.8")])

    # ---- 2.8: đoạn mở đầu, bảng thiết bị
    p = doc.find("Ngoài trình duyệt web, bạn có thể dùng AI Workspace như một ứng dụng", "p")
    doc.set_para(p, "Ngoài trình duyệt web, bạn có thể dùng AI Workspace như một ứng dụng cài trên máy tính Windows 11 hoặc "
                    "điện thoại Android [[ĐANG PHÁT TRIỂN]]. Ứng dụng có cùng giao diện và chức năng như trên trình duyệt; "
                    "điểm khác là bạn chọn Tổ chức ngay trên màn hình đăng nhập (mục 2.4). Bản thử nghiệm "
                    f"{VER} cho Windows và Android tải được từ trang phát hành {RELEASE}. Ứng dụng chưa có trên Microsoft "
                    "Store, Google Play.")
    t = doc.table("Thiết bịYêu cầuTệp cài đặt")
    r = doc.find_row(t, "Điện thoại, máy tính bảng Android")
    doc.set_cell(t, r, 2, f"{APK} [[ĐANG PHÁT TRIỂN]]")

    # ---- Bộ phát hành dùng thử
    p = doc.find(f"Tải tệp nén {ZIP_W} ở trang phát hành", "p")
    doc.set_para(p, f"Tải ở trang phát hành {RELEASE} (hoặc nhận từ Quản trị viên) tệp nén cho thiết bị của bạn: **{ZIP_W}** "
                    f"(Windows) hoặc **{ZIP_A}** (Android), giải nén, gồm:")
    items = [f"{EXE} (Windows) hoặc {APK} (Android): tệp cài đặt.",
             "iThunder_AI_OS_TL1_Huong_dan_su_dung_AI_Workspace_v<bản>.pdf: tài liệu này.",
             "HUONG_DAN_CAI_DAT.txt: hướng dẫn cài đặt ngắn gọn, giới hạn của bản dùng thử, kênh góp ý, đặt hàng phát triển.",
             "SHA256SUMS.txt: mã kiểm tra tính toàn vẹn của tệp cài đặt và tài liệu."]
    el = p.getnext()
    for text in items:
        doc.set_para(el, text)
        el = el.getnext()
    p = doc.find("Nên kiểm tra tệp trước khi cài. Trên Windows, mở PowerShell", "p")
    doc.set_para(p, "Nên kiểm tra tệp trước khi cài. Trên Windows, mở PowerShell trong thư mục chứa tệp, chạy Get-FileHash với "
                    "tên tệp và tham số -Algorithm SHA256; kết quả phải trùng với dòng tương ứng trong SHA256SUMS.txt. Tệp "
                    "Android (.apk) được ký bằng chứng thư của iThunder AI Workspace, dấu vân tay SHA-256 ghi trong "
                    f"HUONG_DAN_CAI_DAT.txt ({CERT[:8]}…). Góp ý, báo lỗi, đặt hàng phát triển: Wiki "
                    "https://github.com/hoangleminhvlu/ai-workspace/wiki hoặc Email ai-tech.center@vlu.edu.vn.")

    # ---- Cài trên Android: bước 1, bước 3, bước 4, ảnh màn hình 4
    p = doc.find("Nhận tệp cài đặt iThunder-AI-Workspace-6.11.16-Android.apk từ Quản trị viên", "p")
    doc.step_after(p, p, 1, f"Tải và giải nén **{ZIP_A}** từ trang phát hành (hoặc nhận tệp {APK} từ Quản trị viên), "
                            "rồi mở tệp .apk trên điện thoại.")
    p.getparent().remove(p)
    p = doc.find("Nếu Google Play Protect cảnh báo ứng dụng chưa được xác minh", "p")
    doc.step_after(p, p, 3, "Nếu Google Play Protect cảnh báo ứng dụng chưa được xác minh, chọn **Thông tin chi tiết** rồi "
                            "**Vẫn cài đặt**. Chỉ làm vậy với tệp từ trang phát hành chính thức hoặc do Tổ chức cung cấp. Đã cài "
                            "bản 6.11.16: cài đè lên, không cần gỡ.")
    p.getparent().remove(p)
    p = doc.find("Mở ứng dụng iThunder AI Workspace, nhập hoặc chọn AI Gateway của Tổ chức", "p")
    doc.step_after(p, p, 4, "Mở ứng dụng **iThunder AI Workspace**. Màn hình đăng nhập mở sẵn AI Workspace của AI-TECH Center "
                            "(Ảnh màn hình 4); chọn Tổ chức của bạn nếu khác (mục 2.4), rồi đăng nhập bằng PassCode. Khi quét "
                            "khuôn mặt hoặc gọi thoại, video, ứng dụng hỏi quyền dùng camera, micro: hãy cho phép.")
    p.getparent().remove(p)
    cap = replace_picture_fit(doc, "Ảnh màn hình 4. ", SHOT)
    set_caption(doc, cap, f"Ứng dụng AI Workspace {VER} trên Android (bản thử nghiệm): trang Đăng nhập lần đầu mở, với ô "
                          "Chọn Tổ chức, và màn hình làm việc")

    # ---- Cập nhật
    p = doc.find("Cập nhật: khi Tổ chức phát hành bản mới", "p")
    doc.set_para(p, f"Cập nhật: tải bản mới ở trang phát hành {RELEASE} (hoặc nhận từ Quản trị viên) và cài đè lên bản đang "
                    "dùng; Tổ chức đã chọn được giữ nguyên.")
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
