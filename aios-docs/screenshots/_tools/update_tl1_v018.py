# -*- coding: utf-8 -*-
"""TL1 bản 0.18 (21/11/2026): trang Đăng nhập của ứng dụng máy trạm chọn Tổ chức thay cho khối AI Gateway (ai-workspace
PR #14); bộ phát hành Windows 6.11.21 trên GitHub Releases (PR #12, #13).
- Trang bìa; mục "Những điểm mới trong bản 0.18".
- 2.1, 2.4 (ứng dụng máy trạm), 2.8 (bộ phát hành, cài trên Windows, ảnh màn hình 3 chụp từ ứng dụng thật), 3.1, 12.1, 12.2.
Đầu vào: TL1 bản 0.17 ĐÃ điền mục lục. Đầu ra: bản 0.18 chưa điền mục lục ở <thư mục ra>/<tên>.src.docx, sau đó chép vào
_tools/toc_work/ và chạy build_toc.py TL1.
    python update_tl1_v018.py <TL1_v017.docx> <thư mục ra>"""
import os
import sys

from docx.table import Table

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docedit import Doc  # noqa: E402
from doc_versioning import new_section, set_caption, unfill_toc  # noqa: E402
from update_tl1_v016 import replace_picture_fit  # noqa: E402

T1 = "iThunder_AI_OS_TL1_Huong_dan_su_dung_AI_Workspace.docx"
DATE = "21/11/2026"
VER = "6.11.21"
SHOT = os.path.join(HERE, "..", "tl1_09_windows_login.jpg")
RELEASE = "https://github.com/hoangleminhvlu/ai-workspace/releases/latest"
ZIP = f"iThunder-AI-Workspace-{VER}-Windows.zip"
EXE = f"iThunder-AI-Workspace-{VER}-Windows-x64-setup.exe"

APP_LOGIN = [   # 2.4: thay 5 gạch đầu dòng về mục AI Gateway của ứng dụng
    "Lần đầu mở, ứng dụng tự vào AI Workspace của AI-TECH Center. Nếu bạn thuộc Tổ chức khác, chọn Tổ chức trong ô "
    "**Chọn Tổ chức** rồi bấm **Chuyển**.",
    "Tổ chức của bạn không có trong danh sách (ví dụ máy chủ chỉ chạy trong mạng nội bộ của Tổ chức): chọn mục cuối **Tổ "
    "chức khác: nhập địa chỉ do Quản trị viên cung cấp…**, nhập địa chỉ rồi bấm **Kết nối**. Ô nhập che ký tự; bấm biểu "
    "tượng con mắt để xem khi gõ.",
    "Tổ chức đã chọn được ghi nhớ sau khi bạn đăng nhập thành công; lần sau ứng dụng vào thẳng Tổ chức đó. Ứng dụng không "
    "hiển thị địa chỉ máy chủ của Tổ chức.",
    "Địa chỉ qua Internet phải dùng https để bảo vệ PassCode; ứng dụng từ chối địa chỉ http đi qua Internet. Máy chủ trong "
    "mạng nội bộ của Tổ chức: máy tính, điện thoại của bạn phải nối vào mạng đó.",
    "Nếu thấy **Chưa kết nối được AI Workspace của Tổ chức: AI Gateway không phản hồi sau 8 giây.**, hãy kiểm tra kết nối "
    "mạng rồi bấm **Thử lại**; ô nhập địa chỉ hiện ra để bạn chọn lại nếu cần.",
]


def tl1(src, dst):
    doc = Doc(src)
    unfill_toc(doc)
    t = doc.table("Phiên bản tài liệu")
    doc.set_cell(t, doc.find_row(t, "Phiên bản tài liệu"), 1,
                 f"0.18, bản dự thảo ngày {DATE} (cập nhật bản 0.17: chọn Tổ chức trên ứng dụng, bộ phát hành Windows {VER})")
    doc.set_cell(t, doc.find_row(t, "Phiên bản phần mềm"), 1,
                 f"AI Workspace {VER} (Web, ứng dụng Windows; Android 6.11.16) và AI Platform 6.11.18, nhánh main ngày {DATE}")
    new_section(doc, "Những điểm mới trong bản 0.17", "Những điểm mới trong bản 0.18",
                "So với bản 0.17, trang Đăng nhập của ứng dụng trên máy tính, điện thoại gọn hơn, và bản cài đặt cho Windows "
                "được phát hành trên Internet:",
                "So với bản 0.16, trang Giám sát Hệ thống",
                "Thống kê AI Tokens theo ngày, tháng",
                [("Ứng dụng tự vào AI Workspace của AI-TECH Center; đổi Tổ chức bằng ô Chọn Tổ chức, không còn mục AI Gateway", "2.4"),
                 (f"Tải bộ cài Windows {VER} kèm tài liệu này từ trang phát hành; góp ý qua Wiki, Email", "2.8")])

    # ---- 2.1
    p = doc.find("Nếu dùng ứng dụng, bạn cần địa chỉ AI Gateway của Tổ chức.", "p")
    doc.replace_in(p, "Nếu dùng ứng dụng, bạn cần địa chỉ AI Gateway của Tổ chức.",
                   "Ứng dụng tự vào AI Workspace của AI-TECH Center; Tổ chức khác chọn trong danh sách trên màn hình đăng nhập.")

    # ---- 2.4: đoạn dẫn và 5 gạch đầu dòng
    intro = doc.find("Khi dùng ứng dụng AI Workspace cho Windows hoặc Android (mục 2.8), màn hình đăng nhập", "p")
    doc.set_para(intro, "Khi dùng ứng dụng AI Workspace cho Windows hoặc Android (mục 2.8) [[ĐANG PHÁT TRIỂN]]:")
    el = intro.getnext()
    for text in APP_LOGIN:
        assert doc.text_of(el).strip(), "thiếu gạch đầu dòng"
        doc.set_para(el, text)
        el = el.getnext()

    # ---- 2.8
    p = doc.find("Ngoài trình duyệt web, bạn có thể dùng AI Workspace như một ứng dụng", "p")
    doc.set_para(p, "Ngoài trình duyệt web, bạn có thể dùng AI Workspace như một ứng dụng cài trên máy tính Windows 11 hoặc "
                    "điện thoại Android [[ĐANG PHÁT TRIỂN]]. Ứng dụng có cùng giao diện và chức năng như trên trình duyệt; "
                    f"điểm khác là bạn chọn Tổ chức ngay trên màn hình đăng nhập (mục 2.4). Bản thử nghiệm cho Windows {VER} "
                    f"tải được từ trang phát hành {RELEASE}; bản Android 6.11.16 do Quản trị viên của Tổ chức cung cấp. Ứng "
                    "dụng chưa có trên Microsoft Store, Google Play.")
    t = doc.table("Thiết bịYêu cầuTệp cài đặt")
    r = doc.find_row(t, "Windows 11 (64-bit)")
    doc.set_cell(t, r, 2, f"{EXE} [[ĐANG PHÁT TRIỂN]]")
    p = doc.find("Tổ chức gửi bạn tệp nén iThunder-AI-Workspace-6.11.16.zip", "p")
    doc.set_para(p, f"Tải tệp nén **{ZIP}** ở trang phát hành {RELEASE} (hoặc nhận từ Quản trị viên), giải nén, gồm:")
    items = [f"{EXE}: bộ cài cho máy tính Windows 11.",
             "iThunder_AI_OS_TL1_Huong_dan_su_dung_AI_Workspace_v<bản>.pdf: tài liệu này.",
             "HUONG_DAN_CAI_DAT.txt: hướng dẫn cài đặt ngắn gọn, giới hạn của bản dùng thử, kênh góp ý, đặt hàng phát triển.",
             "SHA256SUMS.txt: mã kiểm tra tính toàn vẹn của bộ cài và tài liệu."]
    el = p.getnext()
    for text in items:
        doc.set_para(el, text)
        el = el.getnext()
    p = doc.find("Khi nhận tệp qua mạng (Zalo, Email), nên kiểm tra tệp trước khi cài.", "p")
    doc.set_para(p, "Nên kiểm tra tệp trước khi cài. Trên Windows, mở PowerShell trong thư mục chứa tệp, chạy Get-FileHash với "
                    "tên tệp và tham số -Algorithm SHA256; kết quả phải trùng với dòng tương ứng trong SHA256SUMS.txt. Bản "
                    "Android 6.11.16 (tệp .apk) được ký bằng chứng thư của iThunder AI Workspace. Góp ý, báo lỗi, đặt hàng phát "
                    "triển: Wiki https://github.com/hoangleminhvlu/ai-workspace/wiki hoặc Email ai-tech.center@vlu.edu.vn.")
    p = doc.find("Nhận tệp iThunder-AI-Workspace-6.11.16-Windows-x64-setup.exe", "p")
    doc.step_after(p, p, 1, f"Tải và giải nén **{ZIP}** từ trang phát hành, hoặc nhận bộ cài từ Quản trị viên. Chỉ cài tệp từ "
                            "trang phát hành chính thức hoặc do Tổ chức cung cấp.")
    p.getparent().remove(p)
    p = doc.find("Ở màn hình đăng nhập, nhập hoặc chọn AI Gateway của Tổ chức, bấm Kết nối", "p")
    doc.step_after(p, p, 4, "Màn hình đăng nhập mở sẵn AI Workspace của AI-TECH Center. Chọn Tổ chức của bạn nếu khác (mục 2.4), "
                            "rồi đăng nhập bằng PassCode. Nút **Cỡ hiển thị** ở cuối thanh bên trái (hoặc Ctrl + / Ctrl -) phóng "
                            "to, thu nhỏ màn hình.")
    p.getparent().remove(p)
    cap = replace_picture_fit(doc, "Ảnh màn hình 3. ", SHOT)
    set_caption(doc, cap, f"Ứng dụng AI Workspace {VER} trên Windows 11 (bản thử nghiệm), lần đầu mở: AI Workspace của AI-TECH "
                          "Center và ô Chọn Tổ chức")
    p = doc.find("Mở ứng dụng iThunder AI Workspace, chọn AI Gateway và đăng nhập như trên Windows.", "p")
    doc.step_after(p, p, 4, "Mở ứng dụng **iThunder AI Workspace**, nhập hoặc chọn AI Gateway của Tổ chức (bản Android 6.11.16 "
                            "còn mục AI Gateway như Ảnh màn hình 4), rồi đăng nhập. Khi quét khuôn mặt hoặc gọi thoại, video, "
                            "ứng dụng hỏi quyền dùng camera, micro: hãy cho phép.")
    p.getparent().remove(p)
    cap4 = doc.find("Ảnh màn hình 4. ", "p")
    set_caption(doc, cap4, "Ứng dụng AI Workspace 6.11.16 trên Android (bản thử nghiệm): trang Đăng nhập với mục AI Gateway và "
                           "màn hình làm việc")

    # ---- 3.1, 12.1, 12.2
    t = doc.table("Rê chuột lên dòng tên để xem Tổ chức và AI Gateway đang dùng")
    r = doc.find_row(t, "Rê chuột lên dòng tên")
    doc.set_cell(t, r, 1, "Tên AI Workspace và logo. Rê chuột lên dòng tên để xem Tổ chức đang dùng.")
    t = doc.table("Ứng dụng trên máy tính, điện thoại có khác bản web không?")
    r = doc.find_row(t, "Ứng dụng trên máy tính, điện thoại có khác bản web không?")
    doc.set_cell(t, r, 1, "Cùng giao diện và chức năng. Ứng dụng mặc định vào AI Workspace của AI-TECH Center, bạn chọn Tổ chức "
                          "khác trên màn hình đăng nhập; Tổ chức được ghi nhớ sau lần đăng nhập thành công đầu tiên. Bản web dùng "
                          "địa chỉ AI Workspace của Tổ chức.")
    r = doc.find_row(t, "Không cài được bản mới đè lên bản cũ trên Android?")
    doc.set_cell(t, r, 1, "Bản đang cài là bản thử nghiệm nội bộ, ký bằng khóa khác. Gỡ bản cũ rồi cài bản mới, sau đó chọn lại "
                          "Tổ chức và đăng nhập.")
    t = doc.table("Mã PassCode không chính xác hoặc tài khoản chưa được phê duyệt")
    r = doc.find_row(t, "Chưa xác định được Tổ chức của AI Gateway này")
    doc.set_cell(t, r, 0, "Chưa kết nối được AI Workspace của Tổ chức: AI Gateway không phản hồi sau 8 giây.")
    doc.set_cell(t, r, 2, "Kiểm tra mạng, bấm Thử lại; hoặc chọn Tổ chức khác, nhập địa chỉ đúng do Quản trị viên cung cấp.")
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
