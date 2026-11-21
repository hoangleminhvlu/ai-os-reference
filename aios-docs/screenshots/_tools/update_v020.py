# -*- coding: utf-8 -*-
"""TL2 bản 0.20 (22/11/2026): bộ phát hành Windows, Android 6.11.21 trên GitHub Releases (ai-workspace PR #12 đến #15) và
trang Đăng nhập của ứng dụng máy trạm chọn Tổ chức, không hiện địa chỉ AI Gateway (PR #9, #14).
- Trang bìa; mục "Những điểm mới trong bản 0.20".
- 2.4: ứng dụng máy trạm, Ảnh màn hình 1 (ứng dụng Windows 6.11.21 thật).
- 4.5; 4.6: phiên bản, lệnh, thông số, bộ phát hành (đầy đủ, --windows, --android), GitHub Releases, kết quả kiểm tra.
- 8.3 hộp Lưu ý; 10.1; 11.3; Phụ lục G.
Đầu vào: TL2 bản 0.19 ĐÃ điền mục lục. Đầu ra: bản 0.20 chưa điền mục lục ở <thư mục ra>/<tên>.src.docx, sau đó chép vào
_tools/toc_work/ và chạy build_toc.py TL2.
    python update_v020.py <TL2_v019.docx> <thư mục ra>"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docedit import Doc  # noqa: E402
from doc_versioning import new_section, set_caption, unfill_toc  # noqa: E402
from update_tl1_v016 import replace_picture_fit  # noqa: E402

T2 = "iThunder_AI_OS_TL2_Cai_dat_va_Van_hanh_AI_Platform.docx"
DATE = "22/11/2026"
VER = "6.11.21"
SHOT = os.path.join(HERE, "..", "tl1_09_windows_login.jpg")
RELEASES = "https://github.com/hoangleminhvlu/ai-workspace/releases"


def tl2(src, dst):
    doc = Doc(src)
    unfill_toc(doc)
    t = doc.table("Phiên bản tài liệu")
    doc.set_cell(t, doc.find_row(t, "Phiên bản tài liệu"), 1,
                 f"0.20, bản dự thảo ngày {DATE} (cập nhật bản 0.19: bộ phát hành Windows, Android {VER}, trang Đăng nhập "
                 "của ứng dụng máy trạm)")
    doc.set_cell(t, doc.find_row(t, "Phiên bản phần mềm"), 1,
                 f"AI Workspace {VER} và AI Platform 6.11.18 (nhánh main ngày {DATE}, sau khi hợp nhất PR #5 đến #8 của "
                 "iThunder-AI-OS và PR #6 đến #15 của ai-workspace)")
    new_section(doc, "Những điểm mới trong bản 0.19", "Những điểm mới trong bản 0.20",
                f"Bổ sung bộ phát hành dùng thử Windows và Android {VER} trên GitHub Releases, trang Đăng nhập của ứng dụng máy "
                "trạm chọn Tổ chức (ai-workspace PR #9, #12 đến #15). Các thay đổi và nơi mô tả:",
                "Bổ sung thống kê AI Tokens theo ngày, tháng",
                "Sổ iAIP_ai_usage, thống kê, phạm vi admin / manager",
                [("Ứng dụng máy trạm: mặc định AI-TECH Center, ô Chọn Tổ chức, không hiện địa chỉ AI Gateway", "2.4"),
                 ("release.sh --windows / --android: gói ZIP kèm TL1; phát hành trên GitHub Releases; kết quả kiểm tra", "4.6"),
                 ("Xử lý sự cố khi cài, cập nhật ứng dụng", "4.5, 8.3, 11.3")])

    # ---- 2.4
    p = doc.find("Ứng dụng máy trạm không có tên miền của Tổ chức", "p")
    doc.set_para(p, "Ứng dụng máy trạm không có tên miền của Tổ chức: AI Workspace mở tại http://localhost khi phát triển, ứng "
                    "dụng Windows có nguồn http://tauri.localhost, Android http://localhost, iOS capacitor://localhost; "
                    "CLIENT_APP_ORIGINS cho phép các nguồn này gọi API. Lần đầu mở, ứng dụng kết nối DEFAULT_GATEWAY_URL (AI "
                    "Workspace của AI-TECH Center, chờ tối đa 8 giây). Trang Đăng nhập chỉ hiện khung Tổ chức (không phân biệt "
                    "On-Premise / On-Clouds) với ô **Chọn Tổ chức** gộp: danh mục Tổ chức của AI Gateway đang kết nối "
                    "(ORG_DIRECTORY_FILE), các AI Gateway đã đăng nhập thành công và KNOWN_GATEWAYS (tên đọc từ /api/org/info), "
                    "cùng mục cuối **Tổ chức khác: nhập địa chỉ do Quản trị viên cung cấp…**. Ô nhập địa chỉ chỉ hiện khi chọn "
                    "mục này hoặc khi AI Gateway không phản hồi; ô che ký tự, chặn sao chép. Ứng dụng không hiển thị địa chỉ AI "
                    "Gateway; địa chỉ chỉ được ghi nhớ khi xác thực thành công; AI Gateway qua Internet bắt buộc https.")
    cap = replace_picture_fit(doc, "Ảnh màn hình 1. ", SHOT)
    set_caption(doc, cap, f"Trang Đăng nhập của ứng dụng Windows {VER} lần đầu mở: AI Workspace của AI-TECH Center và ô Chọn "
                          "Tổ chức, không hiện địa chỉ AI Gateway")

    # ---- 4.5
    t = doc.table("Cài bộ cài Windows hoặc APK, chọn AI Gateway của Tổ chức")
    r = doc.find_row(t, "Cài bộ cài Windows hoặc APK")
    doc.set_cell(t, r, 1, f"Cài bộ cài Windows hoặc APK {VER} từ trang phát hành, chọn Tổ chức (hoặc Tổ chức khác, nhập địa "
                          "chỉ AI Gateway)", code=False)

    # ---- 4.6
    p = doc.find("Cùng mã nguồn ai-work được đóng gói thành ứng dụng máy trạm", "p")
    doc.set_para(p, "Cùng mã nguồn ai-work được đóng gói thành ứng dụng máy trạm [[ĐANG PHÁT TRIỂN]]. Bản thử nghiệm "
                    f"{VER} cho Windows và Android được build ngay trong WSL Ubuntu và phát hành công khai tại {RELEASES}; "
                    "hướng dẫn chi tiết nằm ở docs/NATIVE_APPS.md của kho ai-work, các script ở scripts/native/.")
    t = doc.table("npm run android:emulator")
    lines = [x.text for x in t.cell(0, 0).paragraphs]
    i = next(k for k, x in enumerate(lines) if x.startswith("npm run release"))
    lines[i:i + 1] = [
        "npm run release -- --copy-windows   # bộ đầy đủ: bộ cài + APK, release/<phiên bản>/ và tệp .zip",
        "bash scripts/native/release.sh --windows [--copy-windows]   # bộ cài Windows + TL1 (PDF)",
        "bash scripts/native/release.sh --android [--copy-windows]   # APK Android + TL1 (PDF)",
    ]
    doc.set_code(t, lines)
    p = doc.find("Phiên bản lấy theo package.json (6.11.16)", "p")
    doc.set_para(p, f"Phiên bản lấy theo package.json ({VER}, theo quy ước 6.tháng.ngày). Android versionCode và iOS build "
                    "(61121) sửa tay trong android/app/build.gradle và ios/App/App.xcodeproj/project.pbxproj khi phát hành bản "
                    "mới; release.sh dừng nếu versionName khác package.json.")
    p = doc.find("DEFAULT_GATEWAY_URL, KNOWN_GATEWAYS trong index.html được đóng gói", "p")
    doc.set_para(p, "DEFAULT_GATEWAY_URL (mặc định AI-TECH Center), KNOWN_GATEWAYS trong index.html được đóng gói vào ứng "
                    "dụng; Tổ chức trong KNOWN_GATEWAYS hiện sẵn trong ô Chọn Tổ chức. Muốn thêm Tổ chức thì sửa index.html "
                    "rồi build lại; không ghi địa chỉ IP nội bộ vào đây vì ứng dụng được phát hành công khai.")

    p = doc.find("npm run release (scripts/native/release.sh) build bộ cài Windows và APK", "p")
    doc.set_para(p, "scripts/native/release.sh có ba chế độ; thư mục release/ bị bỏ qua trong git. Bộ đầy đủ (npm run "
                    "release) build bộ cài Windows và APK, tạo release/<phiên bản>/ và release/iThunder-AI-Workspace-<phiên "
                    "bản>.zip gồm:")
    p = doc.find(f"iThunder-AI-Workspace-6.11.16-Windows-x64-setup.exe: bộ cài NSIS", "p")
    doc.set_para(p, "iThunder-AI-Workspace-<phiên bản>-Windows-x64-setup.exe: bộ cài NSIS, cài cho người dùng hiện tại, chưa "
                    "ký số.")
    p = doc.find("iThunder-AI-Workspace-6.11.16-Android.apk: APK ký bằng khóa phát hành", "p")
    doc.set_para(p, "iThunder-AI-Workspace-<phiên bản>-Android.apk: APK ký bằng khóa phát hành, không debuggable.")
    p = doc.find("SHA256SUMS.txt: mã SHA-256 của hai tệp cài đặt.", "p")
    p = doc.para_after(p, p, "Chế độ **--windows** và **--android** (dùng cho trang phát hành) chỉ build một nền tảng, tạo "
                             "release/<phiên bản>-Windows/ hoặc release/<phiên bản>-Android/ và tệp "
                             "iThunder-AI-Workspace-<phiên bản>-Windows.zip / -Android.zip gồm tệp cài đặt, tài liệu TL1 (PDF, "
                             "lấy từ biến TL1_PDF, đặt tên theo số phiên bản đọc từ trang bìa), HUONG_DAN_CAI_DAT.txt (mẫu "
                             "HUONG_DAN_CAI_DAT_WINDOWS.txt / HUONG_DAN_CAI_DAT_ANDROID.txt) và SHA256SUMS.txt. --copy-windows "
                             "chép tệp ZIP vào thư mục Downloads của Windows.")
    doc.para_after(p, p, f"**GitHub Releases**: bản {VER} được đăng tại {RELEASES} (thẻ v{VER}, đánh dấu Latest, ghi rõ bản "
                         "dùng thử), kèm hai tệp ZIP; ghi chú phát hành ghi mã SHA-256 của từng tệp ZIP, dấu vân tay chứng thư "
                         "APK và kênh góp ý (Wiki, Email, Issues). Khi đóng gói lại cùng phiên bản (ví dụ thay tài liệu TL1), "
                         "tải tệp mới đè lên (gh release upload --clobber) và sửa mã SHA-256 trong ghi chú. Kho ai-workspace "
                         "công khai: không đưa địa chỉ nội bộ, khóa, dữ liệu người dùng vào mã nguồn hay ghi chú phát hành.")

    box = doc.find("Đã kiểm tra ngày 17/11/2026 với bộ phát hành 6.11.16", "tbl")
    doc.set_note(box, f"Đã kiểm tra ngày 21/11/2026 với bộ phát hành {VER}: trên Windows 11 cài bộ cài, mở ứng dụng thật (điều "
                      "khiển WebView2 qua cổng gỡ lỗi của bản thử), trang Đăng nhập tự vào AI-TECH Center và ô Chọn Tổ chức "
                      "liệt kê các Tổ chức của AI Gateway. Trên máy ảo Android 16 (API 36): cài đè APK lên bản 6.11.16 cùng "
                      "khóa ký, versionCode 61121, mở ứng dụng lần đầu (đã xóa dữ liệu) hiện đúng trang Đăng nhập mới; APK tắt "
                      "gỡ lỗi WebView. Ngày 17/11/2026 (6.11.16): cài im lặng (/S), gỡ cài đặt, cài lại và nâng cấp trên "
                      "Windows. Chưa kiểm tra đăng nhập bằng PassCode, FaceID có kiểm tra người thật, gọi thoại / video trên "
                      "ứng dụng và trên điện thoại thật; chưa build iOS.")

    # ---- 8.3
    box = doc.find("Ứng dụng Windows, Android bản 6.11.16 phải build lại", "tbl")
    doc.replace_in(box, "Ứng dụng Windows, Android bản 6.11.16 phải build lại; trong lúc chờ có thể đặt FACEID_LIVENESS=optional.",
                   f"Ứng dụng Windows, Android bản 6.11.16 chưa hỗ trợ kiểm tra người thật: người dùng cài bản {VER} từ trang "
                   "phát hành; trong lúc chờ có thể đặt FACEID_LIVENESS=optional.")

    # ---- 10.1
    p = doc.find("APK phát hành không debuggable, tắt sao lưu dữ liệu và gỡ lỗi WebView", "p")
    doc.set_para(p, "APK phát hành không debuggable, tắt sao lưu dữ liệu và gỡ lỗi WebView; release.sh (cả ba chế độ) kiểm tra "
                    "chữ ký và cấu hình trước khi đóng gói. Ứng dụng máy trạm không hiển thị địa chỉ AI Gateway, ô nhập địa chỉ "
                    "chặn sao chép; bảo vệ thật sự vẫn là xác thực của AI Gateway, TLS, tường lửa / VPN.")

    # ---- 11.3
    t = doc.table("Android không cài được bản mới đè lên bản cũ")
    r = doc.find_row(t, "Máy trạm báo AI Gateway yêu cầu kiểm tra người thật")
    doc.set_cell(t, r, 2, f"Triển khai bản Web mới (no-store cho index.html); người dùng cài ứng dụng {VER} từ trang phát hành; "
                          "tạm thời FACEID_LIVENESS=optional", code=False)
    r = doc.find_row(t, "Android không cài được bản mới đè lên bản cũ")
    doc.add_row(t, ["Người dùng báo mã SHA-256 của tệp ZIP không khớp ghi chú phát hành",
                    "Tệp ZIP đã được đóng gói lại cùng phiên bản (thay tài liệu) sau khi người dùng tải",
                    "Tải lại từ trang phát hành; nếu chỉ đổi tài liệu thì tệp cài đặt bên trong không đổi (đối chiếu "
                    "SHA256SUMS.txt)"], like=r, after=r)

    # ---- Phụ lục G
    t = doc.table("Căn cứ chính của nội dung hiện trạng")
    r = doc.find_row(t, "Căn cứ chính của nội dung hiện trạng")
    doc.set_cell(t, r, 0, f"Mã nguồn ai-platform (AI Platform 6.11.18), ai-work (AI Workspace {VER}), nhánh main ngày {DATE} "
                          "sau khi hợp nhất FaceID liveness (iThunder-AI-OS PR #5, ai-workspace PR #6), Kết nối Kho Đám mây "
                          "(PR #6, PR #7), Ảnh FaceID trong hồ sơ (PR #7, PR #8), ẩn địa chỉ AI Gateway và phóng to (ai-workspace "
                          "PR #9), AI Workspace 6.11.21 (PR #10), thống kê AI Tokens (PR #8, PR #11), bộ phát hành Windows và "
                          "trang Đăng nhập chọn Tổ chức (ai-workspace PR #12 đến #14), bộ phát hành Android (PR #15)", code=False)
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
