# -*- coding: utf-8 -*-
"""TL2 bản 0.22 (23/11/2026): ứng dụng 6.11.23 đã đăng trang phát hành (v6.11.23, ai-workspace PR #17).
- Trang bìa; mục "Những điểm mới trong bản 0.22"; 4.5, 4.6, 8.3 (hộp Lưu ý), Phụ lục G.
Đầu vào: TL2 bản 0.21 ĐÃ điền mục lục. Đầu ra: bản 0.22 chưa điền mục lục ở <thư mục ra>/<tên>.src.docx, sau đó chép vào
_tools/toc_work/ và chạy build_toc.py TL2.
    python update_v022.py <TL2_v021.docx> <thư mục ra>"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docedit import Doc  # noqa: E402
from doc_versioning import new_section, unfill_toc  # noqa: E402
from update_tl1_v021 import bump  # noqa: E402

T2 = "iThunder_AI_OS_TL2_Cai_dat_va_Van_hanh_AI_Platform.docx"
DATE, OLD, VER = "23/11/2026", "6.11.21", "6.11.23"


def tl2(src, dst):
    doc = Doc(src)
    unfill_toc(doc)
    t = doc.table("Phiên bản tài liệu")
    doc.set_cell(t, doc.find_row(t, "Phiên bản tài liệu"), 1,
                 f"0.22, bản dự thảo ngày {DATE} (cập nhật bản 0.21: ứng dụng {VER} trên trang phát hành)")
    doc.set_cell(t, doc.find_row(t, "Phiên bản phần mềm"), 1,
                 f"AI Workspace {VER} và AI Platform 6.11.18 (nhánh main ngày {DATE}, sau khi hợp nhất PR #5 đến #9 của iThunder-AI-OS và "
                 "PR #6 đến #17 của ai-workspace)")
    new_section(doc, "Những điểm mới trong bản 0.21", "Những điểm mới trong bản 0.22",
                f"Bổ sung bộ phát hành ứng dụng {VER} (ai-workspace PR #17). Các thay đổi và nơi mô tả:",
                "Bổ sung quy trình FaceID mới (iThunder-AI-OS PR #9",
                "Ngưỡng độ khớp thống nhất 0,6 (FACEID_ACCURACY)",
                [(f"Bộ phát hành {VER} trên GitHub Releases, kết quả kiểm tra", "4.5, 4.6"),
                 (f"Ứng dụng {VER} với AI Platform chưa cập nhật", "4.6, 8.3")])

    # ---- 4.5, 4.6
    t = doc.table("Cài bộ cài Windows hoặc APK 6.11.21 từ trang phát hành")
    r = doc.find_row(t, "Cài bộ cài Windows hoặc APK 6.11.21")
    doc.set_cell(t, r, 1, f"Cài bộ cài Windows hoặc APK {VER} từ trang phát hành, chọn Tổ chức (hoặc Tổ chức khác, nhập địa chỉ AI Gateway)",
                 code=False)
    doc.set_cell(t, r, 2, "Trang Đăng nhập hiện đúng Tổ chức; thiết bị mới nhận được mã qua Email, đăng ký và đăng nhập được")
    bump(doc, doc.find("Cùng mã nguồn ai-work được đóng gói thành ứng dụng máy trạm", "p"), 1)
    p = doc.find("Phiên bản lấy theo package.json (6.11.21", "p")
    doc.set_para(p, f"Phiên bản lấy theo package.json ({VER}, theo quy ước 6.tháng.ngày). Android versionCode và iOS build (61123) sửa tay trong "
                    "android/app/build.gradle và ios/App/App.xcodeproj/project.pbxproj khi phát hành bản mới; release.sh dừng nếu versionName "
                    "khác package.json.")
    p = doc.find("GitHub Releases: bản 6.11.21 được đăng tại", "p")
    bump(doc, p, 2)
    box = doc.find("Đã kiểm tra ngày 21/11/2026 với bộ phát hành 6.11.21", "tbl")
    doc.replace_in(box, "Ngày 23/11/2026, bản thử 6.11.23 (chưa đăng trang phát hành): ứng dụng",
                   f"Ngày 23/11/2026, bộ phát hành {VER} (dùng đúng tệp đã kiểm tra): ứng dụng")
    doc.replace_in(box, "với AI Platform dùng LDAP, Email, khuôn mặt giả lập.",
                   "với AI Platform dùng LDAP, Email, khuôn mặt giả lập. Giao diện mới với AI Platform trước PR #9 (kiểm tra trong Chromium): tự lùi về "
                   "kiểm tra người thật, không hỏi mã.")

    # ---- 8.3
    box = doc.find("Liveness chủ động chống ảnh tĩnh 2D; chưa chống được mặt nạ 3D", "tbl")
    doc.replace_in(box, "Ứng dụng 6.11.21 vẫn dùng được nhưng phải kiểm tra người thật.",
                   f"Ứng dụng 6.11.21 vẫn dùng được nhưng phải kiểm tra người thật; ứng dụng {VER} cũng vậy khi AI Platform chưa cập nhật.")

    # ---- Phụ lục G
    t = doc.table("Căn cứ chính của nội dung hiện trạng")
    r = doc.find_row(t, "Căn cứ chính của nội dung hiện trạng")
    old = t.cell(r, 0).text
    new = old.replace("ứng dụng 6.11.21 và bản thử 6.11.23", f"ứng dụng {VER}").replace(
        "(iThunder-AI-OS PR #9, ai-workspace PR #16)", f"(iThunder-AI-OS PR #9, ai-workspace PR #16), AI Workspace {VER} (PR #17)")
    assert new.count(VER) == 2 and "bản thử" not in new
    doc.set_cell(t, r, 0, new, code=False)
    doc.save(dst)
    return doc


if __name__ == "__main__":
    src, out = sys.argv[1:3]
    os.makedirs(out, exist_ok=True)
    d = tl2(src, os.path.join(out, T2[:-5] + ".src.docx"))
    print("tl2", len(d.changes), "thay đổi")
