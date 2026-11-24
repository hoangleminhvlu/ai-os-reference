# -*- coding: utf-8 -*-
"""TL2 bản 0.23 (24/11/2026): thông báo khuyến nghị cài ứng dụng trước trang Đăng nhập của bản Web (ai-workspace PR #19).
- Trang bìa; mục "Những điểm mới trong bản 0.23"; 2.4, 4.3, 4.5, 8.3 (hộp Lưu ý), 11.3, Phụ lục G.
Đầu vào: TL2 bản 0.22 ĐÃ điền mục lục. Đầu ra: bản 0.23 chưa điền mục lục ở <thư mục ra>/<tên>.src.docx, sau đó chép vào
_tools/toc_work/ và chạy build_toc.py TL2.
    python update_v023.py <TL2_v022.docx> <thư mục ra>"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docedit import Doc  # noqa: E402
from doc_versioning import new_section, unfill_toc  # noqa: E402

T2 = "iThunder_AI_OS_TL2_Cai_dat_va_Van_hanh_AI_Platform.docx"
DATE, VER = "24/11/2026", "6.11.23"


def tl2(src, dst):
    doc = Doc(src)
    unfill_toc(doc)
    t = doc.table("Phiên bản tài liệu")
    doc.set_cell(t, doc.find_row(t, "Phiên bản tài liệu"), 1,
                 f"0.23, bản dự thảo ngày {DATE} (cập nhật bản 0.22: thông báo khuyến nghị cài ứng dụng trên bản Web)")
    doc.set_cell(t, doc.find_row(t, "Phiên bản phần mềm"), 1,
                 f"AI Workspace {VER} và AI Platform 6.11.18 (nhánh main ngày {DATE}, sau khi hợp nhất PR #5 đến #9 của iThunder-AI-OS và "
                 "PR #6 đến #19 của ai-workspace)")
    new_section(doc, "Những điểm mới trong bản 0.22", "Những điểm mới trong bản 0.23",
                "Bổ sung thông báo khuyến nghị cài ứng dụng trước trang Đăng nhập của bản Web (ai-workspace PR #19). Các thay đổi và nơi mô tả:",
                "Bổ sung bộ phát hành ứng dụng 6.11.23 (ai-workspace PR #17)",
                "Bộ phát hành 6.11.23 trên GitHub Releases, kết quả kiểm tra",
                [("Thông báo trên bản Web: khi nào hiện, dẫn tới đâu", "2.4"),
                 ("Triển khai bản Web: thứ tự với AI Platform, mạng không có Internet", "4.3, 11.3"),
                 ("Kiểm tra sau cài đặt; thông báo không thay kiểm soát truy cập", "4.5, 8.3")])

    # ---- 2.4
    p = doc.find("Trước khi nhập PassCode, AI Workspace gọi GET /api/org/info", "p")
    last = doc.find("Ứng dụng máy trạm không có tên miền của Tổ chức", "p")
    tpl = doc.find("GATEWAY_NETWORK=internet: trang Đăng nhập cho chọn Tổ chức khác", "p")
    doc.para_after(last, tpl, "Bản Web (clientKind = web) hiện thông báo Nên dùng ứng dụng AI Workspace trước trang Đăng nhập (ai-workspace PR #19): "
                              "lý do dùng ứng dụng, hai nút mở trang phát hành mới nhất trên GitHub, nút Tiếp tục bằng Web App. Thông báo hiện một lần "
                              "mỗi thẻ trình duyệt (sessionStorage, khóa tara_web_notice); không hiện trong ứng dụng Windows, Android và khi phiên đăng "
                              "nhập còn lưu trên trình duyệt. Trang Đăng nhập có nút Tải ứng dụng Windows / Android để mở lại.")
    assert p is not None

    # ---- 4.3
    p = doc.find("Bản build dùng chung cho mọi Tổ chức: AI Workspace gọi /api và /ws", "p")
    q = doc.para_after(p, p, "Thông báo khuyến nghị cài ứng dụng nằm sẵn trong bản build, không có biến cấu hình. Địa chỉ tải và hướng dẫn là hằng số "
                             "RELEASE_URL (https://github.com/hoangleminhvlu/ai-workspace/releases/latest) và GUIDE_URL trong src/auth/appDownload.js; "
                             "nút tải mở trang phát hành mới nhất nên không phải build lại Web khi có bản ứng dụng mới.")
    doc.para_after(q, p, "Thứ tự triển khai: cập nhật AI Platform có khóa thiết bị (iThunder-AI-OS PR #9) trước hoặc cùng lúc với bản Web này. Với AI "
                         f"Platform cũ, ứng dụng {VER} vẫn phải kiểm tra người thật nên lời khuyến nghị chưa đúng. Tổ chức có mạng không ra "
                         "Internet: đặt bộ cài trên máy chủ nội bộ, sửa RELEASE_URL (và GUIDE_URL) rồi build lại.")

    # ---- 4.5
    t = doc.table("Cài bộ cài Windows hoặc APK 6.11.23 từ trang phát hành")
    r = doc.find_row(t, "Cài bộ cài Windows hoặc APK 6.11.23")
    doc.add_row(t, ["Thông báo trên bản Web", "Mở AI Workspace bằng trình duyệt trong thẻ mới (chưa đăng nhập); bấm nút tải, rồi Tiếp tục bằng Web App",
                    "Thông báo hiện trước trang Đăng nhập, nút tải mở trang phát hành; ứng dụng Windows, Android không hiện"], like=r, after=r)

    # ---- 8.3
    box = doc.find("Liveness chủ động chống ảnh tĩnh 2D; chưa chống được mặt nạ 3D", "tbl")
    doc.replace_in(box, f"ứng dụng {VER} cũng vậy khi AI Platform chưa cập nhật.",
                   f"ứng dụng {VER} cũng vậy khi AI Platform chưa cập nhật. Thông báo Nên dùng ứng dụng trên bản Web (mục 2.4) chỉ là khuyến nghị ở "
                   "giao diện, không phải biện pháp kiểm soát: AI Platform vẫn tự phân loại máy trạm và bắt buộc Liveness với Web.")

    # ---- 11.3
    t = doc.table("Người dùng không thấy giao diện mới sau khi triển khai")
    r = doc.find_row(t, "Người dùng không thấy giao diện mới sau khi triển khai")
    for row in [["Bấm nút tải ứng dụng trên bản Web nhưng không mở được trang GitHub", "Mạng không ra Internet hoặc chặn GitHub",
                 "Cấp bộ cài qua máy chủ nội bộ; sửa RELEASE_URL trong src/auth/appDownload.js, build lại (mục 4.3)"],
                ["Đã cài ứng dụng theo khuyến nghị nhưng vẫn phải quay mặt", "AI Platform chưa có khóa thiết bị, hoặc FACEID_APP_LIVENESS=required",
                 "Cập nhật AI Platform (iThunder-AI-OS PR #9); kiểm tra FACEID_APP_LIVENESS (mục 8.3)"]]:
        r = doc.add_row(t, row, like=r, after=r)

    # ---- Phụ lục G
    t = doc.table("Căn cứ chính của nội dung hiện trạng")
    r = doc.find_row(t, "Căn cứ chính của nội dung hiện trạng")
    old = t.cell(r, 0).text
    new = old.replace("nhánh main ngày 23/11/2026", f"nhánh main ngày {DATE}").replace(
        f"AI Workspace {VER} (PR #17)", f"AI Workspace {VER} (PR #17), thông báo khuyến nghị cài ứng dụng trên bản Web (ai-workspace PR #19)")
    assert DATE in new and "PR #19" in new
    doc.set_cell(t, r, 0, new, code=False)
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
