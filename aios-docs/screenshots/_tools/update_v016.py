# -*- coding: utf-8 -*-
"""TL2 bản 0.16 (20/11/2026): chức năng Kết nối Kho Dữ liệu với kho đám mây (iThunder-AI-OS PR #6, ai-workspace PR #7).
- Trang bìa; mục "Những điểm mới trong bản 0.16".
- 2.2 phân hệ, 2.3 cổng dịch vụ: Cloud Storage MCP Server.
- 7.7 (mới): dịch vụ hỗ trợ, quy tắc đồng bộ, quyền, thông tin đăng nhập, cài đặt, ảnh màn hình; 7.6 thêm hàng nguồn ngoài.
- 9.1 lệnh khởi động; 10.1 cơ chế an toàn; 10.2 SFTP chưa kiểm tra khóa máy chủ; 11.1 kiểm thử, Bảng 3.
- Phụ lục A.2 (API), B (bảng iAIP_storage_links), C (biến, giới hạn), G (căn cứ mã nguồn).
- Ảnh màn hình mới đánh số theo thứ tự trong tài liệu; đánh lại số các ảnh sau và chỗ dẫn chiếu.
Đầu vào: TL2 bản 0.15 ĐÃ điền mục lục. Đầu ra: bản 0.16 chưa điền mục lục ở <thư mục ra>/<tên>.src.docx, sau đó chép vào
_tools/toc_work/ và chạy build_toc.py TL2.
    python update_v016.py <TL2_v015.docx> <thư mục ra>"""
import copy
import os
import re
import sys

from docx.oxml.ns import qn
from docx.shared import Cm
from docx.table import Table

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docedit import Doc  # noqa: E402
from doc_versioning import blank_para, clone_heading, clone_table, new_section, unfill_toc  # noqa: E402

T2 = "iThunder_AI_OS_TL2_Cai_dat_va_Van_hanh_AI_Platform.docx"
DATE = "20/11/2026"
SHOT = os.path.join(HERE, "..", "tl2_05_cloud_links.png")
SHOT_TITLE = ("Quản trị Kho Tri thức, Kết nối Kho Đám mây: kết nối OneDrive (gương, mỗi ngày) đã đồng bộ và kết nối "
              "WebDAV báo lỗi đăng nhập (dữ liệu mẫu)")

PROVIDERS = [
    ("Microsoft OneDrive / SharePoint, Google Drive, Dropbox",
     "Mã ủy quyền OAuth: chạy rclone authorize \"onedrive\" (\"drive\", \"dropbox\") trên máy có trình duyệt, dán chuỗi JSON. "
     "Không cần đăng ký ứng dụng với nhà cung cấp; Google Drive luôn dùng quyền chỉ đọc"),
    ("S3: Amazon S3, IDrive e2, MinIO, Wasabi, Cloudflare R2", "Access Key ID, Secret Access Key, vùng, endpoint"),
    ("SFTP: đám mây riêng iThunder Clouds, máy chủ dữ liệu", "Máy chủ, cổng, tài khoản; mật khẩu hoặc khóa riêng SSH"),
    ("WebDAV: Nextcloud, ownCloud", "Địa chỉ WebDAV, tài khoản, mật khẩu ứng dụng"),
]
RULES = [
    "**Một chiều, từ đám mây về Kho.** Tệp PDF, DOCX, TXT, MD (bỏ tệp ẩn và tệp lớn hơn CLOUD_LINK_MAX_FILE_MB) được kéo về "
    "thư mục <Kho>/Cloud/<tên kết nối>. AI Platform không bao giờ sửa, xóa dữ liệu trên đám mây; các thư mục khác của Kho "
    "không bị đụng tới.",
    "**Gương** (mặc định): thư mục của kết nối giống hệt đám mây, tệp bị xóa trên đám mây được gỡ khỏi Kho. **Chỉ thêm, cập "
    "nhật**: không gỡ tệp đã kéo về.",
    "Sau mỗi lần đồng bộ có thay đổi, thư mục của kết nối được nạp vào Kho Tri thức như nút Đồng bộ hóa (chỉ tệp mới hoặc đã "
    "sửa; đoạn tri thức của tệp đã gỡ cũng được gỡ). Tắt được theo từng kết nối.",
    "**Lịch**: chỉ khi bấm Đồng bộ ngay, mỗi giờ, 6 giờ hoặc 1 ngày. Đồng bộ theo lịch chạy dưới quyền người cấu hình gần "
    "nhất, kiểm tra lại trên OpenLDAP mỗi lần: còn trong Tổ chức, không bị khóa, còn quyền quản trị Kho. Nhiều worker không "
    "chạy chồng một kết nối. Mỗi lần kéo tối đa CLOUD_LINK_MAX_TOTAL_MB, lần sau kéo tiếp.",
    "**Quyền**: người có quyền quản trị (manage) trên Kho; Kho Cá nhân chỉ chính người dùng thấy, kể cả Quản trị viên cũng "
    "không thấy. Tạo, sửa, xóa kết nối cần xác nhận lại PassCode (mục 6.4); mọi thao tác ghi nhật ký storage_link.*.",
    "**Thông tin đăng nhập** mã hóa Fernet trong bảng iAIP_storage_links, khóa dẫn xuất từ CLOUD_AI_SECRET (hoặc "
    "JWT_SECRET_KEY) và ORGID, gắn với Kho và thư mục của kết nối. API không bao giờ trả lại; ô nhập để trống khi sửa là giữ "
    "nguyên. Mã ủy quyền OAuth được rclone làm mới thì lưu lại.",
]
INSTALL = [
    "sudo apt install rclone                    # hoặc bản mới từ rclone.org (đặt RCLONE_BIN)",
    "openssl rand -hex 32                       # đặt vào CLOUD_STORAGE_MCP_TOKEN trong .env",
    "/usr/local/venv/bin/python3 -m integrations.cloud_storage_mcp check",
    "scripts/start_cloud_storage_mcp.sh         # nghe 127.0.0.1:9120, nhật ký cloud-storage-mcp.log",
]


def add_rows(doc, table, rows, like=-1):
    for values in rows:
        doc.add_row(table, list(values), like=like)


def bullet_after(doc, anchor, template, markup):
    return doc.para_after(anchor, template, markup)


def picture_after(doc, anchor, img, title, number):
    """Ảnh + chú thích theo mẫu Ảnh màn hình 2 (rộng 16 cm)."""
    cap_tpl = doc.find("Ảnh màn hình 2. ", "p")
    pic_tpl = cap_tpl.getprevious()
    p = doc.d.add_paragraph()
    p.add_run().add_picture(img, width=Cm(16.0))
    el = p._p
    old = el.find(qn("w:pPr"))
    if old is not None:
        el.remove(old)
    el.insert(0, copy.deepcopy(pic_tpl.find(qn("w:pPr"))))
    anchor.addnext(el)
    cap = copy.deepcopy(cap_tpl)
    runs = cap.findall(qn("w:r"))
    runs[0].find(qn("w:t")).text = f"Ảnh màn hình {number}. "
    runs[1].find(qn("w:t")).text = title
    for r in runs[2:]:
        cap.remove(r)
    el.addnext(cap)
    doc.changes.append(f"Thêm ảnh màn hình: {title[:60]}")
    return cap


def renumber_shots(doc):
    """Đánh lại số Ảnh màn hình theo thứ tự trong tài liệu; sửa chỗ dẫn chiếu 'Ảnh màn hình N' trong câu."""
    mapping, n = {}, 0
    for e in doc.elements():
        if e.tag != qn("w:p"):
            continue
        runs = e.findall(qn("w:r"))
        t0 = runs[0].find(qn("w:t")) if runs else None
        m = re.fullmatch(r"Ảnh màn hình (\d+)\. ", t0.text or "") if t0 is not None else None
        if m:
            n += 1
            mapping[m.group(1)] = n
            t0.text = f"Ảnh màn hình {n}. "
    for e in doc.elements():
        for t in e.iter(qn("w:t")):
            if t.text and "Ảnh màn hình" in t.text and not re.fullmatch(r"Ảnh màn hình \d+\. ", t.text):
                t.text = re.sub(r"Ảnh màn hình (\d+)(?=[,) ])", lambda m: f"Ảnh màn hình {mapping.get(m.group(1), m.group(1))}"
                                if m.group(1) in mapping else m.group(0), t.text)
    return mapping


def tl2(src, dst):
    doc = Doc(src)
    unfill_toc(doc)
    blank = blank_para(doc)

    # ---- Trang bìa, những điểm mới
    t = doc.table("Phiên bản tài liệu")
    doc.set_cell(t, doc.find_row(t, "Phiên bản tài liệu"), 1,
                 f"0.16, bản dự thảo ngày {DATE} (cập nhật bản 0.15: Kết nối Kho Dữ liệu với kho đám mây)")
    doc.set_cell(t, doc.find_row(t, "Phiên bản phần mềm"), 1,
                 f"AI Workspace 6.11.16 và AI Platform 6.11.18 (nhánh main ngày {DATE}, sau khi hợp nhất PR #6 của "
                 "iThunder-AI-OS và PR #7 của ai-workspace: Kết nối Kho Đám mây)")
    new_section(doc, "Những điểm mới trong bản 0.15", "Những điểm mới trong bản 0.16",
                "Bổ sung chức năng Kết nối Kho Dữ liệu với kho đám mây công cộng (OneDrive, Google Drive, Dropbox, S3) và đám "
                "mây riêng (SFTP của iThunder Clouds, WebDAV) qua Cloud Storage MCP Server (iThunder-AI-OS PR #6, "
                "ai-workspace PR #7). Các thay đổi và nơi mô tả:",
                "Đối chiếu hai kho mã nguồn ngày",
                "Số phiên bản 6.11.18 trong aiplatform.py",
                [("Kết nối Kho Dữ liệu với kho đám mây: dịch vụ hỗ trợ, quy tắc đồng bộ, quyền, cài đặt", "7.7"),
                 ("Cloud Storage MCP Server (rclone), cổng 9120 chỉ trên 127.0.0.1", "2.2, 2.3, 9.1"),
                 ("Cơ chế an toàn, điểm cần tăng cường", "10.1, 10.2"),
                 ("API /api/storage-links, bảng iAIP_storage_links, biến cấu hình và giới hạn", "A.2, B, C"),
                 ("Kết quả kiểm thử tests/storage_links (gồm rclone thật với WebDAV, S3 giả lập)", "11.1")])

    # ---- 2.2, 2.3
    t = doc.table("Phân hệTệp chính")
    doc.add_row(t, ["Kết nối Kho đám mây", "storage_links/, integrations/cloud_storage_mcp/",
                    "Kéo tệp từ OneDrive, Google Drive, Dropbox, S3, SFTP, WebDAV về Kho Dữ liệu qua MCP Server dùng "
                    "rclone (tiến trình riêng); lịch đồng bộ, nạp Kho Tri thức", "[[ĐÃ CÓ]]"])
    t = doc.table("Dịch vụCổng mặc định")
    doc.add_row(t, ["Cloud Storage MCP Server", "9120/TCP, 127.0.0.1", "AI Platform",
                    "Tiến trình riêng cùng máy, cùng tài khoản với AI Platform; ra Internet qua HTTPS (OneDrive, S3...) "
                    "hoặc tới đám mây riêng khai trong CLOUD_STORAGE_MCP_PRIVATE_HOSTS"])

    # ---- 7.6: thêm hàng; 7.7 mới
    t = doc.table("Yêu cầu TARAHiện trạng")
    doc.add_row(t, ["Kết nối nguồn tài liệu bên ngoài (OneDrive, S3, đám mây riêng)",
                    "Kéo một chiều qua Cloud Storage MCP Server, có lịch, nạp Kho Tri thức (mục 7.7); chưa giữ chủ sở hữu, "
                    "phiên bản của tệp ở nguồn", "[[ĐANG PHÁT TRIỂN]]"])
    h8 = doc.heading("8. Các chức năng AI")
    anchor = h8.getprevious()
    h = clone_heading(doc, "7.6. Khoảng cách tới Knowledge Object của TPS-KR", anchor,
                      "7.7. Kết nối Kho Dữ liệu với kho đám mây")
    intro = doc.para_after(h, "Chia đoạn (knowledge/chunking.py)",
                           "[[ĐÃ CÓ]] Mỗi Kho Dữ liệu (kho của Tổ chức, Kho phân vùng, Kho Cá nhân của từng người) có tối đa "
                           "5 kết nối (CLOUD_LINK_MAX_PER_STORE) tới kho đám mây. AI Platform là MCP Client của **Cloud "
                           "Storage MCP Server** (integrations/cloud_storage_mcp), tiến trình dùng rclone để đọc đám mây và "
                           "kéo tệp về Kho. Giao diện: **Quản trị Kho Tri thức** › chọn Kho › **Kết nối Kho Đám mây** (chỉ "
                           "hiện với người có quyền quản trị Kho).")
    tbl_el, tbl = clone_table(doc, "Sự kiệnHệ thống làm gì", intro, 2)
    doc.replace_in(tbl.cell(0, 0)._tc, "Sự kiện", "Dịch vụ")          # giữ định dạng hàng tiêu đề
    doc.replace_in(tbl.cell(0, 1)._tc, "Hệ thống làm gì", "Đăng nhập")
    doc.set_cell(tbl, 1, 0, PROVIDERS[0][0], code=False)
    doc.set_cell(tbl, 1, 1, PROVIDERS[0][1], code=False)
    add_rows(doc, tbl, PROVIDERS[1:])
    gap = copy.deepcopy(blank)
    tbl_el.addnext(gap)
    bullet_tpl = doc.find("Khóa là DN. gidNumber", "p")
    last = gap
    for rule in RULES:
        last = bullet_after(doc, last, bullet_tpl, rule)
    p = doc.para_after(last, "Chia đoạn (knowledge/chunking.py)",
                       "Cài đặt trên máy AI Platform (README của ai-platform, mục 6.4; integrations/cloud_storage_mcp/README.md):")
    code = copy.deepcopy(doc.find("cd ai-work", "tbl", nth=0))      # khối mã mẫu của mục 4.3
    p.addnext(code)
    doc.set_code(Table(code, doc.d), INSTALL)
    gap2 = copy.deepcopy(blank)
    code.addnext(gap2)
    last = doc.note_after(gap2, "Lưu ý",
                          "Mặc định chỉ cho kết nối tới địa chỉ Internet công cộng, để người quản trị một Kho không dò quét "
                          "được mạng nội bộ. Máy chủ đám mây riêng (iThunder Clouds, MinIO, Nextcloud nội bộ) phải được khai "
                          "trong CLOUD_STORAGE_MCP_PRIVATE_HOSTS (tên máy, IP hoặc dải CIDR). MCP Server chỉ nhận các tùy chọn "
                          "rclone trong danh sách cho phép (chặn ssh, key_file, env_auth...) và chỉ ghi vào thư mục Kho. Nên "
                          "dùng tài khoản hoặc khóa chỉ có quyền đọc trên đám mây.", blank_template=copy.deepcopy(blank))
    picture_after(doc, last, SHOT, SHOT_TITLE, 99)

    # ---- 9.1
    t = Table(doc.heading("9.1. Khởi động và dừng").getnext(), doc.d)
    assert doc.text_of(t._tbl).startswith("scripts/start_aiplatform.sh")
    lines = [p.text for p in t.cell(0, 0).paragraphs]
    doc.set_code(t, lines + ["scripts/start_cloud_storage_mcp.sh   # Cloud Storage MCP Server (khi .env có CLOUD_STORAGE_MCP_TOKEN)"])

    # ---- 10.1, 10.2
    p = doc.find("Khóa API dịch vụ AI On-Clouds do admin nhập", "p")
    doc.para_after(p, p, "Kết nối Kho đám mây: thông tin đăng nhập mã hóa, không trả về qua API, thao tác ghi cần step-up; "
                         "MCP Server chỉ nghe 127.0.0.1, xác thực bằng mã riêng, chặn tùy chọn rclone nguy hiểm, máy chủ nội "
                         "bộ chưa khai báo và đường dẫn ra ngoài Kho; đồng bộ một chiều, không ghi lên đám mây (mục 7.7).")
    t = doc.table("Ưu tiênVấn đềKhuyến nghị")
    doc.add_row(t, ["P2", "Kết nối SFTP tới kho đám mây chưa kiểm tra khóa máy chủ (host key); máy chủ được phép kiểm tra "
                    "theo DNS lúc kết nối", "Khai báo dấu vân tay khóa máy chủ cho từng kết nối SFTP; dùng tên máy cố định "
                    "trong CLOUD_STORAGE_MCP_PRIVATE_HOSTS"])

    # ---- 11.1
    t = doc.table("AIP=.venv")
    lines = [p.text for p in t.cell(0, 0).paragraphs]
    i = lines.index("$AIP tests/ai_server/test_ai_server.py") + 1
    lines[i:i] = ["$AIP tests/storage_links/test_storage_links.py   # thêm RCLONE_BIN=<rclone>: chạy rclone thật"]
    doc.set_code(t, lines)
    t = doc.table("Bộ kiểm thửKết quả ngày")
    r = doc.find_row(t, "tests/identity/run_all.py")
    doc.set_cell(t, r, 1, "Đạt (16 tệp, gồm test_14_rag_chat.py so sánh trực tiếp lời nhắc của AI Chat/Work và Trao đổi trực "
                          "tuyến, test_17_storage_link_hooks.py)", code=False)
    r = doc.find_row(t, "tests/cloud_ai/test_cloud_ai.py")
    doc.add_row(t, ["tests/storage_links/test_storage_links.py", "Đạt (11 nhóm)",
                    "MCP Server, dịch vụ, API với bộ máy giả lập; rclone v1.68.2 thật với WebDAV, S3 giả lập trên 127.0.0.1; "
                    "bảng iAIP_storage_links thử trên PostgreSQL thật. Chưa thử OneDrive, Google Drive, S3 thật"],
                like=r, after=r)

    # ---- Phụ lục A.2, B, C, G
    t = doc.table("POST /api/chat/rag-stream")
    for values in [
        ("GET /api/storage-links/providers; GET /api/storage-links?storage=<kho>",
         "Dịch vụ đám mây hỗ trợ, trường cấu hình, lịch; kết nối của Kho (không có thông tin đăng nhập)"),
        ("POST /api/storage-links; PATCH, DELETE /api/storage-links/{id}",
         "Tạo, sửa, xóa kết nối đám mây; quyền quản trị Kho, cần step-up"),
        ("POST /api/storage-links/browse, /api/storage-links/{id}/sync",
         "Xem thư mục trên đám mây; Đồng bộ ngay (chạy nền)"),
    ]:
        doc.add_row(t, list(values))
    t = doc.table("BảngDùng cho")
    r = doc.find_row(t, "iAIP_cloud_ai_keys")
    doc.add_row(t, ["iAIP_storage_links", "Kết nối Kho đám mây: thông tin đăng nhập (mã hóa), lịch, kết quả lần đồng bộ "
                    "gần nhất", "init_database.sh, mã nguồn"], like=r, after=r)
    t = doc.table("NhómBiến trong .env")
    doc.add_row(t, ["Kết nối Kho đám mây", "CLOUD_STORAGE_MCP_TOKEN, CLOUD_LINKS_ENABLED, CLOUD_LINK_MCP_URL, CLOUD_LINK_SUBDIR, "
                    "CLOUD_LINK_MAX_PER_STORE, CLOUD_LINK_MAX_FILE_MB, CLOUD_LINK_MAX_TOTAL_MB, CLOUD_LINK_SYNC_TIMEOUT; MCP "
                    "Server: CLOUD_STORAGE_MCP_ROOTS, CLOUD_STORAGE_MCP_HOST, CLOUD_STORAGE_MCP_PORT, "
                    "CLOUD_STORAGE_MCP_PRIVATE_HOSTS, CLOUD_STORAGE_MCP_TIMEOUT, CLOUD_STORAGE_MCP_MAX_PARALLEL, RCLONE_BIN"])
    t = doc.table("Giới hạn trong config.py")
    doc.add_row(t, ["Kết nối Kho đám mây", "5 kết nối mỗi Kho; 100 MB mỗi tệp; 5.000 MB và 3.600 giây mỗi lần đồng bộ; "
                    "lịch 1 giờ, 6 giờ, 1 ngày"])
    t = doc.table("Căn cứ chính của nội dung hiện trạng")
    r = doc.find_row(t, "Căn cứ chính của nội dung hiện trạng")
    doc.set_cell(t, r, 0, f"Mã nguồn ai-platform (AI Platform 6.11.18), ai-work (AI Workspace 6.11.16), nhánh main ngày {DATE} "
                          "sau khi hợp nhất Kết nối Kho Đám mây (iThunder-AI-OS PR #6, ai-workspace PR #7)", code=False)

    mapping = renumber_shots(doc)
    doc.changes.append(f"Đánh lại số ảnh màn hình: {mapping}")
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
