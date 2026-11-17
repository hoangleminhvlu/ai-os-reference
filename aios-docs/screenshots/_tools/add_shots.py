# -*- coding: utf-8 -*-
"""Chèn ảnh màn hình thật (AI Workspace nhánh main, AI Gateway giả lập, dữ liệu mẫu) vào TL1, TL2 bản 0.10.
Chạy trên bản sao lưu trước khi chèn ảnh (bak_before_shots) để chạy lại được nhiều lần."""
import copy
import os
import sys

from docx.oxml.ns import qn
from docx.shared import Cm

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docedit import Doc  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
IMG = os.path.join(HERE, "..", "shots2", "doc")
SRC = os.path.join(HERE, "bak_before_shots")
OUT = "/home/local/ai-reference/aios-docs"
NOTE = ("Ảnh màn hình trong tài liệu được chụp từ AI Workspace bản 0.10 với tài khoản và dữ liệu mẫu, "
        "không phải dữ liệu thật của Tổ chức; giao diện của bạn có thể khác đôi chút về màu sắc và nội dung.")


class Shots:
    def __init__(self, doc):
        self.doc = doc
        self.cap_tpl = doc.find("Hình 1. ", "p")
        self.pic_tpl = self.cap_tpl.getprevious()
        assert self.pic_tpl.find(".//" + qn("w:drawing")) is not None
        self.n = 0

    def _make(self, img, width_cm, title):
        self.n += 1
        p = self.doc.d.add_paragraph()
        p.add_run().add_picture(os.path.join(IMG, img), width=Cm(width_cm))
        el = p._p
        old = el.find(qn("w:pPr"))
        if old is not None:
            el.remove(old)
        el.insert(0, copy.deepcopy(self.pic_tpl.find(qn("w:pPr"))))
        cap = copy.deepcopy(self.cap_tpl)
        runs = cap.findall(qn("w:r"))
        runs[0].find(qn("w:t")).text = f"Ảnh màn hình {self.n}. "
        runs[1].find(qn("w:t")).text = title
        for r in runs[2:]:
            cap.remove(r)
        self.doc.changes.append(f"Ảnh màn hình {self.n}: {title}")
        return el, cap

    def replace(self, placeholder_needle, img, title, width_cm=16.0):
        tbl = self.doc.find(placeholder_needle, "tbl")
        el, cap = self._make(img, width_cm, title)
        tbl.addprevious(el)
        el.addnext(cap)
        tbl.getparent().remove(tbl)

    def before_heading(self, heading, items):
        anchor = self.doc.heading(heading).getprevious()
        for img, title, *w in items:
            el, cap = self._make(img, w[0] if w else 16.0, title)
            anchor.addnext(el)
            el.addnext(cap)
            anchor = cap


def note(doc, intro_needle, after_table_needle):
    tpl = doc.find(intro_needle, "p")
    tbl = doc.find(after_table_needle, "tbl")
    doc.para_after(tbl, tpl, NOTE)


def tl1():
    name = "iThunder_AI_OS_TL1_Huong_dan_su_dung_AI_Workspace.docx"
    doc = Doc(os.path.join(SRC, name))
    s = Shots(doc)
    note(doc, "So với bản 0.9, AI Workspace có thêm", "Tùy vai trò và cấu hình của tổ chức")
    # theo thứ tự xuất hiện trong tài liệu để số ảnh tăng dần
    s.replace("[Chèn ảnh màn hình 1]", "tl1_01_login.jpg",
              "Trang Đăng nhập: khung AI Workspace của Tổ chức, ô PassCode, GetCode và Register")
    s.replace("[Chèn ảnh màn hình 2]", "tl1_02_faceid_c.png",
              "Kích hoạt Trắc sinh học FaceID (ảnh camera được thay bằng hình minh họa)", 7.0)
    s.replace("[Chèn ảnh màn hình 3]", "tl1_03_workspace.png",
              "AI Workspace với đủ thanh bên trái, khung AI Chat/Work và thanh bên phải")
    s.before_heading("4.7. Sao chép và lưu câu trả lời", [
        ("tl1_07_answer.png", "Một câu trả lời trên Kho Công khai: dòng thông báo, số nguồn [n], danh mục nguồn, "
                              "mô hình On-Premises, thời gian và số token")])
    s.replace("[Chèn ảnh màn hình 4]", "tl1_04_public.png", "Hộp thoại Mở Kho của phân vùng Public với kết quả tìm “LAB”")
    s.replace("[Chèn ảnh màn hình 5]", "tl1_05_history.png", "Lịch sử Làm việc: danh sách phiên và nội dung phiên đang chọn")
    s.before_heading("8.4. Quản lý hội thoại", [
        ("tl1_08_comm.png", "Gọi @ai trong hội thoại nhóm: Trợ lý AI chào từng thành viên, trích nguồn và ghi "
                            "mô hình, thời gian, số token")])
    s.replace("[Chèn ảnh màn hình 6]", "tl1_06_identity.png", "Quản trị Định danh, tab Người dùng và hồ sơ chi tiết (tài khoản mẫu)")
    doc.save(os.path.join(OUT, name))
    return doc


def tl2():
    name = "iThunder_AI_OS_TL2_Cai_dat_va_Van_hanh_AI_Platform.docx"
    doc = Doc(os.path.join(SRC, name))
    s = Shots(doc)
    note(doc, "Bản này đối chiếu lại với mã nguồn nhánh main", "AI Chat/Work và Trợ lý AI của Trao đổi trực tuyến dùng chung")
    s.before_heading("2.5. Luồng một câu hỏi RAG", [
        ("tl2_01_login_gw.jpg", "Trang Đăng nhập của ứng dụng máy trạm: AI Gateway On-Clouds đã chọn, Tổ chức lấy từ "
                                "/api/org/info với GATEWAY_NETWORK=internet")])
    s.before_heading("7.4. Vòng đời Kho Tri thức của phân vùng", [
        ("tl2_04_clouds.png", "Quản trị Định danh, tab Phân vùng: phân vùng iDragonCloud đang dùng và Kho Tri thức đã lưu trữ")])
    s.before_heading("8.7. Giám sát GPU và máy chủ", [
        ("tl1_08_comm.png", "Trợ lý AI trả lời @ai trong hội thoại nhóm, kèm mô hình, vị trí On-Premises, thời gian và số token")])
    s.before_heading("8.8. Bằng chứng của từng câu trả lời", [
        ("tl2_02_server.png", "Giám sát Hệ thống, tab Máy chủ AI: mô hình đã cài và đang nạp, GPU, CPU, bộ nhớ, nhật ký"),
        ("tl2_03_platform_c.png", "Giám sát Hệ thống, tab AI Platform: trạng thái dịch vụ nền, cảnh báo cấu hình, tiến trình và Tổ chức")])
    doc.save(os.path.join(OUT, name))
    return doc


if __name__ == "__main__":
    for f in (tl1, tl2):
        d = f()
        print(f.__name__, len(d.changes), "thay đổi")
        for c in d.changes:
            print("  -", c)
