# -*- coding: utf-8 -*-
"""Điền sẵn mục lục (trường TOC \\o "1-2") cho TL1, TL2 khi không có Word:
- gắn bookmark _Toc… vào mỗi tiêu đề Heading1/Heading2,
- viết kết quả trường TOC gồm các dòng có liên kết, dấu chấm dẫn và PAGEREF (Word vẫn Update Field được),
- số trang lấy từ PDF do LibreOffice xuất; chạy hai lượt để số trang ổn định sau khi mục lục chiếm chỗ.
"""
import copy
import os
import re
import shutil
import subprocess
import sys

import docx
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

DOCS = "/home/local/ai-reference/aios-docs"
WIN_DIR = "/mnt/c/Users/Public/ithunder_shots/docs/toc"
WIN_DIR_W = r"C:\Users\Public\ithunder_shots\docs\toc"
SOFFICE = "/mnt/c/Program Files/LibreOffice/program/soffice.exe"
NOTE = ("Mục lục cập nhật ngày 20/11/2026. Nếu sửa tài liệu trong Word, bấm chuột phải vào mục lục "
        "rồi chọn Update Field để cập nhật số trang.")
FILES = ["iThunder_AI_OS_TL1_Huong_dan_su_dung_AI_Workspace", "iThunder_AI_OS_TL2_Cai_dat_va_Van_hanh_AI_Platform"]
W = nsdecls("w")


def text_of(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def headings(d):
    out = []
    for p in d.paragraphs:
        name = p.style.name if p.style is not None else ""
        if name in ("Heading 1", "Heading 2") and p.text.strip():
            out.append((1 if name == "Heading 1" else 2, p))
    return out


def rpr(level):
    if level == 1:
        return '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/><w:b/><w:color w:val="1F3A5F"/><w:sz w:val="22"/></w:rPr>'
    return '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/><w:color w:val="1F2937"/><w:sz w:val="21"/></w:rPr>'


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def entry(level, title, bm, page, first=False, toc_instr=""):
    rp = rpr(level)
    head = ""
    if first:
        head = (f'<w:r>{rp}<w:fldChar w:fldCharType="begin"/></w:r>'
                f'<w:r>{rp}<w:instrText xml:space="preserve">{esc(toc_instr)}</w:instrText></w:r>'
                f'<w:r>{rp}<w:fldChar w:fldCharType="separate"/></w:r>')
    before = 140 if level == 1 else 0
    ind = 0 if level == 1 else 400
    return parse_xml(
        f'<w:p {W}><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs>'
        f'<w:spacing w:before="{before}" w:after="50"/><w:ind w:left="{ind}" w:right="500" w:hanging="0"/></w:pPr>'
        f'{head}<w:hyperlink w:anchor="{bm}" w:history="1">'
        f'<w:r>{rp}<w:t xml:space="preserve">{esc(title)}</w:t></w:r>'
        f'<w:r>{rp}<w:tab/></w:r>'
        f'<w:r>{rp}<w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r>{rp}<w:instrText xml:space="preserve"> PAGEREF {bm} \\h </w:instrText></w:r>'
        f'<w:r>{rp}<w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r>{rp}<w:t>{page}</w:t></w:r>'
        f'<w:r>{rp}<w:fldChar w:fldCharType="end"/></w:r>'
        f'</w:hyperlink></w:p>')


def build(src, dst, pages):
    d = docx.Document(src)
    body = d.element.body
    toc_p = next(p for p in body.iter(qn("w:p")) if any("TOC " in (i.text or "") for i in p.iter(qn("w:instrText"))))
    toc_instr = next(i.text for i in toc_p.iter(qn("w:instrText")) if "TOC " in (i.text or ""))
    note_p = toc_p.getnext()
    assert "Update Field" in text_of(note_p), text_of(note_p)
    ids = [int(b.get(qn("w:id"))) for b in body.iter(qn("w:bookmarkStart"))]
    next_id = max(ids, default=0) + 1
    hs = headings(d)
    new = []
    for n, (level, p) in enumerate(hs):
        bm = f"_Toc{8000000 + n:08d}"
        el = p._p
        start = parse_xml(f'<w:bookmarkStart {W} w:id="{next_id + n}" w:name="{bm}"/>')
        end = parse_xml(f'<w:bookmarkEnd {W} w:id="{next_id + n}"/>')
        ppr = el.find(qn("w:pPr"))
        (ppr.addnext(start) if ppr is not None else el.insert(0, start))
        el.append(end)
        new.append(entry(level, p.text.strip(), bm, pages[n] if pages else "00", first=(n == 0), toc_instr=toc_instr))
    # thay đoạn TOC rỗng bằng các dòng mục lục; đoạn cuối đóng trường
    anchor = toc_p
    for e in new:
        anchor.addnext(e)
        anchor = e
    closing = parse_xml(f'<w:p {W}><w:pPr><w:spacing w:after="120"/></w:pPr><w:r>{rpr(2)}<w:fldChar w:fldCharType="end"/></w:r></w:p>')
    anchor.addnext(closing)
    body.remove(toc_p)
    for t in note_p.iter(qn("w:t")):
        t.text = ""
    note_p.findall(".//" + qn("w:t"))[0].text = NOTE
    d.save(dst)
    return [(lv, p.text.strip()) for lv, p in hs]


def to_pdf(path):
    os.makedirs(WIN_DIR, exist_ok=True)
    name = os.path.basename(path)
    shutil.copy(path, os.path.join(WIN_DIR, name))
    subprocess.run([SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", WIN_DIR_W, WIN_DIR_W + "\\" + name],
                   check=True, capture_output=True, timeout=600)
    return os.path.join(WIN_DIR, name[:-5] + ".pdf")


def page_texts(pdf):
    n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout).group(1))
    out = []
    for i in range(1, n + 1):
        t = subprocess.run(["pdftotext", "-f", str(i), "-l", str(i), pdf, "-"], capture_output=True, text=True).stdout
        out.append([re.sub(r"\s+", " ", ln).strip() for ln in t.splitlines() if ln.strip()])
    return out


def find_pages(pdf, hs):
    pages = page_texts(pdf)
    first = hs[0][1]
    # trang thân bài đầu tiên: tiêu đề đầu tiên nằm ở đầu trang (Heading1 có ngắt trang trước)
    cur = next(i for i, lines in enumerate(pages) if i > 1 and lines and lines[0].startswith(first[:25]))
    res = []
    for _, title in hs:
        key = re.sub(r"\s+", " ", title)[:28]
        for i in range(cur, len(pages)):
            if any(ln.startswith(key) for ln in pages[i]):
                cur = i
                res.append(str(i + 1))
                break
        else:
            raise SystemExit(f"Không thấy tiêu đề trong PDF: {title}")
    return res, len(pages)


def main():
    work = os.path.join(os.path.dirname(os.path.abspath(__file__)), "toc_work")
    os.makedirs(work, exist_ok=True)
    # Tham số: TL1 / TL2 để chỉ dựng một tài liệu (đầu vào toc_work/<tên>.src.docx là bản chưa điền mục lục)
    files = [f for f in FILES if any(a in f for a in sys.argv[1:])] or FILES
    for f in files:
        src = os.path.join(work, f + ".src.docx")
        if not os.path.exists(src):
            shutil.copy(os.path.join(DOCS, f + ".docx"), src)   # bản trước khi điền mục lục
        tmp = os.path.join(work, f + ".docx")
        hs = build(src, tmp, None)
        pages, _ = find_pages(to_pdf(tmp), hs)
        for _ in range(3):
            build(src, tmp, pages)
            again, total = find_pages(to_pdf(tmp), hs)
            if again == pages:
                break
            pages = again
        else:
            raise SystemExit("Số trang không ổn định")
        shutil.copy(tmp, os.path.join(DOCS, f + ".docx"))
        shutil.copy(os.path.join(WIN_DIR, f + ".pdf"), os.path.join(DOCS, f + ".pdf"))
        print(f, f"{len(hs)} mục, {total} trang")
        for (lv, t), pg in zip(hs, pages):
            print("   " * (lv - 1), t, "....", pg)


if __name__ == "__main__":
    sys.exit(main())
