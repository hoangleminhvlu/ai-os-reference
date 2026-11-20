# -*- coding: utf-8 -*-
"""Hàm dùng chung để dựng bản mới của TL1, TL2 từ bản trước (tách từ update_v011.py, update_tl2_v011.py, update_v013.py
ngày 20/11/2026; các script dựng bản 0.11 tới 0.13 chứa nội dung Phụ lục G nội bộ nên chỉ lưu ngoài kho công khai).
Script dựng bản sau (update_v014.py trở đi) chỉ nhập hàm từ đây và từ docedit.py, add_shots.py."""
import copy
import os
import sys

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.table import Table
from docx.text.paragraph import Paragraph

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def strip_bookmarks(el):
    for tag in ("w:bookmarkStart", "w:bookmarkEnd"):
        for b in list(el.iter(qn(tag))):
            b.getparent().remove(b)
    return el


def clone_heading(doc, template_text, anchor, text):
    el = strip_bookmarks(copy.deepcopy(doc.heading(template_text)))
    anchor.addnext(el)
    doc.set_para(el, text, note=f"Thêm tiêu đề: {text}")
    return el


def clone_table(doc, needle, anchor, keep_rows):
    """Nhân bản bảng có sẵn (giữ định dạng), chỉ giữ keep_rows hàng đầu (gồm hàng tiêu đề)."""
    el = copy.deepcopy(doc.find(needle, "tbl"))
    anchor.addnext(el)
    for tr in el.findall(qn("w:tr"))[keep_rows:]:
        el.remove(tr)
    return el, Table(el, doc.d)


def blank_para(doc):
    """Đoạn trống kiểu Normal làm mẫu khoảng cách (sau bảng, sau hộp)."""
    for el in doc.elements():
        if el.tag == qn("w:p") and doc.text_of(el) == "" and el.find(".//" + qn("w:drawing")) is None \
                and el.find(".//" + qn("w:sectPr")) is None and Paragraph(el, doc.d).style.name == "Normal" \
                and el.find(".//" + qn("w:fldChar")) is None:
            return el
    raise SystemExit("Không có đoạn trống mẫu")


def code_after(doc, anchor, lines, _tpl={}):
    if doc not in _tpl:
        _tpl[doc] = copy.deepcopy(doc.find("cd ai-work", "tbl"))   # khối mã mẫu (mục 4.3), lấy trước khi nhân bản
    el = copy.deepcopy(_tpl[doc])
    anchor.addnext(el)
    doc.set_code(Table(el, doc.d), lines)
    return el


def note_plain(doc, anchor, kind, markup, blank_template):
    """Hộp kind với thân chữ Arial thường: hộp mẫu đầu tiên của TL2 bắt đầu bằng mã nguồn (Consolas), nên lấy định dạng
    thân hộp từ hộp Lưu ý, giữ chữ đậm."""
    last = doc.note_after(anchor, kind, markup, blank_template=blank_template)
    box = last.getprevious()
    ref = next(e for e in doc.elements() if e.tag == qn("w:tbl") and doc.text_of(e).startswith("Lưu ý."))
    body = copy.deepcopy(ref.find(".//" + qn("w:p")).findall(qn("w:r"))[1].find(qn("w:rPr")))
    for tag in ("w:b", "w:bCs"):
        x = body.find(qn(tag))
        if x is not None:
            body.remove(x)
    for r in box.find(".//" + qn("w:p")).findall(qn("w:r"))[1:]:
        old = r.find(qn("w:rPr"))
        new = copy.deepcopy(body)
        if old is not None:
            if is_bold(old):
                new.append(docx.oxml.OxmlElement("w:b"))
            r.remove(old)
        r.insert(0, new)
    return last


def unfill_toc(doc):
    """Gỡ kết quả mục lục do build_toc.py điền (các dòng mục lục, bookmark _Toc8…) để dựng lại sau khi sửa nội dung."""
    body = doc.body
    first = next(p for p in body.iter(qn("w:p")) if any("TOC " in (i.text or "") for i in p.iter(qn("w:instrText"))))
    instr = next(i.text for i in first.iter(qn("w:instrText")) if "TOC " in (i.text or ""))
    el, removed = first, []
    while True:
        removed.append(el)
        chars = [f.get(qn("w:fldCharType")) for f in el.iter(qn("w:fldChar"))]
        if el is not first and chars == ["end"] and not doc.text_of(el).strip():
            break
        el = el.getnext()
    W = nsdecls("w")
    toc = parse_xml(f'<w:p {W}><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">'
                    f'{instr}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                    f'<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
    first.addprevious(toc)
    for e in removed:
        body.remove(e)
    names = {}
    for b in list(body.iter(qn("w:bookmarkStart"))):
        if (b.get(qn("w:name")) or "").startswith("_Toc8"):
            names[b.get(qn("w:id"))] = True
            b.getparent().remove(b)
    for b in list(body.iter(qn("w:bookmarkEnd"))):
        if b.get(qn("w:id")) in names:
            b.getparent().remove(b)
    assert "Update Field" in doc.text_of(toc.getnext())


def new_section(doc, prev_heading, title, intro, intro_template, table_needle, rows):
    """'Những điểm mới trong bản …': đoạn dẫn và bảng 2 cột nhân bản từ bảng của bản trước (thêm hàng khi cần)."""
    h_prev = doc.heading(prev_heading)
    h = clone_heading(doc, prev_heading, h_prev.getprevious(), title)
    p = doc.para_after(h, intro_template, intro)
    tbl = copy.deepcopy(doc.find(table_needle, "tbl", nth=0))
    p.addnext(tbl)
    trs = tbl.findall(qn("w:tr"))
    for tr in trs[1 + len(rows):]:
        tbl.remove(tr)
    while len(tbl.findall(qn("w:tr"))) < 1 + len(rows):
        tbl.append(copy.deepcopy(tbl.findall(qn("w:tr"))[-1]))
    t = Table(tbl, doc.d)
    for r, (a, b) in enumerate(rows, start=1):
        doc.set_cell(t, r, 0, a, code=False)
        doc.set_cell(t, r, 1, b, code=False)
    tbl.addnext(copy.deepcopy(blank_para(doc)))


def replace_picture(doc, caption_needle, path):
    """Thay ảnh của đoạn ngay trước chú thích (giữ kích thước, vị trí trong tài liệu)."""
    cap = doc.find(caption_needle, "p")
    blip = cap.getprevious().find(".//" + qn("a:blip"))
    part = doc.d.part.related_parts[blip.get(qn("r:embed"))]
    with open(path, "rb") as f:
        part._blob = f.read()
    doc.changes.append(f"Thay ảnh: {caption_needle}")
    return cap


def set_caption(doc, cap, title):
    runs = cap.findall(qn("w:r"))
    runs[1].find(qn("w:t")).text = title
    for r in runs[2:]:
        cap.remove(r)


def last_shot_number(doc):
    import re
    nums = [int(m.group(1)) for e in doc.elements() for m in [re.match(r"Ảnh màn hình (\d+)\. ", doc.text_of(e))] if m]
    return max(nums)


def insert_picture(doc, anchor, img_path, title, width_cm, caption_needle):
    """Ảnh + chú thích "Ảnh màn hình 99. <title>" sau anchor, theo mẫu ảnh có chú thích caption_needle; chạy
    renumber_shots sau khi chèn để đánh số theo thứ tự trong tài liệu."""
    from docx.shared import Cm
    cap_tpl = doc.find(caption_needle, "p")
    pic_tpl = cap_tpl.getprevious()
    p = doc.d.add_paragraph()
    p.add_run().add_picture(img_path, width=Cm(width_cm))
    el = p._p
    old = el.find(qn("w:pPr"))
    if old is not None:
        el.remove(old)
    el.insert(0, copy.deepcopy(pic_tpl.find(qn("w:pPr"))))
    anchor.addnext(el)
    cap = copy.deepcopy(cap_tpl)
    runs = cap.findall(qn("w:r"))
    runs[0].find(qn("w:t")).text = "Ảnh màn hình 99. "
    runs[1].find(qn("w:t")).text = title
    for r in runs[2:]:
        cap.remove(r)
    el.addnext(cap)
    doc.changes.append(f"Thêm ảnh màn hình: {title[:60]}")
    return cap


def renumber_shots(doc):
    """Đánh lại số "Ảnh màn hình N" theo thứ tự trong tài liệu và sửa chỗ dẫn chiếu trong câu. Trả {số cũ: số mới}."""
    import re
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
                t.text = re.sub(r"Ảnh màn hình (\d+)(?=[,) .])",
                                lambda m: f"Ảnh màn hình {mapping.get(m.group(1), m.group(1))}", t.text)
    doc.changes.append(f"Đánh lại số ảnh màn hình: {mapping}")
    return mapping
