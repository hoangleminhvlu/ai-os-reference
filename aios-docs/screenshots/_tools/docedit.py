# -*- coding: utf-8 -*-
"""
Sửa tài liệu DOCX dựng sẵn (TL1, TL2) mà vẫn giữ định dạng gốc: nhân bản phần tử mẫu có trong tài liệu rồi thay nội dung.
Cú pháp nội dung: **chữ đậm** (tên nút, mục trên màn hình), [[ĐÃ CÓ]] / [[ĐANG PHÁT TRIỂN]] / [[ĐỊNH HƯỚNG]] (nhãn trạng thái),
"\n" trong đoạn = xuống dòng (w:br). Định vị theo nội dung chữ: không thấy đúng một phần tử thì dừng (không sửa nhầm).
"""
import copy
import re

import docx
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

LABELS = ("ĐÃ CÓ", "ĐANG PHÁT TRIỂN", "ĐỊNH HƯỚNG")
TOKEN = re.compile(r"(\*\*.+?\*\*|\[\[(?:ĐÃ CÓ|ĐANG PHÁT TRIỂN|ĐỊNH HƯỚNG)\]\])")


def is_bold(rpr) -> bool:
    """<w:b/> hoặc <w:b w:val="1|true"/>; <w:b w:val="0"/> là tắt đậm tường minh."""
    b = rpr.find(qn("w:b")) if rpr is not None else None
    return b is not None and (b.get(qn("w:val")) or "1").lower() not in ("0", "false", "off")


def set_bold(rpr):
    b = rpr.find(qn("w:b"))
    if b is None:
        rpr.append(docx.oxml.OxmlElement("w:b"))
    else:
        b.attrib.pop(qn("w:val"), None)


class Doc:
    def __init__(self, path):
        self.d = docx.Document(path)
        self.body = self.d.element.body
        self.label_rpr = {}
        for r in self.body.iter(qn("w:r")):
            t = "".join(x.text or "" for x in r.iter(qn("w:t"))).strip()
            if t in LABELS and t not in self.label_rpr and r.find(qn("w:rPr")) is not None:
                self.label_rpr[t] = copy.deepcopy(r.find(qn("w:rPr")))
        assert set(self.label_rpr) == set(LABELS), self.label_rpr.keys()
        # Định dạng chuẩn của chữ trong ô bảng (bảng nhiều cột, hàng dữ liệu): định dạng phổ biến nhất của chữ thường Arial
        # và của mã nguồn Consolas (không đậm, không phải nhãn trạng thái)
        from collections import Counter
        seen = {True: Counter(), False: Counter()}
        sample = {}
        for tbl in self.body.iter(qn("w:tbl")):
            rows = tbl.findall(qn("w:tr"))
            if len(rows) < 2 or len(rows[0].findall(qn("w:tc"))) < 2:
                continue
            for tr in rows[1:]:
                for r in tr.iter(qn("w:r")):
                    rpr = r.find(qn("w:rPr"))
                    if rpr is None or is_bold(rpr) or rpr.find(qn("w:shd")) is not None:
                        continue
                    color, sz = rpr.find(qn("w:color")), rpr.find(qn("w:sz"))
                    key = (sz.get(qn("w:val")) if sz is not None else None, color.get(qn("w:val")) if color is not None else None)
                    code = self._is_code(rpr)
                    seen[code][key] += 1
                    sample.setdefault((code, key), copy.deepcopy(rpr))
        self._code_rpr = sample[(True, seen[True].most_common(1)[0][0])] if seen[True] else None
        self._cell_text_rpr = sample[(False, seen[False].most_common(1)[0][0])] if seen[False] else None
        self.changes = []

    # ------------------------------------------------------------------ tìm kiếm
    @staticmethod
    def text_of(el):
        return "".join(x.text or "" for x in el.iter(qn("w:t")))

    def elements(self):
        return list(self.body.iterchildren())

    def find(self, needle, kind=None, start_after=None, nth=None):
        """Phần tử cấp thân tài liệu (đoạn 'p' hoặc bảng 'tbl') chứa needle; phải duy nhất (hoặc chọn nth)."""
        els = self.elements()
        if start_after is not None:
            els = els[els.index(start_after) + 1:]
        hits = [e for e in els if (kind is None or e.tag == qn("w:" + kind)) and needle in self.text_of(e)]
        if nth is not None:
            return hits[nth]
        if len(hits) != 1:
            raise SystemExit(f"Không tìm thấy duy nhất '{needle[:60]}' ({len(hits)} kết quả)")
        return hits[0]

    def para(self, needle, **kw):
        return Paragraph(self.find(needle, "p", **kw), self.d)

    def table(self, needle, **kw):
        return Table(self.find(needle, "tbl", **kw), self.d)

    def heading(self, text):
        for e in self.elements():
            if e.tag == qn("w:p") and Paragraph(e, self.d).style.name.startswith("Heading") and self.text_of(e).strip() == text:
                return e
        raise SystemExit(f"Không có tiêu đề '{text}'")

    # ------------------------------------------------------------------ viết nội dung
    def _fill(self, p_el, markup, base_rpr):
        """Xóa các run của đoạn rồi dựng lại từ markup, dùng base_rpr cho chữ thường."""
        for r in list(p_el.findall(qn("w:r"))) + list(p_el.findall(qn("w:hyperlink"))):
            p_el.remove(r)
        for part in TOKEN.split(markup):
            if not part:
                continue
            if part.startswith("[["):
                label = part[2:-2]
                self._run(p_el, f" {label} ", copy.deepcopy(self.label_rpr[label]))
            elif part.startswith("**"):
                rpr = copy.deepcopy(base_rpr) if base_rpr is not None else docx.oxml.OxmlElement("w:rPr")
                if not is_bold(rpr):
                    rpr.append(docx.oxml.OxmlElement("w:b"))
                self._run(p_el, part[2:-2], rpr)
            else:
                self._run(p_el, part, copy.deepcopy(base_rpr) if base_rpr is not None else None)

    @staticmethod
    def _run(p_el, text, rpr):
        lines = text.split("\n")
        for i, line in enumerate(lines):
            r = docx.oxml.OxmlElement("w:r")
            if rpr is not None:
                r.append(copy.deepcopy(rpr))
            if i:
                r.append(docx.oxml.OxmlElement("w:br"))
            if line:
                t = docx.oxml.OxmlElement("w:t")
                t.set(qn("xml:space"), "preserve")
                t.text = line
                r.append(t)
            p_el.append(r)

    @staticmethod
    def _is_code(rpr):
        fonts = rpr.find(qn("w:rFonts")) if rpr is not None else None
        return fonts is not None and "Consolas" in (fonts.get(qn("w:ascii")) or "")

    @staticmethod
    def _plain_rpr(p_el, allow_code=False):
        """rPr của run chữ thường đầu tiên (không đậm, không phải mã nguồn Consolas) trong đoạn, làm định dạng gốc."""
        runs = p_el.findall(qn("w:r"))
        for r in runs:
            rpr = r.find(qn("w:rPr"))
            if rpr is not None and not is_bold(rpr) and rpr.find(qn("w:shd")) is None \
                    and (allow_code or not Doc._is_code(rpr)):
                return copy.deepcopy(rpr)
        for r in runs:   # chỉ có run mã nguồn: dùng định dạng đó nhưng đổi font về Arial
            rpr = r.find(qn("w:rPr"))
            if rpr is not None and Doc._is_code(rpr) and not allow_code:
                rpr = copy.deepcopy(rpr)
                fonts = rpr.find(qn("w:rFonts"))
                for k in ("ascii", "hAnsi", "cs", "eastAsia"):
                    fonts.set(qn("w:" + k), "Arial")
                for tag in ("w:b", "w:shd", "w:color"):
                    x = rpr.find(qn(tag))
                    if x is not None:
                        rpr.remove(x)
                return rpr
        if runs and runs[0].find(qn("w:rPr")) is not None:
            rpr = copy.deepcopy(runs[0].find(qn("w:rPr")))
            for tag in ("w:b", "w:shd", "w:color"):
                x = rpr.find(qn(tag))
                if x is not None:
                    rpr.remove(x)
            return rpr
        return None

    def set_para(self, p_el, markup, note="", like=None):
        """like: đoạn mẫu lấy định dạng chữ thường (khi đoạn đích bắt đầu bằng mã nguồn)."""
        self._fill(p_el, markup, self._plain_rpr(like if like is not None else p_el))
        self.changes.append(note or f"Sửa đoạn: {markup[:70]}")
        return p_el

    def replace_in(self, el, old, new, note=""):
        """Thay chuỗi trong các w:t của phần tử (giữ định dạng run). old phải nằm trọn trong một run."""
        hits = [t for t in el.iter(qn("w:t")) if old in (t.text or "")]
        if len(hits) != 1:
            raise SystemExit(f"replace_in: '{old[:60]}' thấy {len(hits)} lần")
        hits[0].text = hits[0].text.replace(old, new)
        self.changes.append(note or f"Thay: {old[:50]} -> {new[:50]}")

    def clone_after(self, anchor, template, markup=None):
        el = copy.deepcopy(template)
        anchor.addnext(el)
        if markup is not None and el.tag == qn("w:p"):
            self._fill(el, markup, self._plain_rpr(template))
        return el

    def para_after(self, anchor, template_needle_or_el, markup):
        tpl = template_needle_or_el if not isinstance(template_needle_or_el, str) else self.find(template_needle_or_el, "p")
        el = self.clone_after(anchor, tpl, markup)
        self.changes.append(f"Thêm đoạn: {markup[:70]}")
        return el

    def step_after(self, anchor, template_step_el, n, markup):
        el = copy.deepcopy(template_step_el)
        anchor.addnext(el)
        runs = el.findall(qn("w:r"))
        head, brk = runs[0], runs[1]
        head.find(qn("w:t")).text = f"Bước {n}"
        base = self._plain_rpr(template_step_el)
        for r in runs[2:]:
            el.remove(r)
        tmp = docx.oxml.OxmlElement("w:p")
        self._fill(tmp, markup, base)
        for r in tmp.findall(qn("w:r")):
            el.append(r)
        self.changes.append(f"Thêm bước {n}: {markup[:60]}")
        return el

    def note_after(self, anchor, kind, markup, blank_template=None):
        """Hộp Lưu ý / Cảnh báo / Mẹo / Nguy hiểm: nhân bản hộp cùng loại có sẵn; thêm đoạn trống sau hộp như bố cục gốc."""
        tpl = None
        for e in self.elements():
            if e.tag == qn("w:tbl") and self.text_of(e).startswith(kind + "."):
                tpl = e
                break
        assert tpl is not None, kind
        box = copy.deepcopy(tpl)
        anchor.addnext(box)
        p = box.find(".//" + qn("w:p"))
        runs = p.findall(qn("w:r"))
        body_rpr = copy.deepcopy(runs[1].find(qn("w:rPr")))
        for r in runs[1:]:
            p.remove(r)
        tmp = docx.oxml.OxmlElement("w:p")
        self._fill(tmp, markup, body_rpr)
        for r in tmp.findall(qn("w:r")):
            p.append(r)
        last = box
        if blank_template is not None:
            last = copy.deepcopy(blank_template)
            box.addnext(last)
        self.changes.append(f"Thêm hộp {kind}: {markup[:60]}")
        return last

    def set_note(self, box_el, markup, note=""):
        """Viết lại thân hộp Lưu ý / Cảnh báo / Mẹo / Nguy hiểm, giữ run nhãn đầu hộp ("Mẹo. ")."""
        p = box_el.find(".//" + qn("w:p"))
        runs = p.findall(qn("w:r"))
        body_rpr = None
        for r in runs[1:]:
            rpr = r.find(qn("w:rPr"))
            if rpr is not None and not is_bold(rpr):
                body_rpr = copy.deepcopy(rpr)
                break
        if body_rpr is None:
            body_rpr = copy.deepcopy(runs[0].find(qn("w:rPr")))
            for tag in ("w:b", "w:color"):
                x = body_rpr.find(qn(tag))
                if x is not None:
                    body_rpr.remove(x)
        for r in runs[1:]:
            p.remove(r)
        tmp = docx.oxml.OxmlElement("w:p")
        self._fill(tmp, markup, body_rpr)
        for r in tmp.findall(qn("w:r")):
            p.append(r)
        self.changes.append(note or f"Sửa hộp: {markup[:60]}")

    # ------------------------------------------------------------------ bảng
    def set_cell(self, table, r, c, markup, note="", code=None):
        """code=True: chữ trong ô dùng font mã nguồn; False: chữ thường; None: theo định dạng sẵn có của ô."""
        cell = table.cell(r, c)
        paras = cell._tc.findall(qn("w:p"))
        if code is True and self._code_rpr is not None:
            base = copy.deepcopy(self._code_rpr)
        elif code is False and self._cell_text_rpr is not None:
            base = copy.deepcopy(self._cell_text_rpr)
        else:
            only_code = all(self._is_code(x) for x in cell._tc.iter(qn("w:rPr"))) and \
                any(True for _ in cell._tc.iter(qn("w:rPr")))
            base = self._plain_rpr(paras[0], allow_code=only_code)
        for extra in paras[1:]:
            cell._tc.remove(extra)
        self._fill(paras[0], markup, base)
        self.changes.append(note or f"Sửa ô: {markup[:70]}")

    def clone_row(self, table, like=-1):
        """Nhân bản nguyên vẹn một hàng (giữ định dạng) và đặt ở cuối bảng."""
        rows = table._tbl.findall(qn("w:tr"))
        new = copy.deepcopy(rows[like])
        rows[-1].addnext(new)
        return new

    def add_row(self, table, values, like=-1, after=None, note=""):
        """Thêm hàng (nhân bản hàng 'like'), đặt sau hàng 'after' (mặc định cuối bảng)."""
        rows = table._tbl.findall(qn("w:tr"))
        new = copy.deepcopy(rows[like])
        (rows[after] if after is not None else rows[-1]).addnext(new)
        idx = table._tbl.findall(qn("w:tr")).index(new)
        for c, v in enumerate(values):
            self.set_cell(table, idx, c, v, note="-")
        self.changes.append(note or f"Thêm hàng: {values[0][:60]}")
        return idx

    def del_row(self, table, r):
        tr = table._tbl.findall(qn("w:tr"))[r]
        table._tbl.remove(tr)
        self.changes.append("Xóa hàng")

    def find_row(self, table, needle):
        hits = [i for i, row in enumerate(table.rows) if needle in "".join(c.text for c in row.cells)]
        if len(hits) != 1:
            raise SystemExit(f"find_row '{needle[:50]}': {len(hits)} hàng")
        return hits[0]

    def set_code(self, table, lines):
        """Ô mã nguồn: mỗi dòng một đoạn (nhân bản đoạn đầu)."""
        tc = table.cell(0, 0)._tc
        paras = tc.findall(qn("w:p"))
        tpl = paras[0]
        for p in paras:
            tc.remove(p)
        for line in lines:
            p = copy.deepcopy(tpl)
            for r in p.findall(qn("w:r")):
                p.remove(r)
            self._run(p, line, copy.deepcopy(tpl.find(qn("w:r")).find(qn("w:rPr"))) if tpl.find(qn("w:r")) is not None else None)
            tc.append(p)
        self.changes.append(f"Sửa khối mã ({len(lines)} dòng)")

    def save(self, path):
        self.d.save(path)
