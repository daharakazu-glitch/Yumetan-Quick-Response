"""
2026年度 論理・表現Ⅱ 中間テスト 解答例
ユーザーが最終確定した解答用紙 DOCX をベースに、赤色 Arial フォントで解答を挿入する。
レイアウトは一切変更しない。
"""
import copy
import shutil
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from lxml import etree

SRC = '/root/.claude/uploads/1fe2160d-2e19-4c6d-bdf0-ff50f20ed0c7/9d049440-2026____________________Complete.docx'
OUT = '/home/user/Yumetan-Quick-Response/2026年度_論理表現Ⅱ_中間テスト_解答例.docx'

RED   = RGBColor(0xFF, 0x00, 0x00)
BLACK = RGBColor(0x00, 0x00, 0x00)
FS    = 10  # font size pt

# ── Font helper ──────────────────────────────────────────────────────────────
def font_xml(run, size=FS, color=None, bold=False):
    r = run._r
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        r.append(rPr)

    # fonts
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'),   'Arial')
    rFonts.set(qn('w:hAnsi'),   'Arial')
    rFonts.set(qn('w:eastAsia'), 'MS PGothic')

    # bold
    b_el = rPr.find(qn('w:b'))
    if bold:
        if b_el is None:
            b_el = OxmlElement('w:b'); rPr.append(b_el)
    else:
        if b_el is not None: rPr.remove(b_el)

    # size
    for tag_local in ['w:sz', 'w:szCs']:
        el = rPr.find(qn(tag_local))
        if el is None:
            el = OxmlElement(tag_local)
            rPr.append(el)
        el.set(qn('w:val'), str(int(size * 2)))

    # color
    if color is not None:
        c_el = rPr.find(qn('w:color'))
        if c_el is None:
            c_el = OxmlElement('w:color'); rPr.append(c_el)
        c_el.set(qn('w:val'), f'{color[0]:02X}{color[1]:02X}{color[2]:02X}')

def add_run(para, text, size=FS, color=None, bold=False):
    """Add a run to paragraph with Arial font."""
    run = para.add_run(text)
    col = color if color else BLACK
    font_xml(run, size, col, bold)
    return run

# ── Cell helpers ──────────────────────────────────────────────────────────────
def clear_para(para):
    """Remove all runs from a paragraph while keeping paragraph properties."""
    p = para._p
    for r in p.findall(qn('w:r')):
        p.remove(r)
    for ins in p.findall(qn('w:ins')):
        p.remove(ins)

def get_raw_tcs(row):
    """Return list of raw TC elements (actual cells, not python-docx merged repeats)."""
    return row._tr.findall('.//' + qn('w:tc'))

def cell_from_tc(tc):
    """Wrap a raw TC element as a python-docx Cell-like object to access paragraphs."""
    from docx.table import _Cell
    # We need the parent table; use a workaround
    return tc

def set_cell_answer(cell, text, size=FS, color=(0xFF, 0x00, 0x00)):
    """Clear cell paragraphs and write answer text in red."""
    para = cell.paragraphs[0]
    clear_para(para)
    add_run(para, text, size, color)

def append_answer(cell, text, size=FS, color=(0xFF, 0x00, 0x00)):
    """Append answer text to existing cell content (keeps label text)."""
    para = cell.paragraphs[0]
    add_run(para, text, size, color)

def set_cell_mixed(cell, pre_text, answer, post_text, size=FS):
    """Set cell content with black pre_text + red answer + black post_text."""
    para = cell.paragraphs[0]
    clear_para(para)
    if pre_text:
        add_run(para, pre_text, size, BLACK)
    add_run(para, answer, size, RED)
    if post_text:
        add_run(para, post_text, size, BLACK)

# ── Raw TC cell wrapper ───────────────────────────────────────────────────────
class TCWrapper:
    """Minimal wrapper around a raw w:tc element to use set_cell_* helpers."""
    def __init__(self, tc):
        self._tc = tc

    @property
    def paragraphs(self):
        from docx.text.paragraph import Paragraph
        ps = self._tc.findall('.//' + qn('w:p'))
        return [Paragraph(p, None) for p in ps]


def build(src, out):
    doc = Document(src)
    tables = doc.tables

    # ──────────────────────────────────────────────────────────────────────────
    # §1 A / B  (Table 1, 1r × 4c)
    # col1 = A answer, col3 = B answer
    # ──────────────────────────────────────────────────────────────────────────
    t1 = tables[1]
    set_cell_answer(t1.rows[0].cells[1], '②')
    set_cell_answer(t1.rows[0].cells[3], '①')

    # ──────────────────────────────────────────────────────────────────────────
    # §1 C  (Table 2, 3r × 5c with varying gridAfter)
    # Row 0: TC0=(1)label, TC1=ans(1), TC2=spacer, TC3=(2)label, TC4=ans(2)
    # Row 1: TC0=(3)label, TC1=(4)label, TC2=ans(3), TC3=ans(4), gridAfter=1
    # Row 2: TC0=(5)label, TC1=ans(5), TC2=spacer, gridAfter=2
    # ──────────────────────────────────────────────────────────────────────────
    t2 = tables[2]
    r2_tcs = [get_raw_tcs(row) for row in t2.rows]
    # Row 0
    set_cell_answer(TCWrapper(r2_tcs[0][1]), 'have been launched')
    set_cell_answer(TCWrapper(r2_tcs[0][4]), 'no longer')
    # Row 1
    set_cell_answer(TCWrapper(r2_tcs[1][2]), 'includes')
    set_cell_answer(TCWrapper(r2_tcs[1][3]), 'affect')
    # Row 2
    set_cell_answer(TCWrapper(r2_tcs[2][1]), 'in order to')

    # ──────────────────────────────────────────────────────────────────────────
    # §2A  (Table 3, 2r × 5c)
    # Each cell = "N." label + answer appended
    # ──────────────────────────────────────────────────────────────────────────
    t3 = tables[3]
    s2a = ['A', 'C', 'D', 'C', 'C', 'A', 'C', 'A', 'B', 'C']
    for ri in range(2):
        for ci in range(5):
            idx = ri * 5 + ci
            append_answer(t3.rows[ri].cells[ci], s2a[idx])

    # ──────────────────────────────────────────────────────────────────────────
    # §2B  (Table 4, 1r × 5c)
    # ──────────────────────────────────────────────────────────────────────────
    t4 = tables[4]
    s2b = ['purchase', 'suspect', 'permitted', 'miserable', 'superstition']
    for ci, ans in enumerate(s2b):
        append_answer(t4.rows[0].cells[ci], ans)

    # ──────────────────────────────────────────────────────────────────────────
    # §3A  (Table 5, 10r × 2c)
    # cell[ri, 1] has pre-printed text; we reconstruct with answer in red
    # ──────────────────────────────────────────────────────────────────────────
    t5 = tables[5]
    # (pre_text, answer, post_text)
    s3a = [
        ('', 'A good number of people visit the country', ' annually.'),
        ('', 'He would rather not see this movie', '.'),
        ('It ', 'is easy to live in Iceland', ' in summer.'),
        ('It is ', 'interesting that the country looks like a large cat', ' on the world map.'),
        ('', 'There are too many new places in the world to explore', '.'),
        ('The author ', 'was inspired by the 100-year-old tree', ' in her garden.'),
        ('It was quite frightening, but ', 'a little more horridness would have made the story better', '.'),
        ('In Japan, ', 'it is very important to entertain guests', '.'),
        ('Enjoy ', 'reading how these dinosaur fossils were found', '.'),
        ('', 'Nobody likes to be made fun of in', ' public.'),
    ]
    for ri, (pre, ans, post) in enumerate(s3a):
        cell = t5.rows[ri].cells[1]
        set_cell_mixed(cell, pre, ans, post)

    # ──────────────────────────────────────────────────────────────────────────
    # §3B  (Table 6, 6r × 15c with merged cells)
    # Access first-occurrence column index in cells() view
    # ──────────────────────────────────────────────────────────────────────────
    t6 = tables[6]

    def s3b_ans(row_idx, col_idx, text):
        cell = t6.rows[row_idx].cells[col_idx]
        set_cell_answer(cell, text)

    s3b_ans(0,  1, '(It)')                          # Q1
    s3b_ans(0,  3, '(It), (to)')                     # Q2
    s3b_ans(0,  9, '(renewable), (energy\'s)')        # Q3
    s3b_ans(1,  1, '(enabled)(her)(to)')              # Q4
    s3b_ans(1,  7, '(prevented)(from)(going)')        # Q5
    s3b_ans(2,  1, '(has)(been)(translated)')         # Q6
    s3b_ans(2,  7, '(contains)')                      # Q7
    s3b_ans(3,  1, '(located), (center)')             # Q8
    s3b_ans(3,  5, '(climate), (varies)')             # Q9
    s3b_ans(4,  1, '(water), (pollution)')            # Q10
    s3b_ans(4,  5, '(aging), (society)')              # Q11
    s3b_ans(4, 12, '(cutting-edge)')                  # Q12
    s3b_ans(5,  1, '(amazing)')                       # Q13

    # ──────────────────────────────────────────────────────────────────────────
    # §3C  (Table 7, 3r × 2c)
    # Each cell = number label + answer appended
    # ──────────────────────────────────────────────────────────────────────────
    t7 = tables[7]
    s3c = [
        ['was studying', 'are moving'],
        ['have been closed', 'kicks'],
        ['is getting', 'lived'],
    ]
    for ri in range(3):
        for ci in range(2):
            append_answer(t7.rows[ri].cells[ci], s3c[ri][ci])

    # ──────────────────────────────────────────────────────────────────────────
    # §4  (Table 8, 6r × 1c) — English translations
    # ──────────────────────────────────────────────────────────────────────────
    t8 = tables[8]
    s4 = [
        'My cell phone is not as new as hers.',
        'No other secretary is as capable as Yuki.',
        'What the man said was far from the truth.',
        'You will have a wonderful experience in New Zealand.',
        'The painting has been protected by bullet-proof glass since 1956.',
        'Many places are now being hit by power cuts.',
    ]
    for ri, ans in enumerate(s4):
        append_answer(t8.rows[ri].cells[0], ans)

    # ──────────────────────────────────────────────────────────────────────────
    # §5  (Table 9, 3r × 1c)
    # ──────────────────────────────────────────────────────────────────────────
    t9 = tables[9]
    s5 = [
        'Brunei is a small country in South-East Asia.',
        'There are around 100 mammals.',
        'We can see unique animals such as the proboscis monkey.',
    ]
    for ri, ans in enumerate(s5):
        append_answer(t9.rows[ri].cells[0], ans)

    # ──────────────────────────────────────────────────────────────────────────
    # §6 Q1/Q2  (Table 10, 1r × 5c)
    # TC0='1.(2点)', TC1=Q1 ans, TC2='2.(2x2=4点)', TC3=Q2 ans1, TC4=Q2 ans2
    # ──────────────────────────────────────────────────────────────────────────
    t10 = tables[10]
    raw_tcs10 = get_raw_tcs(t10.rows[0])
    set_cell_answer(TCWrapper(raw_tcs10[1]), 'B')           # Q1 answer
    set_cell_answer(TCWrapper(raw_tcs10[3]), '(i)')         # Q2 answer 1
    set_cell_answer(TCWrapper(raw_tcs10[4]), '(iv)')        # Q2 answer 2

    # ──────────────────────────────────────────────────────────────────────────
    # §6 Q3  (Table 11, 2r × 1c) — Japanese explanation (2 rows)
    # スパルタの女性がアテネの女性より「自由」を得られた理由
    # ──────────────────────────────────────────────────────────────────────────
    t11 = tables[11]
    q3_ans = [
        'スパルタでは夫が軍とともに家を離れて生活していたため、女性は家に一人で残り、',
        '多くの自由を得ることができたから。',
    ]
    for ri, text in enumerate(q3_ans):
        set_cell_answer(t11.rows[ri].cells[0], text)

    # ──────────────────────────────────────────────────────────────────────────
    # §6 Q4  (Table 12, 1r × 2c)
    # TC0='4.(2点)', TC1=Q4 ans
    # ──────────────────────────────────────────────────────────────────────────
    t12 = tables[12]
    raw_tcs12 = get_raw_tcs(t12.rows[0])
    set_cell_answer(TCWrapper(raw_tcs12[1]), 'A')           # Q4 answer

    # ──────────────────────────────────────────────────────────────────────────
    # §6 Q5  (Table 13, 2r × 1c) — Japanese translation (2 rows)
    # 下線部(2): "Sparta allowed the people in Athens to continue their way of
    # life as long as they did not attack other cities."
    # ──────────────────────────────────────────────────────────────────────────
    t13 = tables[13]
    q5_ans = [
        'スパルタはアテネの人々が他の都市を攻撃しない限り、',
        '自分たちのやり方で生活を続けることを許した。',
    ]
    for ri, text in enumerate(q5_ans):
        set_cell_answer(t13.rows[ri].cells[0], text)

    # ── Save ──────────────────────────────────────────────────────────────────
    doc.save(out)
    print(f'Saved: {out}')


if __name__ == '__main__':
    build(SRC, OUT)
