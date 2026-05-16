from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# ── ページ設定 ──────────────────────────────────────────────
for sec in doc.sections:
    sec.page_width    = Cm(21)
    sec.page_height   = Cm(29.7)
    sec.left_margin   = Cm(1.5)
    sec.right_margin  = Cm(1.5)
    sec.top_margin    = Cm(1.5)
    sec.bottom_margin = Cm(1.5)

W = 18.0  # 本文幅 (cm)

# ── ユーティリティ ─────────────────────────────────────────

def set_row_h(row, h_cm):
    trPr = row._tr.get_or_add_trPr()
    e = OxmlElement('w:trHeight')
    e.set(qn('w:val'), str(int(h_cm * 567)))
    e.set(qn('w:hRule'), 'exact')
    trPr.append(e)

def jp_run(run, size, bold=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = 'MS Gothic'
    rpr = run._element.get_or_add_rPr()
    rpr.get_or_add_rFonts().set(qn('w:eastAsia'), 'MS Gothic')

def cell_write(cell, text='', size=9, bold=False,
               align=WD_ALIGN_PARAGRAPH.CENTER,
               valign=WD_ALIGN_VERTICAL.CENTER):
    cell.vertical_alignment = valign
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    r = p.add_run(text)
    jp_run(r, size, bold)

def label(doc, text, size=9, bold=True, before=4, after=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after  = Pt(after)
    r = p.add_run(text)
    jp_run(r, size, bold)

def sub_label(doc, text, size=8.5, bold=False, before=2, after=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after  = Pt(after)
    r = p.add_run(text)
    jp_run(r, size, bold)

def no_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    b = OxmlElement('w:tcBorders')
    for side in ('top','left','bottom','right','insideH','insideV'):
        e = OxmlElement(f'w:{side}')
        e.set(qn('w:val'), 'none')
        e.set(qn('w:sz'), '0')
        e.set(qn('w:space'), '0')
        e.set(qn('w:color'), 'auto')
        b.append(e)
    tcPr.append(b)

def shade(cell, fill='D9D9D9'):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    s = OxmlElement('w:shd')
    s.set(qn('w:val'), 'clear')
    s.set(qn('w:color'), 'auto')
    s.set(qn('w:fill'), fill)
    tcPr.append(s)

def page_break(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    p.add_run()._element.append(br)

def new_table(rows, cols):
    t = doc.add_table(rows=rows, cols=cols)
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    return t

# ─────────────────────────────────────────────────────────────
# ▌ ヘッダー
# ─────────────────────────────────────────────────────────────
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after  = Pt(3)
r = p.add_run('2026年度　山形県立山形南高等学校　２学年１学期中間テスト　論理・表現Ⅱ　解答用紙')
jp_run(r, 11, True)

t_hd = new_table(2, 4)
for i, row in enumerate(t_hd.rows):
    set_row_h(row, 0.65)
for j, w in enumerate([Cm(2.2), Cm(3.0), Cm(2.2), Cm(10.6)]):
    for i in range(2): t_hd.cell(i, j).width = w
cell_write(t_hd.cell(0,0), 'Class', 9, True)
cell_write(t_hd.cell(0,2), 'No',    9, True)
cell_write(t_hd.cell(1,0), 'Name',  9, True)
t_hd.cell(1,1).merge(t_hd.cell(1,3))

# ─────────────────────────────────────────────────────────────
# ▌ １　Listening（思考・判断・表現　14点）
# ─────────────────────────────────────────────────────────────
label(doc, '１　（思考・判断・表現　14点）', before=5)

# A(2点) / B(2点) 横並び
t1ab = new_table(2, 6)
for j, w in enumerate([Cm(2.0), Cm(2.0), Cm(0.4), Cm(2.0), Cm(2.0), Cm(9.6)]):
    for i in range(2): t1ab.cell(i, j).width = w
set_row_h(t1ab.rows[0], 0.44)
set_row_h(t1ab.rows[1], 0.68)
cell_write(t1ab.cell(0,0), 'A（2点）', 8, True); shade(t1ab.cell(0,0))
cell_write(t1ab.cell(1,0), '答え', 7.5)
cell_write(t1ab.cell(0,1), ''); cell_write(t1ab.cell(1,1), '')
no_border(t1ab.cell(0,2)); no_border(t1ab.cell(1,2))
cell_write(t1ab.cell(0,3), 'B（2点）', 8, True); shade(t1ab.cell(0,3))
cell_write(t1ab.cell(1,3), '答え', 7.5)
cell_write(t1ab.cell(0,4), ''); cell_write(t1ab.cell(1,4), '')
no_border(t1ab.cell(0,5)); no_border(t1ab.cell(1,5))

# C(10点) — 各問の空欄数 (1)3語 (2)2語 (3)1語 (4)3語 (5)3語
sub_label(doc, 'C（2×5=10点）')
C_blanks = [3, 2, 1, 3, 3]
# 列：問ごとに「番号1列＋解答n列」、計 5*(1+3)=20列
total_c = 20
cw = Cm(W / total_c)
t1c = new_table(2, total_c)
for j in range(total_c):
    for i in range(2): t1c.cell(i,j).width = cw
set_row_h(t1c.rows[0], 0.44)
set_row_h(t1c.rows[1], 0.68)
col = 0
for qi, nb in enumerate(C_blanks):
    cell_write(t1c.cell(0, col), f'({qi+1})', 8, True); shade(t1c.cell(0, col))
    no_border(t1c.cell(1, col))
    col += 1
    for b in range(3):
        if b < nb:
            cell_write(t1c.cell(0, col), '', 7); cell_write(t1c.cell(1, col), '')
        else:
            no_border(t1c.cell(0, col)); no_border(t1c.cell(1, col))
        col += 1

# ─────────────────────────────────────────────────────────────
# ▌ ２　LEAP Vocabulary（知識・技能　15点）
# ─────────────────────────────────────────────────────────────
label(doc, '２　（知識・技能　15点）')

# A 10問
sub_label(doc, 'A（1×10=10点）')
t2a = new_table(2, 10)
for j in range(10):
    for i in range(2): t2a.cell(i,j).width = Cm(W/10)
set_row_h(t2a.rows[0], 0.44); set_row_h(t2a.rows[1], 0.68)
for j in range(10):
    cell_write(t2a.cell(0,j), f'({j+1})', 8, True); shade(t2a.cell(0,j))
    cell_write(t2a.cell(1,j), '')

# B 5問
sub_label(doc, 'B（1×5=5点）')
t2b = new_table(2, 5)
for j in range(5):
    for i in range(2): t2b.cell(i,j).width = Cm(W/5)
set_row_h(t2b.rows[0], 0.44); set_row_h(t2b.rows[1], 0.68)
for j in range(5):
    cell_write(t2b.cell(0,j), f'({j+1})', 8, True); shade(t2b.cell(0,j))
    cell_write(t2b.cell(1,j), '')

# ─────────────────────────────────────────────────────────────
# ▌ ３　Grammar and Usage（知識・技能　28点）
# ─────────────────────────────────────────────────────────────
label(doc, '３　（知識・技能　28点）')

# A 語句整序 10問 — 5行×2列
sub_label(doc, 'A（1×10=10点）　※完全正答のみ得点')
nw = Cm(0.7); hw = Cm(W/2 - 0.7)
t3a = new_table(5, 4)
for i in range(5):
    set_row_h(t3a.rows[i], 0.70)
    for j, w in enumerate([nw, hw, nw, hw]):
        t3a.cell(i, j).width = w
for i in range(5):
    for s in range(2):
        q = i*2 + s + 1
        cell_write(t3a.cell(i, s*2),   f'({q})', 8, True); shade(t3a.cell(i, s*2))
        cell_write(t3a.cell(i, s*2+1), '', 9, align=WD_ALIGN_PARAGRAPH.LEFT)

# B 空所補充 12問（各問の空欄数）
sub_label(doc, 'B（1×12=12点）')
b_blanks = [1, 2, 2, 3, 3, 3,   1, 2, 2, 2, 2, 1]
# 左Q1-6 / 右Q7-12 の2列レイアウト
# 各半分：番号0.6cm + 解答ボックス最大3つ
bx_w = Cm((W/2 - 0.6) / 3)
t3b = new_table(6, 8)
b_widths = [Cm(0.6), bx_w, bx_w, bx_w, Cm(0.6), bx_w, bx_w, bx_w]
for j, w in enumerate(b_widths):
    for i in range(6): t3b.cell(i, j).width = w
for i in range(6): set_row_h(t3b.rows[i], 0.68)

for i in range(6):
    for side in range(2):
        qi  = i + side*6
        bc  = side*4
        nb  = b_blanks[qi]
        cell_write(t3b.cell(i, bc), f'({qi+1})', 8, True); shade(t3b.cell(i, bc))
        for b in range(3):
            c = bc + 1 + b
            if b < nb:
                cell_write(t3b.cell(i, c), '', 9)
            else:
                no_border(t3b.cell(i, c))

# C 動詞活用 6問
sub_label(doc, 'C（1×6=6点）')
t3c = new_table(2, 6)
for j in range(6):
    for i in range(2): t3c.cell(i,j).width = Cm(W/6)
set_row_h(t3c.rows[0], 0.44); set_row_h(t3c.rows[1], 0.68)
for j in range(6):
    cell_write(t3c.cell(0,j), f'({j+1})', 8, True); shade(t3c.cell(0,j))
    cell_write(t3c.cell(1,j), '')

# ════════════════════ ページ区切り ════════════════════
page_break(doc)

# ─────────────────────────────────────────────────────────────
# ▌ ４　Writing（思考・判断・表現　12点）英訳
# ─────────────────────────────────────────────────────────────
label(doc, '４　（思考・判断・表現　12点）', before=0)

t4 = new_table(6, 2)
for i in range(6):
    t4.cell(i,0).width = Cm(0.7)
    t4.cell(i,1).width = Cm(W - 0.7)
    set_row_h(t4.rows[i], 0.82)
    cell_write(t4.cell(i,0), f'({i+1})', 8, True); shade(t4.cell(i,0))
    cell_write(t4.cell(i,1), '', 9, align=WD_ALIGN_PARAGRAPH.LEFT)

# ─────────────────────────────────────────────────────────────
# ▌ ５　Reading Comprehension Unit1&2（思考・判断・表現　10点）
# ─────────────────────────────────────────────────────────────
label(doc, '５　（思考・判断・表現　10点）')

# A と B を2列並び
sub_label(doc, 'A（1×5=5点）　　　　　　　　　　　　　B（1×5=5点）', before=1)
nw5 = Cm(0.7); aw5 = Cm(W/2 - 0.7)
t5 = new_table(5, 4)
for i in range(5):
    set_row_h(t5.rows[i], 0.70)
    for s in range(2):
        t5.cell(i, s*2).width   = nw5
        t5.cell(i, s*2+1).width = aw5
        cell_write(t5.cell(i, s*2),   f'({i+1})', 8, True); shade(t5.cell(i, s*2))
        cell_write(t5.cell(i, s*2+1), '', 9, align=WD_ALIGN_PARAGRAPH.LEFT)

# ─────────────────────────────────────────────────────────────
# ▌ ６　Reading Comprehension: Ancient Greece（思考・判断・表現　15点）
# ─────────────────────────────────────────────────────────────
label(doc, '６　（思考・判断・表現　15点）')

# 問1(1点)・問2(2点)・問5(2点) — 横並び選択ボックス
sub_label(doc, '問1（1点）　問2（2点）　問5（2点）　　各選択（A・B・C・D）', before=1)
t6a = new_table(2, 6)
for j, w in enumerate([Cm(2.2), Cm(2.6), Cm(2.2), Cm(2.6), Cm(2.2), Cm(6.2)]):
    for i in range(2): t6a.cell(i,j).width = w
set_row_h(t6a.rows[0], 0.44); set_row_h(t6a.rows[1], 0.68)
for idx, (lbl, col) in enumerate([('問1', 0), ('問2', 2), ('問5', 4)]):
    cell_write(t6a.cell(0, col), lbl, 8, True); shade(t6a.cell(0, col))
    cell_write(t6a.cell(1, col), '答え', 7.5)
    cell_write(t6a.cell(0, col+1), ''); cell_write(t6a.cell(1, col+1), '')
no_border(t6a.cell(0,5)); no_border(t6a.cell(1,5))

# 問3(5点) — (i)〜(v) S/A
sub_label(doc, '問3（1×5=5点）　　SはスパルタのS、AはアテネのA', before=2)
t6b = new_table(2, 5)
for j in range(5):
    for i in range(2): t6b.cell(i,j).width = Cm(W/5)
set_row_h(t6b.rows[0], 0.44); set_row_h(t6b.rows[1], 0.68)
for j, lbl in enumerate(['(i)', '(ii)', '(iii)', '(iv)', '(v)']):
    cell_write(t6b.cell(0,j), lbl, 8, True); shade(t6b.cell(0,j))
    cell_write(t6b.cell(1,j), '')

# 問4(2点) — 日本語記述
sub_label(doc, '問4（2点）　スパルタの女性がアテネの女性より「自由」を得られた理由（日本語）', before=2)
t6c = new_table(2, 1)
for i in range(2):
    t6c.cell(i,0).width = Cm(W)
    set_row_h(t6c.rows[i], 0.80)
    cell_write(t6c.cell(i,0), '', 9)

# 問6(3点) — 日本語訳
sub_label(doc, '問6（3点）　下線部(2)を日本語にしなさい', before=2)
t6d = new_table(2, 1)
for i in range(2):
    t6d.cell(i,0).width = Cm(W)
    set_row_h(t6d.rows[i], 0.80)
    cell_write(t6d.cell(i,0), '', 9)

# ─────────────────────────────────────────────────────────────
# ▌ ７　Writing: Book Report（思考・判断・表現　6点）
# ─────────────────────────────────────────────────────────────
label(doc, '７　（思考・判断・表現　6点）　50語以上で書くこと')

t7 = new_table(6, 1)
for i in range(6):
    t7.cell(i,0).width = Cm(W)
    set_row_h(t7.rows[i], 0.82)
    cell_write(t7.cell(i,0), '', 9)

# ─────────────────────────────────────────────────────────────
# ▌ 得点集計
# ─────────────────────────────────────────────────────────────
label(doc, '得点集計', before=5)

ts = new_table(2, 4)
s_labels = ['１〜３（知識・技能）　/43', '１・４〜７（思考・判断・表現）　/57', '合計　/100', '']
s_widths  = [Cm(5.5), Cm(6.5), Cm(4.0), Cm(2.0)]
for j, (lbl, w) in enumerate(zip(s_labels, s_widths)):
    for i in range(2): ts.cell(i,j).width = w
set_row_h(ts.rows[0], 0.44); set_row_h(ts.rows[1], 0.78)
for j, lbl in enumerate(s_labels[:3]):
    cell_write(ts.cell(0,j), lbl, 8, True); shade(ts.cell(0,j))
    cell_write(ts.cell(1,j), '')
no_border(ts.cell(0,3)); no_border(ts.cell(1,3))

# ─── 保存 ──────────────────────────────────────────────────
out = '/home/user/Yumetan-Quick-Response/2026年度_論理表現Ⅱ_中間テスト_解答用紙.docx'
doc.save(out)
print('Saved:', out)
