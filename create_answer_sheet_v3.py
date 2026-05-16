from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()
for sec in doc.sections:
    sec.page_width    = Cm(21)
    sec.page_height   = Cm(29.7)
    sec.left_margin   = Cm(1.5)
    sec.right_margin  = Cm(1.5)
    sec.top_margin    = Cm(1.5)
    sec.bottom_margin = Cm(1.5)

W = 18.0  # 本文幅 (cm)

# ──────────── ユーティリティ ────────────

def set_row_h(row, h_cm):
    trPr = row._tr.get_or_add_trPr()
    e = OxmlElement('w:trHeight')
    e.set(qn('w:val'), str(int(h_cm * 567)))
    e.set(qn('w:hRule'), 'exact')
    trPr.append(e)

def jp(run, size, bold=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = 'MS Gothic'
    rpr = run._element.get_or_add_rPr()
    rpr.get_or_add_rFonts().set(qn('w:eastAsia'), 'MS Gothic')

def cw(cell, text='', size=9, bold=False,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        valign=WD_ALIGN_VERTICAL.CENTER):
    cell.vertical_alignment = valign
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    r = p.add_run(text); jp(r, size, bold)

def lbl(text, size=9, bold=True, before=5, after=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after  = Pt(after)
    r = p.add_run(text); jp(r, size, bold)

def slbl(text, size=8.5, bold=False, before=2, after=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after  = Pt(after)
    r = p.add_run(text); jp(r, size, bold)

def no_b(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement('w:tcBorders')
    for s in ('top','left','bottom','right','insideH','insideV'):
        e = OxmlElement(f'w:{s}')
        e.set(qn('w:val'), 'none'); e.set(qn('w:sz'), '0')
        e.set(qn('w:space'), '0'); e.set(qn('w:color'), 'auto')
        b.append(e)
    tcPr.append(b)

def sh(cell, fill='D9D9D9'):
    tcPr = cell._tc.get_or_add_tcPr()
    s = OxmlElement('w:shd')
    s.set(qn('w:val'),'clear'); s.set(qn('w:color'),'auto'); s.set(qn('w:fill'),fill)
    tcPr.append(s)

def pgbrk():
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    br = OxmlElement('w:br'); br.set(qn('w:type'), 'page')
    p.add_run()._element.append(br)

def mktbl(rows, cols, col_widths_cm, row_heights_cm=None):
    """テーブル作成・列幅・行高さを一括設定"""
    t = doc.add_table(rows=rows, cols=cols)
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    for j, w in enumerate(col_widths_cm):
        for i in range(rows): t.cell(i, j).width = Cm(w)
    if row_heights_cm:
        for i, h in enumerate(row_heights_cm): set_row_h(t.rows[i], h)
    return t

# ─────────────────────────────────────────────────────
# ▌ ヘッダー
# ─────────────────────────────────────────────────────
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after  = Pt(3)
r = p.add_run('2026年度　山形県立山形南高等学校　２学年１学期中間テスト　論理・表現Ⅱ　解答用紙')
jp(r, 11, True)

th = mktbl(2, 4, [2.2, 3.0, 2.2, 10.6], [0.65, 0.65])
cw(th.cell(0,0), 'Class', 9, True); cw(th.cell(0,2), 'No', 9, True)
cw(th.cell(1,0), 'Name',  9, True)
th.cell(1,1).merge(th.cell(1,3))

# ─────────────────────────────────────────────────────
# ▌ １　Listening（思考・判断・表現　14点）
# ─────────────────────────────────────────────────────
# A/B：正答 ①②③ の 1 文字 → 小ボックス 1.8 cm
# C  ：(1)3語 (2)2語 (3)1語 (4)3語 (5)3語 → 語ボックス 2.65 cm
lbl('１　（思考・判断・表現　14点）')

# ── A / B ──
slbl('A（2点）　　　　　　　　　　　　　B（2点）　　各①〜③から選択')
# レイアウト: [A label 2.5][A ans 1.8][gap 0.5][B label 2.5][B ans 1.8][rest 9.9]
t1ab = mktbl(1, 6, [2.5, 1.8, 0.5, 2.5, 1.8, 8.9], [0.70])
cw(t1ab.cell(0,0), 'Aの答え', 8, True); sh(t1ab.cell(0,0))
cw(t1ab.cell(0,1), '')
no_b(t1ab.cell(0,2))
cw(t1ab.cell(0,3), 'Bの答え', 8, True); sh(t1ab.cell(0,3))
cw(t1ab.cell(0,4), '')
no_b(t1ab.cell(0,5))

# ── C ──
# 語ボックス幅：(W - 3×NW) / 6 blanks_row1 = (18-3×0.7)/6 = 15.9/6 = 2.65
# Row1: (1)[2.65][2.65][2.65]  (2)[2.65][2.65]  (3)[2.65]
#  cols: 0.7+2.65×3 + 0.7+2.65×2 + 0.7+2.65×1 = 8.65+6.0+3.35=18.0 ✓
# Row2: (4)[2.65][2.65][2.65]  (5)[2.65][2.65][2.65]  rest
#  cols: 0.7+2.65×3 + 0.7+2.65×3 + rest = 8.65+8.65+0.7=18.0 ✓
slbl('C（2×5=10点）')

NW_C = 0.7; BW_C = (W - 3*NW_C) / 6  # = 2.65

# Row1 テーブル (3+2+1=6 boxes, 3 labels → 9 cols)
r1w = [NW_C, BW_C, BW_C, BW_C,  NW_C, BW_C, BW_C,  NW_C, BW_C]
t1c1 = mktbl(2, 9, r1w, [0.44, 0.70])
for col, ql in [(0,'(1)'),(4,'(2)'),(7,'(3)')]:
    cw(t1c1.cell(0,col), ql, 8, True); sh(t1c1.cell(0,col))
    no_b(t1c1.cell(1,col))
for col in [1,2,3,5,6,8]:
    cw(t1c1.cell(0,col), ''); cw(t1c1.cell(1,col), '')

# Row2 テーブル (3+3=6 boxes, 2 labels, rest → 9 cols)
rest_r2 = W - (NW_C + 3*BW_C)*2  # = 18 - 8.65×2 = 0.7
r2w = [NW_C, BW_C, BW_C, BW_C,  NW_C, BW_C, BW_C, BW_C,  rest_r2]
t1c2 = mktbl(2, 9, r2w, [0.44, 0.70])
for col, ql in [(0,'(4)'),(4,'(5)')]:
    cw(t1c2.cell(0,col), ql, 8, True); sh(t1c2.cell(0,col))
    no_b(t1c2.cell(1,col))
for col in [1,2,3,5,6,7]:
    cw(t1c2.cell(0,col), ''); cw(t1c2.cell(1,col), '')
no_b(t1c2.cell(0,8)); no_b(t1c2.cell(1,8))

# ─────────────────────────────────────────────────────
# ▌ ２　LEAP Vocabulary（知識・技能　15点）
# ─────────────────────────────────────────────────────
# A：A/B/C/D の 1 文字 → 各列 1.8 cm
# B：英単語 1 語（purchase/suspect 等、最長12文字）→ 各列 3.6 cm
lbl('２　（知識・技能　15点）')

slbl('A（1×10=10点）　　各選択（A・B・C・D）')
t2a = mktbl(2, 10, [W/10]*10, [0.44, 0.68])
for j in range(10):
    cw(t2a.cell(0,j), f'({j+1})', 8, True); sh(t2a.cell(0,j))
    cw(t2a.cell(1,j), '')

slbl('B（1×5=5点）　　英単語1語')
t2b = mktbl(2, 5, [W/5]*5, [0.44, 0.70])
for j in range(5):
    cw(t2b.cell(0,j), f'({j+1})', 8, True); sh(t2b.cell(0,j))
    cw(t2b.cell(1,j), '')

# ─────────────────────────────────────────────────────
# ▌ ３　Grammar and Usage（知識・技能　28点）
# ─────────────────────────────────────────────────────
lbl('３　（知識・技能　28点）')

# ── A 語句整序：正答 = 文全体（8〜15 語）→ 1 問 1 行・フル幅 ──
slbl('A（1×10=10点）　※完全正答のみ得点')
NW_A = 0.7
t3a = mktbl(10, 2, [NW_A, W-NW_A], [0.72]*10)
for i in range(10):
    cw(t3a.cell(i,0), f'({i+1})', 8, True); sh(t3a.cell(i,0))
    cw(t3a.cell(i,1), '', 9, align=WD_ALIGN_PARAGRAPH.LEFT)

# ── B 空所補充：正答 1〜3 語 → 2 列レイアウト（左Q1-6・右Q7-12）──
# 各語ボックス幅 = (W/2 - 0.6) / 3 = (9 - 0.6) / 3 = 2.8 cm
# Q毎の空欄数: Q1=1, Q2=2, Q3=2, Q4=3, Q5=3, Q6=3, Q7=1, Q8=2, Q9=2, Q10=2, Q11=2, Q12=1
slbl('B（1×12=12点）')
b_blanks = [1,2,2,3,3,3,  1,2,2,2,2,1]
bx = (W/2 - 0.6) / 3  # = 2.8
t3b = mktbl(6, 8, [0.6, bx, bx, bx, 0.6, bx, bx, bx], [0.68]*6)
for i in range(6):
    for side in range(2):
        qi = i + side*6; bc = side*4; nb = b_blanks[qi]
        cw(t3b.cell(i,bc), f'({qi+1})', 8, True); sh(t3b.cell(i,bc))
        for b in range(3):
            c = bc+1+b
            if b < nb: cw(t3b.cell(i,c), '')
            else:       no_b(t3b.cell(i,c))

# ── C 動詞活用：1〜3 語（was studying / have been closed 等）→ 各列 3.0 cm ──
slbl('C（1×6=6点）')
t3c = mktbl(2, 6, [W/6]*6, [0.44, 0.70])
for j in range(6):
    cw(t3c.cell(0,j), f'({j+1})', 8, True); sh(t3c.cell(0,j))
    cw(t3c.cell(1,j), '')

# ══════════════ ページ区切り ══════════════
pgbrk()

# ─────────────────────────────────────────────────────
# ▌ ４　Writing（思考・判断・表現　12点）英訳
# ─────────────────────────────────────────────────────
# 正答：英文 1 文（8〜15 語）→ フル幅（17.3 cm）
lbl('４　（思考・判断・表現　12点）', before=0)
t4 = mktbl(6, 2, [0.7, W-0.7], [0.85]*6)
for i in range(6):
    cw(t4.cell(i,0), f'({i+1})', 8, True); sh(t4.cell(i,0))
    cw(t4.cell(i,1), '', 9, align=WD_ALIGN_PARAGRAPH.LEFT)

# ─────────────────────────────────────────────────────
# ▌ ５　Reading Comprehension Unit1&2（思考・判断・表現　10点）
# ─────────────────────────────────────────────────────
# 正答：英文 1〜2 文（A:5問・B:5問、各 5〜12 語）
# A/B 2 列並び → 各解答欄 8.3 cm
lbl('５　（思考・判断・表現　10点）')
slbl('A（1×5=5点）　　　　　　　　　　　　B（1×5=5点）', before=1)
NW5 = 0.7; AW5 = W/2 - NW5
t5 = mktbl(5, 4, [NW5, AW5, NW5, AW5], [0.80]*5)
for i in range(5):
    for s in range(2):
        cw(t5.cell(i,s*2),   f'({i+1})', 8, True); sh(t5.cell(i,s*2))
        cw(t5.cell(i,s*2+1), '', 9, align=WD_ALIGN_PARAGRAPH.LEFT)

# ─────────────────────────────────────────────────────
# ▌ ６　Reading Comprehension: Ancient Greece（思考・判断・表現　15点）
# ─────────────────────────────────────────────────────
lbl('６　（思考・判断・表現　15点）')

# 問1(1点)・問2(2点)・問5(2点)：正答 = 選択肢 A/B/C/D の 1 文字
# レイアウト: [lbl 3.0][ans 1.5][gap 0.5] ×3 + [rest 4.5]
# = 3×5.0 + 4.5 = 15+... → 3.0+1.5+0.5=5.0×3=15.0+4.5=19.5?
# 修正: [lbl 2.8][ans 1.5][gap 0.4] ×3 + rest
# = (2.8+1.5+0.4)×3=14.1+3.9=18.0 ✓
slbl('問1（1点）　問2（2点）　問5（2点）　　各 A・B・C・D から選択', before=1)
GW = 0.4; LW6 = 2.8; AW6 = 1.5  # lbl/answer/gap
rest6 = W - (LW6+AW6+GW)*3  # = 18 - 4.7×3 = 18 - 14.1 = 3.9
t6s = mktbl(2, 10, [LW6,AW6,GW, LW6,AW6,GW, LW6,AW6, GW, rest6], [0.44, 0.68])
for idx, qlbl in enumerate(['問1（1点）', '問2（2点）', '問5（2点）']):
    bc = idx*3
    cw(t6s.cell(0,bc), qlbl, 8, True); sh(t6s.cell(0,bc))
    cw(t6s.cell(1,bc), '答え', 7.5)
    cw(t6s.cell(0,bc+1), ''); cw(t6s.cell(1,bc+1), '')
    if idx < 2:
        no_b(t6s.cell(0,bc+2)); no_b(t6s.cell(1,bc+2))
    else:
        # 問5 の gap = col8, rest = col9
        no_b(t6s.cell(0,8)); no_b(t6s.cell(1,8))
        no_b(t6s.cell(0,9)); no_b(t6s.cell(1,9))

# 問3(5点)：正答 = S または A の 1 文字 × 5
# [lbl 2.0][ans 1.6] × 5 = 5×3.6 = 18.0 ✓
slbl('問3（1×5=5点）　　S＝スパルタ　A＝アテネ', before=2)
t6q3 = mktbl(2, 10,
             [2.0,1.6, 2.0,1.6, 2.0,1.6, 2.0,1.6, 2.0,1.6],
             [0.44, 0.68])
for j, ql in enumerate(['(i)','(ii)','(iii)','(iv)','(v)']):
    cw(t6q3.cell(0,j*2), ql, 8, True); sh(t6q3.cell(0,j*2))
    cw(t6q3.cell(1,j*2), '')
    cw(t6q3.cell(0,j*2+1), ''); cw(t6q3.cell(1,j*2+1), '')

# 問4(2点)：日本語 1〜2 文 → フル幅 2 行
slbl('問4（2点）　スパルタの女性がアテネの女性より「自由」を得られた理由（日本語）', before=2)
t6q4 = mktbl(2, 1, [W], [0.82, 0.82])
for i in range(2): cw(t6q4.cell(i,0), '', 9)

# 問6(3点)：日本語訳 1〜2 文 → フル幅 2 行
slbl('問6（3点）　下線部(2)を日本語にしなさい', before=2)
t6q6 = mktbl(2, 1, [W], [0.82, 0.82])
for i in range(2): cw(t6q6.cell(i,0), '', 9)

# ─────────────────────────────────────────────────────
# ▌ ７　Writing: Book Report（思考・判断・表現　6点）50語以上
# ─────────────────────────────────────────────────────
# 50語×平均5〜6文字≒280文字、1行18cm≒50文字 → 6行で十分
lbl('７　（思考・判断・表現　6点）　50語以上で書くこと')
t7 = mktbl(6, 1, [W], [0.83]*6)
for i in range(6): cw(t7.cell(i,0), '', 9)

# ─────────────────────────────────────────────────────
# ▌ 得点集計
# ─────────────────────────────────────────────────────
lbl('得点集計', before=5)
ts = mktbl(2, 3, [5.5, 6.5, 6.0], [0.44, 0.78])
for j, t in enumerate(['２・３（知識・技能）　/43', '１・４〜７（思考・判断・表現）　/57', '合計　/100']):
    cw(ts.cell(0,j), t, 8, True); sh(ts.cell(0,j))
    cw(ts.cell(1,j), '', 9)

# ── 保存 ────────────────────────────────────────────
out = '/home/user/Yumetan-Quick-Response/2026年度_論理表現Ⅱ_中間テスト_解答用紙.docx'
doc.save(out)
print('Saved:', out)
