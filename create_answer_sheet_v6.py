"""
2026年度 論理・表現Ⅱ 中間テスト 解答用紙（改訂版 v6）
ページ1: §1 + §2 + §3A  (27.022cm)
ページ2: §3B + §3C + §4 + §5 + §6 + 得点  (26.970cm)
ページ3: §7 英作文  (25.779cm)
"""
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

MARGINS = 1.2
W  = 21 - 2 * MARGINS   # 18.6 cm
FS = 10.5

# ─────────────── ユーティリティ ───────────────

def set_row_h(row, h_cm, rule='exact'):
    trPr = row._tr.get_or_add_trPr()
    e = OxmlElement('w:trHeight')
    e.set(qn('w:val'), str(int(h_cm * 567)))
    e.set(qn('w:hRule'), rule)
    trPr.append(e)

def font_run(run, size=FS, bold=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = 'Arial'
    rpr = run._element.get_or_add_rPr()
    rf  = rpr.get_or_add_rFonts()
    rf.set(qn('w:ascii'),    'Arial')
    rf.set(qn('w:hAnsi'),    'Arial')
    rf.set(qn('w:eastAsia'), 'MS PGothic')

def cw(cell, text='', size=FS, bold=False,
       align=WD_ALIGN_PARAGRAPH.LEFT,
       valign=WD_ALIGN_VERTICAL.CENTER):
    cell.vertical_alignment = valign
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    r = p.add_run(text)
    font_run(r, size, bold)

def para(doc, text, size=FS, bold=False,
         align=WD_ALIGN_PARAGRAPH.LEFT,
         before=0, after=0, line=13):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before      = Pt(before)
    p.paragraph_format.space_after       = Pt(after)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing      = Pt(line)
    r = p.add_run(text)
    font_run(r, size, bold)

def no_b(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement('w:tcBorders')
    for s in ('top','left','bottom','right','insideH','insideV'):
        e = OxmlElement(f'w:{s}')
        e.set(qn('w:val'),'none'); e.set(qn('w:sz'),'0')
        e.set(qn('w:space'),'0'); e.set(qn('w:color'),'auto')
        b.append(e)
    tcPr.append(b)

def pgbrk(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before      = Pt(0)
    p.paragraph_format.space_after       = Pt(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing      = Pt(1)
    br = OxmlElement('w:br'); br.set(qn('w:type'), 'page')
    p.add_run()._element.append(br)

def mktbl(doc, rows, cols, col_w, row_h=None):
    t = doc.add_table(rows=rows, cols=cols)
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    if isinstance(col_w, (int, float)):
        col_w = [col_w] * cols
    for j, w in enumerate(col_w):
        for i in range(rows):
            t.cell(i, j).width = Cm(w)
    if row_h is not None:
        hs = row_h if isinstance(row_h, list) else [row_h]*rows
        for i, h in enumerate(hs):
            set_row_h(t.rows[i], h)
    return t

def set_right_tab(p_obj, pos_cm):
    pPr = p_obj._p.get_or_add_pPr()
    tabs = OxmlElement('w:tabs')
    tab  = OxmlElement('w:tab')
    tab.set(qn('w:val'), 'right')
    tab.set(qn('w:pos'), str(int(pos_cm * 567)))
    tabs.append(tab); pPr.append(tabs)

# ─────────────── §3 データ ───────────────

# Q2: He would rather not see this movie.
# Q10: Nobody likes to be made fun of in public.
S3A = [
    ['(1)',  '',                              'annually.'],
    ['(2)',  '',                              '.'],
    ['(3)',  'It',                            'in summer.'],
    ['(4)',  'It is',                         'on the world map.'],
    ['(5)',  '',                              '.'],
    ['(6)',  'The author',                    'in her garden.'],
    ['(7)',  'It was quite frightening, but', '.'],
    ['(8)',  'In Japan,',                     '.'],
    ['(9)',  'Enjoy',                         '.'],
    ['(10)', '',                              'public.'],
]

# §3B 13問: Q1-13の空欄数
B3B = [1, 2, 2, 3, 3, 3, 1, 2, 2, 2, 2, 1, 1]

def blank_parens(n):
    if n == 1: return '(' + '　' * 16 + ')'
    if n == 2: return ('(' + '　' * 7 + ')') * 2
    if n == 3: return ('(' + '　' * 4 + ')') * 3
    return ''

# ─────────────── ドキュメント生成 ───────────────

def build(output_path):
    doc = Document()
    doc.styles['Normal'].paragraph_format.space_before = Pt(0)
    doc.styles['Normal'].paragraph_format.space_after  = Pt(0)

    for sec in doc.sections:
        sec.page_width  = Cm(21)
        sec.page_height = Cm(29.7)
        for attr in ('left','right','top','bottom'):
            setattr(sec, f'{attr}_margin', Cm(MARGINS))

    # ── ヘッダー  (0pt+14pt+3pt = 17pt = 0.600cm) ──
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before      = Pt(0)
    p.paragraph_format.space_after       = Pt(3)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing      = Pt(14)
    r = p.add_run(
        '2026年度　山形県立山形南高等学校　２学年１学期中間テスト　論理・表現Ⅱ　解答用紙')
    font_run(r, 11, True)

    # 氏名欄 2×0.72cm = 1.440cm
    th = mktbl(doc, 2, 4, [2.0, 3.0, 1.8, 2.5])
    for i in range(2): set_row_h(th.rows[i], 0.72)
    cw(th.cell(0,0), 'Class', 10, True, WD_ALIGN_PARAGRAPH.CENTER)
    cw(th.cell(0,2), 'No',    10, True, WD_ALIGN_PARAGRAPH.CENTER)
    cw(th.cell(1,0), 'Name',  10, True, WD_ALIGN_PARAGRAPH.CENTER)
    th.cell(1,1).merge(th.cell(1,3))

    # ══════════════ ページ１  (27.022cm) ══════════════

    # §1  label (5+13+2=20pt=0.706cm)
    para(doc, '１　（思考・判断・表現　14点）', bold=True, before=5, after=2, line=13)

    # A/B  1×1.0cm = 1.000cm
    t1ab = mktbl(doc, 1, 4, [3.2, 2.0, 3.2, 2.0], 1.0)
    cw(t1ab.cell(0,0), 'A（2点）', 9.5, True, WD_ALIGN_PARAGRAPH.CENTER)
    cw(t1ab.cell(0,2), 'B（2点）', 9.5, True, WD_ALIGN_PARAGRAPH.CENTER)

    # C label (2+12+0=14pt=0.494cm)
    para(doc, 'C（2×5=10点）', size=9.5, before=2, after=0, line=12)
    # C  1×1.2cm = 1.200cm  (語数に合わせた幅)
    t1c = mktbl(doc, 1, 5, [3.7, 2.8, 1.9, 3.7, 3.7], 1.2)
    for j, q in enumerate(['(1)','(2)','(3)','(4)','(5)']):
        cw(t1c.cell(0,j), q, FS)

    # §2  label (4+13+2=19pt=0.670cm)
    para(doc, '２　（知識・技能　15点）', bold=True, before=4, after=2, line=13)

    # A label (2+12+0=14pt=0.494cm)
    para(doc, 'A（1×10=10点）　選択（A・B・C・D）', size=9.5, before=2, after=0, line=12)
    # A  2×0.88cm = 1.760cm
    t2a = mktbl(doc, 2, 5, [3.0]*5, 0.88)
    for i in range(2):
        for j in range(5):
            cw(t2a.cell(i,j), f'({i*5+j+1})', FS)

    # B label (2+12+0=0.494cm)
    para(doc, 'B（1×5=5点）　英単語1語', size=9.5, before=2, after=0, line=12)
    # B  1×1.0cm = 1.000cm
    t2b = mktbl(doc, 1, 5, [3.6]*5, 1.0)
    for j in range(5):
        cw(t2b.cell(0,j), f'({j+1})', FS)

    # §3  label (4+13+2=19pt=0.670cm)
    para(doc, '３　（知識・技能　29点）', bold=True, before=4, after=2, line=13)

    # §3A label (2+12+0=14pt=0.494cm)
    para(doc, 'A（1×10=10点）　語句を並び替えて英文を完成させよ　完全正答のみ得点',
         size=9.5, before=2, after=0, line=12)
    # §3A  10×1.6cm = 16.000cm
    NUM_W = 1.2
    ANS_W = W - NUM_W   # 17.4cm
    t3a = mktbl(doc, 10, 2, [NUM_W, ANS_W], 1.6)
    for i, (num, pre, post) in enumerate(S3A):
        cw(t3a.cell(i,0), num, FS, False, WD_ALIGN_PARAGRAPH.CENTER)
        cell = t3a.cell(i,1)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)
        set_right_tab(p, ANS_W - 0.15)
        r1 = p.add_run((pre + '  ') if pre else '')
        font_run(r1)
        p.add_run('\t')
        r2 = p.add_run(post)
        font_run(r2)

    # ══════════════ ページ区切り ══════════════
    pgbrk(doc)

    # ══════════════ ページ２  (26.970cm) ══════════════

    # §3B label (4+13+2=0.670cm)
    para(doc, 'B（1×13=13点）　空所に適切な語を入れよ', bold=False,
         size=9.5, before=4, after=2, line=13)
    # §3B  7×0.82cm = 5.740cm  (Q1-7左 / Q8-13右, row7右は空)
    NW = 0.8; AW = (W - NW*2) / 2   # 8.50cm
    t3b = mktbl(doc, 7, 4, [NW, AW, NW, AW], 0.82)
    for i in range(7):
        for side in range(2):
            qi = i + side*7   # Q1-7(left) / Q8-14(right)→ Q8-13 = right rows 0-5
            if side == 1:
                qi = i + 7   # Q8..Q14
            nc = side*2; ac = side*2+1
            if side == 0 or qi <= 12:   # Q1-7 left, Q8-13 right (qi 7..12 → i 0..5)
                actual_q = qi + 1
                cw(t3b.cell(i,nc), f'({actual_q})', 9.5, False, WD_ALIGN_PARAGRAPH.CENTER)
                cw(t3b.cell(i,ac), blank_parens(B3B[qi]), FS, False,
                   WD_ALIGN_PARAGRAPH.CENTER)
            else:   # row 6, right side: Q14 doesn't exist → hide
                no_b(t3b.cell(i,nc)); no_b(t3b.cell(i,ac))

    # §3C label (2+12+0=0.494cm)
    para(doc, 'C（1×6=6点）　動詞を正しい形に変えよ', size=9.5, before=2, after=0, line=12)
    # §3C  1×0.82cm = 0.820cm
    t3c = mktbl(doc, 1, 6, [3.0]*6, 0.82)
    for j in range(6):
        cw(t3c.cell(0,j), f'({j+1})', FS)

    # §4  label (4+13+2=0.670cm)
    para(doc, '４　（思考・判断・表現　18点）', bold=True, before=4, after=2, line=13)
    # §4  6×0.88cm = 5.280cm
    t4 = mktbl(doc, 6, 1, [W], 0.88)
    for i in range(6):
        cw(t4.cell(i,0), f'({i+1})', FS)

    # §5  label (4+13+2=0.670cm)
    para(doc, '５　（思考・判断・表現　12点）', bold=True, before=4, after=2, line=13)
    # A/B sub-label (1+12+0=13pt=0.459cm)
    para(doc, 'A（2×3=6点）　　　　　　　　　　　　　　　　B（2×3=6点）',
         size=9.5, before=1, after=0, line=12)
    # §5  3×0.95cm (A列|B列) = 2.850cm
    HW = W / 2   # 9.3cm
    t5 = mktbl(doc, 3, 2, [HW, HW], 0.95)
    for i in range(3):
        for j in range(2):
            cw(t5.cell(i,j), f'({i+1})', FS)

    # §6  label (3+13+2=0.635cm)
    para(doc, '６　（思考・判断・表現　10点）', bold=True, before=3, after=2, line=13)

    # 問1/2/4 sub-label (2+12+0=0.494cm)
    para(doc, '問1（2点）・問2（2点）・問4（1点）', size=9.5, before=2, after=0, line=12)
    # 問1/2/4 combined row  1×0.88cm = 0.880cm
    t6s = mktbl(doc, 1, 3, [5.6, 9.0, 4.0], 0.88)
    cw(t6s.cell(0,0), '問1', 9.5, True, WD_ALIGN_PARAGRAPH.CENTER)
    cw(t6s.cell(0,1), '問2　（i〜vから2つ）', 9.5, True, WD_ALIGN_PARAGRAPH.CENTER)
    cw(t6s.cell(0,2), '問4', 9.5, True, WD_ALIGN_PARAGRAPH.CENTER)

    # 問3 sub-label (2+12+0=0.494cm)
    para(doc, '問3（2点）　スパルタの女性がアテネの女性よりも「自由」を得られていた理由',
         size=9.5, before=2, after=0, line=12)
    # 問3  2×0.88cm = 1.760cm
    t6q3 = mktbl(doc, 2, 1, [W], 0.88)
    for i in range(2): cw(t6q3.cell(i,0), '')

    # 問5 sub-label (2+12+0=0.494cm)
    para(doc, '問5（3点）　下線部(2)の日本語訳', size=9.5, before=2, after=0, line=12)
    # 問5  2×0.88cm = 1.760cm
    t6q5 = mktbl(doc, 2, 1, [W], 0.88)
    for i in range(2): cw(t6q5.cell(i,0), '')

    # 得点集計  label (3+13+1=17pt=0.600cm)
    para(doc, '得点集計', bold=True, before=3, after=1, line=13)
    # 得点  0.55+1.15 = 1.700cm
    ts = mktbl(doc, 2, 3, [5.8, 7.2, 5.6])
    set_row_h(ts.rows[0], 0.55); set_row_h(ts.rows[1], 1.15)
    for j, lbl in enumerate(['２・３（知識・技能）　/44',
                              '１・４〜７（思考・判断・表現）　/60',
                              '合計　/104']):
        cw(ts.cell(0,j), lbl, 9, True, WD_ALIGN_PARAGRAPH.CENTER)
        cw(ts.cell(1,j), '', FS, align=WD_ALIGN_PARAGRAPH.CENTER)

    # ══════════════ ページ区切り（§7用） ══════════════
    pgbrk(doc)

    # ══════════════ ページ３  §7 英作文  (25.779cm) ══════════════

    # 氏名欄（1行） 0.720cm
    th3 = mktbl(doc, 1, 4, [2.0, 3.0, 1.8, 2.5])
    set_row_h(th3.rows[0], 0.72)
    cw(th3.cell(0,0), 'Class', 10, True, WD_ALIGN_PARAGRAPH.CENTER)
    cw(th3.cell(0,2), 'No',    10, True, WD_ALIGN_PARAGRAPH.CENTER)
    th3.cell(0,1).merge(th3.cell(0,3))

    # §7  label (2+14+2=18pt=0.635cm → 使用: before=2,after=2,line=14)
    p = doc.add_paragraph()
    p.paragraph_format.space_before      = Pt(2)
    p.paragraph_format.space_after       = Pt(2)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing      = Pt(14)
    r = p.add_run('７　（思考・判断・表現　6点）　Writing: Book Report')
    font_run(r, 11, True)

    # サブノート (0+12+2=14pt=0.494cm)
    para(doc,
         'Introduce a book you have read recently or a book you like.  (at least 50 words)',
         size=9.5, before=0, after=2, line=12)

    # 英作文 16×1.5cm = 24.000cm
    t7 = mktbl(doc, 16, 1, [W], 1.5)
    for i in range(16): cw(t7.cell(i,0), '')

    doc.save(output_path)
    print(f'Saved: {output_path}')


if __name__ == '__main__':
    build('/home/user/Yumetan-Quick-Response/2026年度_論理表現Ⅱ_中間テスト_解答用紙.docx')
