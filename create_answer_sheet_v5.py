"""
2026年度 論理・表現Ⅱ 中間テスト 解答用紙
参考ファイルに合わせて作成：
  ・英語フォント Arial  / 日本語フォント MS PGothic
  ・番号は解答セル内テキスト（別カラムなし・シェーディングなし）
  ・テーブルは内容に応じた幅（右端まで広げない）
  ・A4 2ページ / マージン 1.2cm
"""
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

MARGINS = 1.2   # cm
W       = 21 - 2 * MARGINS  # = 18.6 cm
FS      = 10.5              # 本文フォントサイズ (pt)

# ──────────────────────────── ユーティリティ ────────────────────────────

def set_row_h(row, h_cm, rule='exact'):
    trPr = row._tr.get_or_add_trPr()
    e = OxmlElement('w:trHeight')
    e.set(qn('w:val'), str(int(h_cm * 567)))
    e.set(qn('w:hRule'), rule)
    trPr.append(e)

def font_run(run, size=FS, bold=False):
    """英語 Arial / 日本語 MS PGothic"""
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = 'Arial'
    rpr = run._element.get_or_add_rPr()
    rf = rpr.get_or_add_rFonts()
    rf.set(qn('w:ascii'),   'Arial')
    rf.set(qn('w:hAnsi'),   'Arial')
    rf.set(qn('w:eastAsia'),'MS PGothic')

def cw(cell, text='', size=FS, bold=False,
       align=WD_ALIGN_PARAGRAPH.LEFT,
       valign=WD_ALIGN_VERTICAL.CENTER):
    """セルにテキストを書く"""
    cell.vertical_alignment = valign
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    r = p.add_run(text)
    font_run(r, size, bold)

def para(doc, text, size=FS, bold=False,
         align=WD_ALIGN_PARAGRAPH.LEFT,
         before=4, after=2):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after  = Pt(after)
    r = p.add_run(text)
    font_run(r, size, bold)

def no_b(cell):
    """セルの罫線を消す"""
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
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    p.add_run()._element.append(br)

def mktbl(doc, rows, cols, col_w, row_h=None, h_rule='exact'):
    """テーブルを作成し列幅・行高を設定"""
    t = doc.add_table(rows=rows, cols=cols)
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    # col_w は float list または float のどちらでも OK
    if isinstance(col_w, (int, float)):
        col_w = [col_w] * cols
    for j, w in enumerate(col_w):
        for i in range(rows):
            t.cell(i, j).width = Cm(w)
    if row_h is not None:
        heights = row_h if isinstance(row_h, list) else [row_h]*rows
        for i, h in enumerate(heights):
            set_row_h(t.rows[i], h, h_rule)
    return t

# ──────────────────────────── ドキュメント生成 ────────────────────────────

def build(output_path):
    doc = Document()
    for sec in doc.sections:
        sec.page_width    = Cm(21)
        sec.page_height   = Cm(29.7)
        for attr in ('left','right','top','bottom'):
            setattr(sec, f'{attr}_margin', Cm(MARGINS))

    # ── ヘッダー ──────────────────────────────────────────
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(3)
    r = p.add_run(
        '2026年度　山形県立山形南高等学校　２学年１学期中間テスト　論理・表現Ⅱ　解答用紙')
    font_run(r, 11, True)

    # Class / No / Name ─ 参考ファイルに倣い右端まで広げない
    # Row1: Class[2.0] ans[3.0] | No[1.8] ans[2.5]
    # Row2: Name[2.0] ans[9.3]  (merged)
    th = mktbl(doc, 2, 4, [2.0, 3.0, 1.8, 2.5])
    for i in range(2):
        set_row_h(th.rows[i], 0.72, 'atLeast')
    cw(th.cell(0,0), 'Class', 10, True, WD_ALIGN_PARAGRAPH.CENTER)
    cw(th.cell(0,2), 'No',    10, True, WD_ALIGN_PARAGRAPH.CENTER)
    cw(th.cell(1,0), 'Name',  10, True, WD_ALIGN_PARAGRAPH.CENTER)
    th.cell(1,1).merge(th.cell(1,3))

    # ══════════════ ページ１ ══════════════

    # ── §1 Listening（思考・判断・表現　14点）────────────────────
    # A/B：答えは①②③ の 1 文字 → 小さめのセル
    # C ：各空欄 1〜3 語 → 語数に合わせた幅
    para(doc, '１　（思考・判断・表現　14点）', bold=True, before=6, after=2)

    # A（2点）/ B（2点）
    para(doc, 'A（2点）・B（2点）', size=9.5, before=2, after=1)
    t1ab = mktbl(doc, 1, 4, [3.2, 2.0, 3.2, 2.0], 0.84)
    cw(t1ab.cell(0,0), 'A（2点）', 9.5, True, WD_ALIGN_PARAGRAPH.CENTER)
    cw(t1ab.cell(0,1), '')
    cw(t1ab.cell(0,2), 'B（2点）', 9.5, True, WD_ALIGN_PARAGRAPH.CENTER)
    cw(t1ab.cell(0,3), '')

    # C（10点）：空欄数に合わせた幅 (1)3語 (2)2語 (3)1語 (4)3語 (5)3語
    # 各セルに番号を記入、幅は語数×2.3cm
    # (1)6.9 (2)4.6 (3)2.3 (4)6.9 (5)6.9 → でも合計27cm超
    # 語ボックスを分けずに1問=1セルにする（幅は語数×1.9cm）
    # (1)3.7 (2)2.8 (3)1.9 (4)3.7 (5)3.7 → 合計15.8cm ✓
    para(doc, 'C（2×5=10点）', size=9.5, before=3, after=1)
    c_widths = [3.7, 2.8, 1.9, 3.7, 3.7]  # (1)3語 (2)2語 (3)1語 (4)3語 (5)3語
    t1c = mktbl(doc, 1, 5, c_widths, 0.84)
    for j, q in enumerate(['(1)', '(2)', '(3)', '(4)', '(5)']):
        cw(t1c.cell(0,j), q, FS, False)

    # ── §2 LEAP Vocabulary（知識・技能　15点）─────────────────────
    # A：答えは A/B/C/D → 3.0cm/問 × 5問 × 2行 = 15cm
    # B：答えは英単語 1 語 → 3.6cm/問 × 5問 = 18cm
    para(doc, '２　（知識・技能　15点）', bold=True, before=5, after=2)

    para(doc, 'A（1×10=10点）　選択（A・B・C・D）', size=9.5, before=2, after=1)
    t2a = mktbl(doc, 2, 5, [3.0]*5, 0.84)
    for i in range(2):
        for j in range(5):
            cw(t2a.cell(i,j), f'({i*5+j+1})', FS)

    para(doc, 'B（1×5=5点）　英単語1語', size=9.5, before=3, after=1)
    t2b = mktbl(doc, 1, 5, [3.6]*5, 0.84)
    for j in range(5):
        cw(t2b.cell(0,j), f'({j+1})', FS)

    # ── §3 Grammar and Usage（知識・技能　28点）──────────────────
    para(doc, '３　（知識・技能　28点）', bold=True, before=5, after=2)

    # A 語句整序：答えは完全な英文 → 1問1行・全幅
    para(doc, 'A（1×10=10点）　完全正答のみ得点', size=9.5, before=2, after=1)
    t3a = mktbl(doc, 5, 2, [9.2, 9.2], 1.0)
    for i in range(5):
        for j in range(2):
            cw(t3a.cell(i,j), f'({i*2+j+1})', FS)

    # B 空所補充：各問 1〜3 語 → 語ボックスを問数分だけ表示
    # 2列レイアウト（左 Q1-6 / 右 Q7-12）
    # 各側：番号[0.8cm] + 語ボックス[2.2cm]×3 = 7.4cm
    # 合計 14.8cm（右端まで使わない）
    para(doc, 'B（1×12=12点）', size=9.5, before=3, after=1)
    b_blanks = [1,2,2,3,3,3, 1,2,2,2,2,1]
    NB = 0.8; BX = 2.2
    t3b = mktbl(doc, 6, 8, [NB,BX,BX,BX, NB,BX,BX,BX], 0.72)
    for i in range(6):
        for side in range(2):
            qi = i + side*6; bc = side*4; nb = b_blanks[qi]
            cw(t3b.cell(i,bc), f'({qi+1})', 9.5, True, WD_ALIGN_PARAGRAPH.CENTER)
            for b in range(3):
                c = bc+1+b
                if b < nb: cw(t3b.cell(i,c), '')
                else:       no_b(t3b.cell(i,c))

    # C 動詞活用：答えは動詞形 1〜3 語 → 3cm × 6 = 18cm
    para(doc, 'C（1×6=6点）', size=9.5, before=3, after=1)
    t3c = mktbl(doc, 1, 6, [3.0]*6, 0.84)
    for j in range(6):
        cw(t3c.cell(0,j), f'({j+1})', FS)

    # ══════════════ ページ区切り ══════════════
    pgbrk(doc)

    # ── §4 Writing / 英訳（思考・判断・表現　12点）───────────────
    # 答えは英文 1 文 → フル幅・1.0cm/行
    para(doc, '４　（思考・判断・表現　12点）', bold=True, before=0, after=2)
    t4 = mktbl(doc, 6, 1, [W], 1.0)
    for i in range(6):
        cw(t4.cell(i,0), f'({i+1})', FS)

    # ── §5 Reading Comprehension Unit1&2（思考・判断・表現　10点）──
    # A/B 各 5 問：英文 1〜2 文 → 2列並び、各 0.9cm
    para(doc, '５　（思考・判断・表現　10点）', bold=True, before=5, after=2)
    para(doc, 'A（1×5=5点）　　　　　　　　　　　　　B（1×5=5点）',
         size=9.5, before=1, after=1)
    HW5 = W / 2  # = 9.3cm
    t5 = mktbl(doc, 5, 2, [HW5, HW5], 0.90)
    for i in range(5):
        for j in range(2):
            cw(t5.cell(i,j), f'({i+1})', FS)

    # ── §6 Reading / Ancient Greece（思考・判断・表現　15点）───────
    para(doc, '６　（思考・判断・表現　15点）', bold=True, before=5, after=2)

    # 問1(1点)・問2(2点)・問5(2点) — A/B/C/D 1文字 → 小さめ
    # [問1 3.2cm][問2 3.2cm][問5 3.2cm] = 9.6cm（右端まで使わない）
    para(doc, '問1（1点）・問2（2点）・問5（2点）　各 A〜D から選択',
         size=9.5, before=2, after=1)
    t6s = mktbl(doc, 1, 3, [3.2, 3.2, 3.2], 0.84)
    for j, lbl in enumerate(['問1（1点）', '問2（2点）', '問5（2点）']):
        cw(t6s.cell(0,j), lbl, 9.5, True, WD_ALIGN_PARAGRAPH.CENTER)

    # 問3(5点) — S または A → 2.4cm × 5 = 12cm（右端まで使わない）
    para(doc, '問3（1×5=5点）　S＝スパルタ　A＝アテネ', size=9.5, before=3, after=1)
    t6q3 = mktbl(doc, 1, 5, [2.4]*5, 0.84)
    for j, lbl in enumerate(['(i)', '(ii)', '(iii)', '(iv)', '(v)']):
        cw(t6q3.cell(0,j), lbl, FS, False, WD_ALIGN_PARAGRAPH.CENTER)

    # 問4(2点) — 日本語記述 → フル幅・2行
    para(doc, '問4（2点）　スパルタの女性がアテネの女性より「自由」を得られた理由',
         size=9.5, before=3, after=1)
    t6q4 = mktbl(doc, 2, 1, [W], 0.88)
    for i in range(2): cw(t6q4.cell(i,0), '')

    # 問6(3点) — 日本語訳 → フル幅・2行
    para(doc, '問6（3点）　下線部(2)の日本語訳', size=9.5, before=3, after=1)
    t6q6 = mktbl(doc, 2, 1, [W], 0.88)
    for i in range(2): cw(t6q6.cell(i,0), '')

    # ── §7 Writing: Book Report（思考・判断・表現　6点）────────────
    # 50語以上 → 6行・フル幅・0.83cm/行
    para(doc, '７　（思考・判断・表現　6点）　50語以上', bold=True,
         before=5, after=2)
    t7 = mktbl(doc, 6, 1, [W], 0.83)
    for i in range(6): cw(t7.cell(i,0), '')

    # ── 得点集計 ─────────────────────────────────────────────────
    para(doc, '得点集計', bold=True, before=5, after=1)
    ts = mktbl(doc, 2, 3, [5.8, 7.2, 5.6])
    set_row_h(ts.rows[0], 0.55); set_row_h(ts.rows[1], 0.75)
    for j, lbl in enumerate(['２・３（知識・技能）　/43',
                              '１・４〜７（思考・判断・表現）　/57',
                              '合計　/100']):
        cw(ts.cell(0,j), lbl, 9, True, WD_ALIGN_PARAGRAPH.CENTER)
        cw(ts.cell(1,j), '', FS, align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.save(output_path)
    print(f'Saved: {output_path}')


if __name__ == '__main__':
    build('/home/user/Yumetan-Quick-Response/2026年度_論理表現Ⅱ_中間テスト_解答用紙.docx')
