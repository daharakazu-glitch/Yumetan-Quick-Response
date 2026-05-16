"""
2026年度 論理・表現Ⅱ 中間テスト
  ① 解答用紙 (blank)
  ② 解答例   (filled)
を生成する。
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ════════════════════════════════════════════
#  解答データ
# ════════════════════════════════════════════
ANSWERS = {
    # §1 Listening：音声なしのため省略表記
    '1a': '（音声による）',
    '1b': '（音声による）',
    # C: 各空欄の語リスト
    '1c': {
        1: ['have', 'been', 'launched'],
        2: ['no', 'longer'],
        3: ['includes'],
        4: ['threaten', 'at', 'risk'],
        5: ['in', 'order', 'to'],
    },
    # §2A 選択肢
    '2a': ['A', 'C', 'D', 'C', 'C', 'A', 'C', 'A', 'B', 'C'],
    # §2B 単語
    '2b': ['purchase', 'suspect', 'permitted', 'miserable', 'superstition'],
    # §3A 語句整序（完成文）
    '3a': [
        'A good number of people visit the country annually.',
        'China is also famous for its large structures in the world.',
        'It is easy to live in Iceland in summer.',
        'It is interesting that the country looks like a large cat on the world map.',
        'There are too many new places in the world to explore.',
        'The author was inspired by the 100-year-old tree in her garden.',
        'a little more horridness would have made the story better',
        'it is very important to entertain guests',
        'reading how these dinosaur fossils were found',
        'is looked up to as a great witch',
    ],
    # §3B 空所補充（各問の語リスト）
    '3b': [
        ['We'],
        ['It', 'to'],
        ['There', 'are'],
        ['enabled', 'her', 'to'],
        ['prevented', 'from', 'going'],
        ['has', 'been', 'translated'],
        ['contains'],
        ['located', 'center'],
        ['climate', 'changes'],
        ['water', 'pollution'],
        ['aging', 'society'],
        ['cutting-edge'],
    ],
    # §3C 動詞活用
    '3c': ['was studying', 'are moving', 'have been closed',
           'kicks', 'is getting', 'lived'],
    # §4 英訳
    '4': [
        "His speech was not as interesting as Alan's.",
        'My cell phone is not as new as hers.',
        'No secretary is as capable as Yuki.',
        'What the man said was far from the truth.',
        'The main character was spoken to by a stranger at the station.',
        'The painting has been protected by bullet-proof glass since 1956.',
    ],
    # §5A・B リーディング解答（英文）
    '5a': [
        'She loves wild animals.',
        'There are over 300 species of birds.',
        'We can see the proboscis monkey.',
        'Brunei is a small country in South-East Asia.',
        'There are around 100 mammals.',
    ],
    '5b': [
        'It is about a boy with a serious disease.',
        'Our differences make us all unique and special.',
        'Because the book is so beautifully written.',
        'R. J. Palacio.',
        'Because of his personality.',
    ],
    # §6 Ancient Greece
    '6_q1': 'B',
    '6_q2': 'B',
    '6_q5': 'A',
    '6_q3': ['S', 'A', 'A', 'S', 'A'],
    '6_q4': [
        'スパルタの夫たちは軍隊と共に家を離れて生活していたため、',
        '妻（女性）は家庭内で自由に行動できたから。',
    ],
    '6_q6': [
        'スパルタはアテネの人々が他の都市を攻撃しない限り、',
        '彼らが自分たちの生活様式を続けることを認めた。',
    ],
    # §7 ブックレポート（解答例）
    '7': [
        'I would like to introduce The Alchemist by Paulo Coelho.',
        'It is about a young shepherd named Santiago who travels from Spain to Egypt',
        'in search of treasure. He meets many people and learns lessons about dreams.',
        'I love this book because it teaches us to listen to our hearts.',
        'It shows that the journey itself is more valuable than the destination.',
        '',
    ],
}

# ════════════════════════════════════════════
#  ユーティリティ
# ════════════════════════════════════════════
def set_row_h(row, h_cm, rule='exact'):
    trPr = row._tr.get_or_add_trPr()
    e = OxmlElement('w:trHeight')
    e.set(qn('w:val'), str(int(h_cm * 567)))
    e.set(qn('w:hRule'), rule)
    trPr.append(e)

def jp(run, size, bold=False, color=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = 'MS Gothic'
    rpr = run._element.get_or_add_rPr()
    rpr.get_or_add_rFonts().set(qn('w:eastAsia'), 'MS Gothic')
    if color:
        run.font.color.rgb = RGBColor(*color)

def cw(cell, text='', size=9, bold=False,
       align=WD_ALIGN_PARAGRAPH.CENTER,
       valign=WD_ALIGN_VERTICAL.CENTER,
       color=None):
    cell.vertical_alignment = valign
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    r = p.add_run(text)
    jp(r, size, bold, color)

def lbl(doc, text, size=9, bold=True, before=5, after=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after  = Pt(after)
    r = p.add_run(text); jp(r, size, bold)

def slbl(doc, text, size=8.5, bold=False, before=2, after=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after  = Pt(after)
    r = p.add_run(text); jp(r, size, bold)

def no_b(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement('w:tcBorders')
    for s in ('top','left','bottom','right','insideH','insideV'):
        e = OxmlElement(f'w:{s}')
        e.set(qn('w:val'),'none'); e.set(qn('w:sz'),'0')
        e.set(qn('w:space'),'0'); e.set(qn('w:color'),'auto')
        b.append(e)
    tcPr.append(b)

def sh(cell, fill='D9D9D9'):
    tcPr = cell._tc.get_or_add_tcPr()
    s = OxmlElement('w:shd')
    s.set(qn('w:val'),'clear'); s.set(qn('w:color'),'auto'); s.set(qn('w:fill'),fill)
    tcPr.append(s)

def pgbrk(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    br = OxmlElement('w:br'); br.set(qn('w:type'), 'page')
    p.add_run()._element.append(br)

def mktbl(doc, rows, cols, col_widths_cm, row_heights_cm=None, h_rule='exact'):
    t = doc.add_table(rows=rows, cols=cols)
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    for j, w in enumerate(col_widths_cm):
        for i in range(rows): t.cell(i, j).width = Cm(w)
    if row_heights_cm:
        for i, h in enumerate(row_heights_cm):
            set_row_h(t.rows[i], h, h_rule)
    return t

ANS_COLOR = (0xC0, 0x00, 0x00)   # 解答の文字色（赤）

def ans_text(text, bold=True):
    """解答モードで使うテキスト・色・サイズ"""
    return {'text': text, 'bold': bold, 'color': ANS_COLOR}

# ════════════════════════════════════════════
#  ドキュメント生成本体
# ════════════════════════════════════════════
def build(filled: bool, output_path: str):
    doc = Document()
    for sec in doc.sections:
        sec.page_width    = Cm(21)
        sec.page_height   = Cm(29.7)
        sec.left_margin   = Cm(1.2)
        sec.right_margin  = Cm(1.2)
        sec.top_margin    = Cm(1.2)
        sec.bottom_margin = Cm(1.2)

    W = 18.6  # 本文幅 (cm)

    def ans(cell, text, size=9, align=WD_ALIGN_PARAGRAPH.CENTER):
        """filled=True のとき赤字で解答を書く"""
        if filled:
            cw(cell, text, size=size, bold=True, align=align, color=ANS_COLOR)

    # ────────── ヘッダー ──────────
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(3)
    title = '2026年度　山形県立山形南高等学校　２学年１学期中間テスト　論理・表現Ⅱ　'
    title += '解答例' if filled else '解答用紙'
    r = p.add_run(title); jp(r, 11, True)

    th = mktbl(doc, 2, 4, [2.2, 3.0, 2.2, 11.2], [0.60, 0.60])
    cw(th.cell(0,0), 'Class', 9, True); cw(th.cell(0,2), 'No', 9, True)
    cw(th.cell(1,0), 'Name',  9, True)
    th.cell(1,1).merge(th.cell(1,3))
    if filled:
        cw(th.cell(0,1), ''); cw(th.cell(0,3), '')
        cw(th.cell(1,1), '（解答例）', 9)

    # ────────── §1 Listening 14点 ──────────
    lbl(doc, '１　（思考・判断・表現　14点）')

    slbl(doc, 'A（2点）　　　　　　　　　　　　　B（2点）　　各①〜③から選択')
    # [A lbl 2.5][A ans 1.8][gap 0.5][B lbl 2.5][B ans 1.8][rest]
    rest_ab = W - (2.5+1.8+0.5) - (2.5+1.8)  # gap は右側には不要
    t1ab = mktbl(doc, 1, 6, [2.5,1.8,0.5,2.5,1.8,rest_ab], [0.68])
    cw(t1ab.cell(0,0), 'Aの答え', 8, True); sh(t1ab.cell(0,0))
    cw(t1ab.cell(0,1), ''); no_b(t1ab.cell(0,2))
    cw(t1ab.cell(0,3), 'Bの答え', 8, True); sh(t1ab.cell(0,3))
    cw(t1ab.cell(0,4), ''); no_b(t1ab.cell(0,5))
    if filled:
        cw(t1ab.cell(0,1), ANSWERS['1a'], 7, True, color=ANS_COLOR)
        cw(t1ab.cell(0,4), ANSWERS['1b'], 7, True, color=ANS_COLOR)

    # C: 2行に分割。語ボックス幅 = (W - 3×0.7) / 6 ≒ 2.75
    slbl(doc, 'C（2×5=10点）')
    NW = 0.7; BW = (W - 3*NW) / 6  # ≈2.75
    # Row1: (1)[3box] (2)[2box] (3)[1box] → 9cols
    r1w = [NW,BW,BW,BW, NW,BW,BW, NW,BW]
    t1c1 = mktbl(doc, 2, 9, r1w, [0.42, 0.68])
    for col, ql in [(0,'(1)'),(4,'(2)'),(7,'(3)')]:
        cw(t1c1.cell(0,col), ql, 8, True); sh(t1c1.cell(0,col))
        no_b(t1c1.cell(1,col))
    for col in [1,2,3,5,6,8]:
        cw(t1c1.cell(0,col), ''); cw(t1c1.cell(1,col), '')
    if filled:
        for qi, cols in [(1,[1,2,3]),(2,[5,6]),(3,[8])]:
            words = ANSWERS['1c'][qi]
            for k, col in enumerate(cols):
                cw(t1c1.cell(1,col), words[k] if k<len(words) else '', 8, True, color=ANS_COLOR)

    # Row2: (4)[3box] (5)[3box] + rest
    pad = W - (NW+3*BW)*2
    r2w = [NW,BW,BW,BW, NW,BW,BW,BW, pad]
    t1c2 = mktbl(doc, 2, 9, r2w, [0.42, 0.68])
    for col, ql in [(0,'(4)'),(4,'(5)')]:
        cw(t1c2.cell(0,col), ql, 8, True); sh(t1c2.cell(0,col))
        no_b(t1c2.cell(1,col))
    for col in [1,2,3,5,6,7]:
        cw(t1c2.cell(0,col), ''); cw(t1c2.cell(1,col), '')
    no_b(t1c2.cell(0,8)); no_b(t1c2.cell(1,8))
    if filled:
        for qi, cols in [(4,[1,2,3]),(5,[5,6,7])]:
            words = ANSWERS['1c'][qi]
            for k, col in enumerate(cols):
                cw(t1c2.cell(1,col), words[k] if k<len(words) else '', 8, True, color=ANS_COLOR)

    # ────────── §2 Vocabulary 15点 ──────────
    lbl(doc, '２　（知識・技能　15点）')

    slbl(doc, 'A（1×10=10点）　　各選択（A・B・C・D）')
    t2a = mktbl(doc, 2, 10, [W/10]*10, [0.42, 0.66])
    for j in range(10):
        cw(t2a.cell(0,j), f'({j+1})', 8, True); sh(t2a.cell(0,j))
        cw(t2a.cell(1,j), '')
        if filled: cw(t2a.cell(1,j), ANSWERS['2a'][j], 9, True, color=ANS_COLOR)

    slbl(doc, 'B（1×5=5点）　　英単語1語')
    t2b = mktbl(doc, 2, 5, [W/5]*5, [0.42, 0.68])
    for j in range(5):
        cw(t2b.cell(0,j), f'({j+1})', 8, True); sh(t2b.cell(0,j))
        cw(t2b.cell(1,j), '')
        if filled: cw(t2b.cell(1,j), ANSWERS['2b'][j], 9, True, color=ANS_COLOR)

    # ────────── §3 Grammar 28点 ──────────
    lbl(doc, '３　（知識・技能　28点）')

    # A 語句整序：2問/行（各解答欄 ≈8.6cm）
    slbl(doc, 'A（1×10=10点）　※完全正答のみ得点')
    NW_A = 0.7; AW_A = W/2 - NW_A
    # 解答モードは atLeast（長文でも折り返せる）、空欄モードは exact
    h_rule_3a = 'atLeast' if filled else 'exact'
    t3a = mktbl(doc, 5, 4, [NW_A, AW_A, NW_A, AW_A],
                [0.82]*5, h_rule=h_rule_3a)
    for i in range(5):
        for s in range(2):
            q = i*2+s+1
            cw(t3a.cell(i, s*2), f'({q})', 8, True); sh(t3a.cell(i, s*2))
            cw(t3a.cell(i, s*2+1), '', 9)
            if filled:
                cw(t3a.cell(i, s*2+1), ANSWERS['3a'][q-1],
                   8, True, align=WD_ALIGN_PARAGRAPH.LEFT, color=ANS_COLOR)

    # B 空所補充：2列レイアウト（各語ボックス 2.8cm）
    slbl(doc, 'B（1×12=12点）')
    b_blanks = [1,2,2,3,3,3, 1,2,2,2,2,1]
    bx = (W/2 - 0.6) / 3
    t3b = mktbl(doc, 6, 8, [0.6,bx,bx,bx, 0.6,bx,bx,bx], [0.66]*6)
    for i in range(6):
        for side in range(2):
            qi = i+side*6; bc = side*4; nb = b_blanks[qi]
            cw(t3b.cell(i,bc), f'({qi+1})', 8, True); sh(t3b.cell(i,bc))
            for b in range(3):
                c = bc+1+b
                if b < nb:
                    cw(t3b.cell(i,c), '')
                    if filled:
                        words = ANSWERS['3b'][qi]
                        cw(t3b.cell(i,c), words[b] if b<len(words) else '',
                           8, True, color=ANS_COLOR)
                else:
                    no_b(t3b.cell(i,c))

    # C 動詞活用
    slbl(doc, 'C（1×6=6点）')
    t3c = mktbl(doc, 2, 6, [W/6]*6, [0.42, 0.68])
    for j in range(6):
        cw(t3c.cell(0,j), f'({j+1})', 8, True); sh(t3c.cell(0,j))
        cw(t3c.cell(1,j), '')
        if filled: cw(t3c.cell(1,j), ANSWERS['3c'][j], 8, True, color=ANS_COLOR)

    # ══════════ ページ区切り ══════════
    pgbrk(doc)

    # ────────── §4 Writing 英訳 12点 ──────────
    lbl(doc, '４　（思考・判断・表現　12点）', before=0)
    t4 = mktbl(doc, 6, 2, [0.7, W-0.7], [0.85]*6)
    for i in range(6):
        cw(t4.cell(i,0), f'({i+1})', 8, True); sh(t4.cell(i,0))
        cw(t4.cell(i,1), '', 9)
        if filled:
            cw(t4.cell(i,1), ANSWERS['4'][i], 8, True,
               align=WD_ALIGN_PARAGRAPH.LEFT, color=ANS_COLOR)

    # ────────── §5 Reading Unit1&2 10点 ──────────
    lbl(doc, '５　（思考・判断・表現　10点）')
    slbl(doc, 'A（1×5=5点）　　　　　　　　　　　　B（1×5=5点）', before=1)
    NW5=0.7; AW5=W/2-NW5
    t5 = mktbl(doc, 5, 4, [NW5,AW5,NW5,AW5], [0.80]*5)
    for i in range(5):
        for s in range(2):
            cw(t5.cell(i,s*2), f'({i+1})', 8, True); sh(t5.cell(i,s*2))
            cw(t5.cell(i,s*2+1), '', 9)
            if filled:
                key = '5a' if s==0 else '5b'
                cw(t5.cell(i,s*2+1), ANSWERS[key][i], 8, True,
                   align=WD_ALIGN_PARAGRAPH.LEFT, color=ANS_COLOR)

    # ────────── §6 Ancient Greece 15点 ──────────
    lbl(doc, '６　（思考・判断・表現　15点）')

    # 問1/2/5 選択
    slbl(doc, '問1（1点）　問2（2点）　問5（2点）　　各 A・B・C・D から選択', before=1)
    GW=0.4; LW6=2.8; AW6=1.5
    rest6 = W-(LW6+AW6+GW)*3
    t6s = mktbl(doc, 2, 10, [LW6,AW6,GW,LW6,AW6,GW,LW6,AW6,GW,rest6], [0.42,0.66])
    for idx, (ql, key) in enumerate([('問1（1点）','6_q1'),('問2（2点）','6_q2'),('問5（2点）','6_q5')]):
        bc = idx*3
        cw(t6s.cell(0,bc), ql, 8, True); sh(t6s.cell(0,bc))
        cw(t6s.cell(1,bc), '答え', 7.5)
        cw(t6s.cell(0,bc+1), ''); cw(t6s.cell(1,bc+1), '')
        no_b(t6s.cell(0,bc+2)); no_b(t6s.cell(1,bc+2))
        if filled: cw(t6s.cell(1,bc+1), ANSWERS[key], 9, True, color=ANS_COLOR)
    no_b(t6s.cell(0,9)); no_b(t6s.cell(1,9))

    # 問3 S/A
    slbl(doc, '問3（1×5=5点）　　S＝スパルタ　A＝アテネ', before=2)
    aw6q3 = W/5 - 2.0  # ≈ 1.72cm
    t6q3 = mktbl(doc, 2, 10,
                 [2.0,aw6q3,2.0,aw6q3,2.0,aw6q3,2.0,aw6q3,2.0,aw6q3], [0.42,0.66])
    for j, ql in enumerate(['(i)','(ii)','(iii)','(iv)','(v)']):
        cw(t6q3.cell(0,j*2), ql, 8, True); sh(t6q3.cell(0,j*2))
        cw(t6q3.cell(1,j*2), '')
        cw(t6q3.cell(0,j*2+1), ''); cw(t6q3.cell(1,j*2+1), '')
        if filled: cw(t6q3.cell(1,j*2+1), ANSWERS['6_q3'][j], 9, True, color=ANS_COLOR)

    # 問4 日本語記述
    slbl(doc, '問4（2点）　スパルタの女性がアテネの女性より「自由」を得られた理由（日本語）', before=2)
    t6q4 = mktbl(doc, 2, 1, [W], [0.80,0.80])
    for i in range(2):
        cw(t6q4.cell(i,0), '')
        if filled: cw(t6q4.cell(i,0), ANSWERS['6_q4'][i], 8, True,
                      align=WD_ALIGN_PARAGRAPH.LEFT, color=ANS_COLOR)

    # 問6 日本語訳
    slbl(doc, '問6（3点）　下線部(2)を日本語にしなさい', before=2)
    t6q6 = mktbl(doc, 2, 1, [W], [0.80,0.80])
    for i in range(2):
        cw(t6q6.cell(i,0), '')
        if filled: cw(t6q6.cell(i,0), ANSWERS['6_q6'][i], 8, True,
                      align=WD_ALIGN_PARAGRAPH.LEFT, color=ANS_COLOR)

    # ────────── §7 Book Report 6点 ──────────
    lbl(doc, '７　（思考・判断・表現　6点）　50語以上で書くこと')
    t7 = mktbl(doc, 6, 1, [W], [0.82]*6)
    for i in range(6):
        cw(t7.cell(i,0), '')
        if filled and i < len(ANSWERS['7']):
            cw(t7.cell(i,0), ANSWERS['7'][i], 8, True,
               align=WD_ALIGN_PARAGRAPH.LEFT, color=ANS_COLOR)

    # ────────── 得点集計 ──────────
    lbl(doc, '得点集計', before=5)
    ts = mktbl(doc, 2, 3, [5.5,7.1,6.0], [0.42,0.76])
    for j, t in enumerate(['２・３（知識・技能）　/43',
                            '１・４〜７（思考・判断・表現）　/57',
                            '合計　/100']):
        cw(ts.cell(0,j), t, 8, True); sh(ts.cell(0,j))
        cw(ts.cell(1,j), '')

    doc.save(output_path)
    print(f'Saved: {output_path}')

# ════════════════════════════════════════════
#  実行
# ════════════════════════════════════════════
BASE = '/home/user/Yumetan-Quick-Response/'
build(filled=False, output_path=BASE + '2026年度_論理表現Ⅱ_中間テスト_解答用紙.docx')
build(filled=True,  output_path=BASE + '2026年度_論理表現Ⅱ_中間テスト_解答例.docx')
