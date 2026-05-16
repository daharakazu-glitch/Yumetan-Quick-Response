from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# Page margins
section = doc.sections[0]
section.page_width = Cm(21)
section.page_height = Cm(29.7)
section.left_margin = Cm(1.5)
section.right_margin = Cm(1.5)
section.top_margin = Cm(1.5)
section.bottom_margin = Cm(1.5)

def set_font(run, size=10, bold=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = 'MS Gothic'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'MS Gothic')

def add_heading(doc, text, size=11, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER):
    p = doc.add_paragraph()
    p.alignment = align
    run = p.add_run(text)
    set_font(run, size, bold)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    return p

def set_cell_border(cell, top=True, bottom=True, left=True, right=True, size=4):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    sides = []
    if top: sides.append('top')
    if bottom: sides.append('bottom')
    if left: sides.append('left')
    if right: sides.append('right')
    for side in sides:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), str(size))
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), '000000')
        tcBorders.append(el)
    tcPr.append(tcBorders)

def set_cell_shading(cell, fill='F2F2F2'):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill)
    tcPr.append(shd)

def set_row_height(row, height_cm):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), str(int(height_cm * 567)))
    trHeight.set(qn('w:hRule'), 'exact')
    trPr.append(trHeight)

def cell_text(cell, text, size=9, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, valign=WD_ALIGN_VERTICAL.CENTER):
    cell.vertical_alignment = valign
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(text)
    set_font(run, size, bold)

def add_section_label(doc, num, title, points, is_thought=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    label = f'{num}　（{"思考・判断・表現" if is_thought else "知識・技能"}　{points}点）'
    run = p.add_run(label)
    set_font(run, 9, True)

# ============================================================
# HEADER
# ============================================================
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(2)
run = p.add_run('2026年度　山形県立山形南高等学校　2学年1学期中間テスト　論理・表現Ⅱ　解答用紙')
set_font(run, 11, True)

# Class / No / Name table
t = doc.add_table(rows=2, cols=4)
t.alignment = WD_TABLE_ALIGNMENT.CENTER
t.style = 'Table Grid'
widths = [Cm(2), Cm(3), Cm(2), Cm(8)]
for i, row in enumerate(t.rows):
    for j, cell in enumerate(row.cells):
        cell.width = widths[j]
        set_row_height(row, 0.7)

cell_text(t.cell(0,0), 'Class', 9, True, WD_ALIGN_PARAGRAPH.CENTER)
cell_text(t.cell(0,2), 'No', 9, True, WD_ALIGN_PARAGRAPH.CENTER)
cell_text(t.cell(1,0), 'Name', 9, True, WD_ALIGN_PARAGRAPH.CENTER)

# Merge Name cell across columns
t.cell(1,1).merge(t.cell(1,3))

doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ============================================================
# SECTION 1: Listening (14点)
# ============================================================
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(2)
p.paragraph_format.space_after = Pt(1)
run = p.add_run('１　Listening　（14点）')
set_font(run, 10, True)

# 1A and 1B side by side
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(1)
run = p.add_run('A　（2点）　　　　　　　　　　　　　　　　　　B　（2点）')
set_font(run, 9)

t = doc.add_table(rows=1, cols=6)
t.style = 'Table Grid'
t.alignment = WD_TABLE_ALIGNMENT.LEFT
labels_a = ['①', '②', '③']
labels_b = ['①', '②', '③']
col_widths = [Cm(1.5), Cm(2), Cm(1.5), Cm(1), Cm(1.5), Cm(2)]
for j, cell in enumerate(t.rows[0].cells):
    cell.width = col_widths[j]
set_row_height(t.rows[0], 0.8)

cell_text(t.cell(0,0), 'A の答え', 8, False, WD_ALIGN_PARAGRAPH.CENTER)
cell_text(t.cell(0,1), '', 9)
cell_text(t.cell(0,2), 'B の答え', 8, False, WD_ALIGN_PARAGRAPH.CENTER)
cell_text(t.cell(0,3), '', 9)
# Add memo note
cell_text(t.cell(0,4), '', 8)
cell_text(t.cell(0,5), '', 8)

# simpler approach
t2 = doc.add_table(rows=1, cols=4)
t2.style = 'Table Grid'
t2.alignment = WD_TABLE_ALIGNMENT.LEFT
set_row_height(t2.rows[0], 0.8)
cw = [Cm(3), Cm(2.5), Cm(3), Cm(2.5)]
for j, c in enumerate(t2.rows[0].cells):
    c.width = cw[j]
cell_text(t2.cell(0,0), 'A（①②③から選択）', 8, False, WD_ALIGN_PARAGRAPH.CENTER)
cell_text(t2.cell(0,1), '', 9)
cell_text(t2.cell(0,2), 'B（①②③から選択）', 8, False, WD_ALIGN_PARAGRAPH.CENTER)
cell_text(t2.cell(0,3), '', 9)

# Remove the first table attempt
t._element.getparent().remove(t._element)

# 1C: fill in the blanks
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(3)
p.paragraph_format.space_after = Pt(1)
run = p.add_run('C　ニュースを聞いて空欄を埋めなさい。（2×5=10点）')
set_font(run, 9, True)

# Table for 1C with the sentence structure
tc = doc.add_table(rows=5, cols=5)
tc.style = 'Table Grid'
tc.alignment = WD_TABLE_ALIGNMENT.LEFT

row_data = [
    ('(1)', 'Many objects (', ')(', ')(', ') into space.'),
    ('(2)', 'Some of these objects are (', ')(', ') in use but they continue to float around the Earth.', ''),
    ('(3)', 'Space debris (', ') satellites that are not working and broken pieces of spacecraft.', '', ''),
    ('(4)', 'It will (', ') future space missions (', ')(', ').'),
    ('(5)', 'A speedy solution is required (', ')(', ')(', ') ensure the safety of space exploration.'),
]

for i, (num, *parts) in enumerate(row_data):
    row = tc.rows[i]
    set_row_height(row, 0.9)
    cell_text(row.cells[0], num, 8, True, WD_ALIGN_PARAGRAPH.CENTER)
    row.cells[0].width = Cm(0.8)
    for j, part in enumerate(parts):
        if j+1 < len(row.cells):
            cell_text(row.cells[j+1], part, 8)

doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ============================================================
# SECTION 2: LEAP Vocabulary (15点)
# ============================================================
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(2)
p.paragraph_format.space_after = Pt(1)
run = p.add_run('２　LEAP Vocabulary　（15点）')
set_font(run, 10, True)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(1)
run = p.add_run('A　適語選択（1×10=10点）')
set_font(run, 9, True)

# 2A: 10 answer boxes in 2 rows of 5
t2a = doc.add_table(rows=2, cols=10)
t2a.style = 'Table Grid'
t2a.alignment = WD_TABLE_ALIGNMENT.LEFT
for i, row in enumerate(t2a.rows):
    set_row_height(row, 0.8)
    for j, cell in enumerate(row.cells):
        cell.width = Cm(1.5)
        num = i*5 + j + 1
        cell_text(cell, f'({num})', 8, False, WD_ALIGN_PARAGRAPH.CENTER)

# Add answer row
t2a2 = doc.add_table(rows=1, cols=10)
t2a2.style = 'Table Grid'
t2a2.alignment = WD_TABLE_ALIGNMENT.LEFT
set_row_height(t2a2.rows[0], 0.8)
for j, cell in enumerate(t2a2.rows[0].cells):
    cell.width = Cm(1.5)
    cell_text(cell, '', 9)

# merge into one proper table
# Actually let's do it as a simple 2-row label + answer table
t2a._element.getparent().remove(t2a._element)
t2a2._element.getparent().remove(t2a2._element)

t2a_final = doc.add_table(rows=2, cols=10)
t2a_final.style = 'Table Grid'
t2a_final.alignment = WD_TABLE_ALIGNMENT.LEFT
for j in range(10):
    t2a_final.cell(0, j).width = Cm(1.5)
    t2a_final.cell(1, j).width = Cm(1.5)
    cell_text(t2a_final.cell(0, j), f'({j+1})', 8, False, WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(t2a_final.cell(1, j), '', 9)
    set_row_height(t2a_final.rows[0], 0.5)
    set_row_height(t2a_final.rows[1], 0.8)
for row in t2a_final.rows:
    for cell in row.cells:
        set_cell_border(cell)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(3)
p.paragraph_format.space_after = Pt(1)
run = p.add_run('B　空欄補充（1×5=5点）')
set_font(run, 9, True)

t2b = doc.add_table(rows=2, cols=5)
t2b.style = 'Table Grid'
t2b.alignment = WD_TABLE_ALIGNMENT.LEFT
for j in range(5):
    t2b.cell(0, j).width = Cm(3)
    t2b.cell(1, j).width = Cm(3)
    cell_text(t2b.cell(0, j), f'({j+1})', 8, False, WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(t2b.cell(1, j), '', 9)
    set_row_height(t2b.rows[0], 0.5)
    set_row_height(t2b.rows[1], 0.8)

doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ============================================================
# SECTION 3: Grammar and Usage (28点)
# ============================================================
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(2)
p.paragraph_format.space_after = Pt(1)
run = p.add_run('３　Grammar and Usage　（28点）')
set_font(run, 10, True)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(1)
run = p.add_run('A　語句整序（1×10=10点）　※完全正答のみ得点')
set_font(run, 9, True)

t3a = doc.add_table(rows=10, cols=2)
t3a.style = 'Table Grid'
t3a.alignment = WD_TABLE_ALIGNMENT.LEFT
questions_3a = [
    '1. 非常に多くの人が毎年その国を訪れます。',
    '2. 中国はその巨大な建造物でも有名です。',
    '3. アイスランドの夏は過ごしやすいです。',
    '4. 世界地図上でその国が巨大な猫のように見えるのが面白い。',
    '5. 世界には探究すべき新しい場所が多すぎる。',
    '6. 著者は庭の樹齢100年の木に着想を得た。',
    '7. かなり怖いですが，あと少しの恐怖が話をよりいっそうおもしろくしていたでしょう。',
    '8. 日本ではゲストをもてなすことがとても大切です。',
    '9. これらの恐竜の化石がどのように発見されたかを楽しく読んでみてください。',
    '10. その魔法の街で彼女は偉大な魔法使いとして尊敬されている。',
]
for i, q in enumerate(questions_3a):
    set_row_height(t3a.rows[i], 0.85)
    t3a.cell(i, 0).width = Cm(5)
    t3a.cell(i, 1).width = Cm(10.5)
    cell_text(t3a.cell(i, 0), q, 7.5)
    cell_text(t3a.cell(i, 1), '', 9)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(3)
p.paragraph_format.space_after = Pt(1)
run = p.add_run('B　空所補充（1×12=12点）')
set_font(run, 9, True)

# 3B questions
b_questions = [
    ('1.', '(　　　　) have a lot of snow in winter in this area.'),
    ('2.', '(　　　　) is interesting (　　　　) learn about other cultures.'),
    ('3.', '(　　　　) (　　　　) more than 500 species of mammal in Indonesia.'),
    ('4.', 'The scholarship (　　　　) (　　　　) (　　　　) go on to college.'),
    ('5.', 'Illness (　　　　) me (　　　　) (　　　　) to school.'),
    ('6.','The Little Prince (　　　　) (　　　　) (　　　　) into over 300 languages.'),
    ('7.', 'The book (　　　　) over 60 pages of fascinating facts about Japan.'),
    ('8.', 'be (　　　　) in the (　　　　) of the city'),
    ('9.', 'The (　　　　) inside the area (　　　　) due to its specific landform.'),
    ('10.', 'Industrial waste is a major cause of (　　　　) (　　　　).'),
    ('11.', 'Japan is rapidly becoming an (　　　　) (　　　　)'),
    ('12.', 'Our company makes good use of (　　　　) technology to improve product quality.'),
]
t3b = doc.add_table(rows=12, cols=2)
t3b.style = 'Table Grid'
t3b.alignment = WD_TABLE_ALIGNMENT.LEFT
for i, (num, sent) in enumerate(b_questions):
    set_row_height(t3b.rows[i], 0.85)
    t3b.cell(i, 0).width = Cm(1)
    t3b.cell(i, 1).width = Cm(14.5)
    cell_text(t3b.cell(i, 0), num, 8, True, WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(t3b.cell(i, 1), sent, 8)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(3)
p.paragraph_format.space_after = Pt(1)
run = p.add_run('C　動詞の活用（1×6=6点）')
set_font(run, 9, True)

c_questions = [
    ('1.', 'Sandy (　　　　　　　　　　) in the library when I saw her two hours ago.　[study]'),
    ('2.', 'We (　　　　　　　　　　) to the countryside next month.　[move]'),
    ('3.', 'In some areas, roads (　　　　　　　　　　) since yesterday.　[close]'),
    ('4.', 'Traditionally, the Tour de France (　　　　　　　　　　) off in July.　[kick]'),
    ('5.', 'The earth (　　　　　　　　　　) hotter every year.　[get]'),
    ('6.', 'The Tyrannosaurus (　　　　　　　　　　) on the continent about 70 million years ago.　[live]'),
]
t3c = doc.add_table(rows=6, cols=2)
t3c.style = 'Table Grid'
t3c.alignment = WD_TABLE_ALIGNMENT.LEFT
for i, (num, sent) in enumerate(c_questions):
    set_row_height(t3c.rows[i], 0.85)
    t3c.cell(i, 0).width = Cm(1)
    t3c.cell(i, 1).width = Cm(14.5)
    cell_text(t3c.cell(i, 0), num, 8, True, WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(t3c.cell(i, 1), sent, 8)

doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ============================================================
# SECTION 4: Writing - Translation (12点)
# ============================================================
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(2)
p.paragraph_format.space_after = Pt(1)
run = p.add_run('４　Writing（12点）　英訳しなさい。（2×6）')
set_font(run, 10, True)

trans_qs = [
    '1. 彼のスピーチはアラン(Alan)のスピーチほど面白く(interesting)なかった。',
    '2. 私の携帯電話(cell phone)は彼女のものほど新しくない。',
    '3. ユキほど有能な(capable)秘書(secretary)はいない。',
    '4. その男性が言ったことは、真実にはほど遠かった(be far from)。',
    '5. 主人公は見知らぬ人(stranger)に駅で話しかけられた(speak to)。',
    '6. その絵は1956年から防弾ガラス(bullet-proof glass)で保護されてきた(protect)。',
]
t4 = doc.add_table(rows=6, cols=2)
t4.style = 'Table Grid'
t4.alignment = WD_TABLE_ALIGNMENT.LEFT
for i, q in enumerate(trans_qs):
    set_row_height(t4.rows[i], 1.2)
    t4.cell(i, 0).width = Cm(5.5)
    t4.cell(i, 1).width = Cm(10)
    cell_text(t4.cell(i, 0), q, 7.5)
    cell_text(t4.cell(i, 1), '', 9)

doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ============================================================
# SECTION 5: Reading Comprehension Unit 1 & 2 (10点)
# ============================================================
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(2)
p.paragraph_format.space_after = Pt(1)
run = p.add_run('５　Reading Comprehension: Textbook Unit 1 and Unit 2 Model　（10点）')
set_font(run, 10, True)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(1)
run = p.add_run('A　プレゼンテーションを読んで答えなさい。（1×5=5点）')
set_font(run, 9, True)

ra_qs = [
    'Q1. What does Emily love?',
    'Q2. How many species of birds are there in Brunei?',
    'Q3. What can you see in Brunei?',
    'Q4. Where is Brunei?',
    'Q5. About how many mammals are there in the rainforests of Brunei?',
]
t5a = doc.add_table(rows=5, cols=2)
t5a.style = 'Table Grid'
t5a.alignment = WD_TABLE_ALIGNMENT.LEFT
for i, q in enumerate(ra_qs):
    set_row_height(t5a.rows[i], 1.0)
    t5a.cell(i, 0).width = Cm(5)
    t5a.cell(i, 1).width = Cm(10.5)
    cell_text(t5a.cell(i, 0), q, 8)
    cell_text(t5a.cell(i, 1), '', 9)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(3)
p.paragraph_format.space_after = Pt(1)
run = p.add_run('B　ブックレポートを読んで答えなさい。（1×5=5点）')
set_font(run, 9, True)

rb_qs = [
    'Q1. What is the book about?',
    'Q2. What can you learn by reading the book?',
    'Q3. Why does Mark think older students can also really appreciate the book?',
    'Q4. Who is the author of the book Wonder?',
    'Q5. Why did the boy\'s classmates like him over time?',
]
t5b = doc.add_table(rows=5, cols=2)
t5b.style = 'Table Grid'
t5b.alignment = WD_TABLE_ALIGNMENT.LEFT
for i, q in enumerate(rb_qs):
    set_row_height(t5b.rows[i], 1.0)
    t5b.cell(i, 0).width = Cm(6)
    t5b.cell(i, 1).width = Cm(9.5)
    cell_text(t5b.cell(i, 0), q, 8)
    cell_text(t5b.cell(i, 1), '', 9)

doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ============================================================
# SECTION 6: Reading Comprehension Ancient Greece (15点)
# ============================================================
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(2)
p.paragraph_format.space_after = Pt(1)
run = p.add_run('６　Reading Comprehension: Ancient Greece　（15点）')
set_font(run, 10, True)

# Q1
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(1)
run = p.add_run('問1　"power over"の意味（1点）')
set_font(run, 9, True)

t6q1 = doc.add_table(rows=1, cols=2)
t6q1.style = 'Table Grid'
t6q1.alignment = WD_TABLE_ALIGNMENT.LEFT
set_row_height(t6q1.rows[0], 0.8)
t6q1.cell(0,0).width = Cm(4)
t6q1.cell(0,1).width = Cm(11.5)
cell_text(t6q1.cell(0,0), '選択（A・B・C・D）', 8, False, WD_ALIGN_PARAGRAPH.CENTER)
cell_text(t6q1.cell(0,1), '', 9)

# Q2
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(3)
p.paragraph_format.space_after = Pt(1)
run = p.add_run('問2　アテネの教育について正しい記述（2点）')
set_font(run, 9, True)

t6q2 = doc.add_table(rows=1, cols=2)
t6q2.style = 'Table Grid'
t6q2.alignment = WD_TABLE_ALIGNMENT.LEFT
set_row_height(t6q2.rows[0], 0.8)
t6q2.cell(0,0).width = Cm(4)
t6q2.cell(0,1).width = Cm(11.5)
cell_text(t6q2.cell(0,0), '選択（A・B・C・D）', 8, False, WD_ALIGN_PARAGRAPH.CENTER)
cell_text(t6q2.cell(0,1), '', 9)

# Q3
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(3)
p.paragraph_format.space_after = Pt(1)
run = p.add_run('問3　SはスパルタAはアテネ（1×5=5点）')
set_font(run, 9, True)

t6q3 = doc.add_table(rows=2, cols=5)
t6q3.style = 'Table Grid'
t6q3.alignment = WD_TABLE_ALIGNMENT.LEFT
labels_q3 = ['(i)', '(ii)', '(iii)', '(iv)', '(v)']
for j in range(5):
    t6q3.cell(0,j).width = Cm(3)
    t6q3.cell(1,j).width = Cm(3)
    cell_text(t6q3.cell(0,j), labels_q3[j], 8, False, WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(t6q3.cell(1,j), '', 9)
    set_row_height(t6q3.rows[0], 0.5)
    set_row_height(t6q3.rows[1], 0.8)

# Q4
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(3)
p.paragraph_format.space_after = Pt(1)
run = p.add_run('問4　スパルタの女性がアテネの女性より「自由」を得られた理由（日本語）（2点）')
set_font(run, 9, True)

t6q4 = doc.add_table(rows=3, cols=1)
t6q4.style = 'Table Grid'
t6q4.alignment = WD_TABLE_ALIGNMENT.LEFT
for i in range(3):
    set_row_height(t6q4.rows[i], 0.85)
    t6q4.cell(i,0).width = Cm(15.5)
    cell_text(t6q4.cell(i,0), '', 9)

# Q5
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(3)
p.paragraph_format.space_after = Pt(1)
run = p.add_run('問5　空所[A]に入る表現（2点）')
set_font(run, 9, True)

t6q5 = doc.add_table(rows=1, cols=2)
t6q5.style = 'Table Grid'
t6q5.alignment = WD_TABLE_ALIGNMENT.LEFT
set_row_height(t6q5.rows[0], 0.8)
t6q5.cell(0,0).width = Cm(4)
t6q5.cell(0,1).width = Cm(11.5)
cell_text(t6q5.cell(0,0), '選択（A・B・C・D）', 8, False, WD_ALIGN_PARAGRAPH.CENTER)
cell_text(t6q5.cell(0,1), '', 9)

# Q6
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(3)
p.paragraph_format.space_after = Pt(1)
run = p.add_run('問6　下線部(2)を日本語にしなさい。（3点）')
set_font(run, 9, True)

t6q6 = doc.add_table(rows=3, cols=1)
t6q6.style = 'Table Grid'
t6q6.alignment = WD_TABLE_ALIGNMENT.LEFT
for i in range(3):
    set_row_height(t6q6.rows[i], 0.85)
    t6q6.cell(i,0).width = Cm(15.5)
    cell_text(t6q6.cell(i,0), '', 9)

doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ============================================================
# SECTION 7: Writing - Book Report (6点)
# ============================================================
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(2)
p.paragraph_format.space_after = Pt(1)
run = p.add_run('７　Writing: Book Report　（6点）　（50語以上）')
set_font(run, 10, True)

t7 = doc.add_table(rows=10, cols=1)
t7.style = 'Table Grid'
t7.alignment = WD_TABLE_ALIGNMENT.LEFT
for i in range(10):
    set_row_height(t7.rows[i], 0.85)
    t7.cell(i,0).width = Cm(15.5)
    cell_text(t7.cell(i,0), '', 9)

doc.add_paragraph().paragraph_format.space_after = Pt(4)

# ============================================================
# SCORE TABLE
# ============================================================
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(4)
p.paragraph_format.space_after = Pt(2)
run = p.add_run('得点集計')
set_font(run, 9, True)

ts = doc.add_table(rows=2, cols=9)
ts.style = 'Table Grid'
ts.alignment = WD_TABLE_ALIGNMENT.LEFT
headers = ['１', '２', '３', '４', '５', '６', '７', '小計', '合計']
widths_s = [Cm(1.5), Cm(1.5), Cm(1.5), Cm(1.5), Cm(1.5), Cm(1.5), Cm(1.5), Cm(2.5), Cm(2.5)]
for j, (h, w) in enumerate(zip(headers, widths_s)):
    ts.cell(0,j).width = w
    ts.cell(1,j).width = w
    cell_text(ts.cell(0,j), h, 8, True, WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(ts.cell(1,j), '', 9)
    set_row_height(ts.rows[0], 0.5)
    set_row_height(ts.rows[1], 0.9)

# Save
output_path = '/home/user/Yumetan-Quick-Response/2026年度_論理表現Ⅱ_中間テスト_解答用紙.docx'
doc.save(output_path)
print(f'Saved: {output_path}')
