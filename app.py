import io
import streamlit as st
import pdfplumber
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

# ============================================================
# ★ 調整パラメータ（フォーマットが変わった場合はここを変更）
# ============================================================
SUBTOTAL_NAMES = {'小 計', '小　計', '計', '合 計', '合　計', '合　　計', '総 合 計'}
SKIP_VALS      = {'数 量 書', '工 種 別 数 量 書', '総 括 数 量 書', '種 別', '内 訳', '工 種',
                  '数    量    書', '工 種 別 内 訳 書', '総　括　書', '数量書', '工種別数量書'}
SUMMARY_KEEP   = {'工事価格 合計', '工事価格　合計', '消費税額', '総 合 計', '総　合　計'}
# ============================================================

# ---------- ユーティリティ ----------

def thin_border():
    s = Side(style='thin')
    return Border(left=s, right=s, top=s, bottom=s)

def apply_cell(cell, value, bold=False, center=False, right=False):
    cell.value = value
    cell.font = Font(name='MS Gothic', size=9, bold=bold)
    halign = 'right' if right else ('center' if center else 'left')
    cell.alignment = Alignment(horizontal=halign, vertical='center', wrap_text=True)
    cell.border = thin_border()

def set_formula(cell, formula, right=False):
    cell.value = formula
    cell.font = Font(name='MS Gothic', size=9)
    cell.alignment = Alignment(horizontal='right' if right else 'center', vertical='center')
    cell.border = thin_border()

def clean_qty(val):
    try:
        f = float(val)
        return int(f) if f == int(f) else f
    except:
        return val

def is_page_num(val):
    if not val: return False
    parts = val.replace(' ', '').split('/')
    return len(parts) == 2 and all(p.isdigit() for p in parts)

# ---------- PDF解析 ----------

def extract_all_pages(pdf_file, progress_cb=None):
    summary_rows, category_rows, quantity_rows = [], [], []
    with pdfplumber.open(pdf_file) as pdf:
        total = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            if progress_cb:
                progress_cb(i / total, f'{i+1} / {total} ページ解析中…')
            text  = page.extract_text() or ''
            first = text.split('\n')[0].strip()
            if   '総' in first and '括' in first: ptype = 'summary'
            elif '工' in first and '種' in first: ptype = 'category'
            elif '数' in first and '量' in first: ptype = 'quantity'
            else: continue
            tables = page.extract_tables()
            if not tables: continue
            for row in tables[0]:
                if not row: continue
                cleaned = [(str(c).replace('\n', ' ').strip() if c else '') for c in row]
                # ページ番号行スキップ
                if is_page_num(cleaned[0]): continue
                if is_page_num(cleaned[1] if len(cleaned) > 1 else ''): continue
                # スキップ対象文字列
                if cleaned[0] in SKIP_VALS: continue
                if len(cleaned) > 1 and cleaned[1] in SKIP_VALS: continue
                if ptype == 'summary':
                    if cleaned[1] in {'内 訳', '内　訳'}: continue
                    if (cleaned[1] and not cleaned[0] and not cleaned[2]
                            and cleaned[1] not in SUMMARY_KEEP): continue
                if ptype == 'category':
                    if cleaned[1] in {'種 別', '種　別'}: continue
                    if (cleaned[1] and not cleaned[0] and not cleaned[2]
                            and cleaned[1] not in SUBTOTAL_NAMES): continue
                if   ptype == 'summary':  summary_rows.append(cleaned)
                elif ptype == 'category': category_rows.append(cleaned)
                elif ptype == 'quantity': quantity_rows.append(cleaned)
    return summary_rows, category_rows, quantity_rows

# ---------- 数量書 ----------

def build_quantity_sheet(ws, rows):
    ws.title = '数量書'
    for col, w in zip('ABCDEFGH', [8, 42, 28, 10, 8, 14, 14, 28]):
        ws.column_dimensions[col].width = w

    data = [row for row in rows if len(row) >= 2 and not all(v == '' for v in row)]

    # ── ヘッダー（1ページ分のみ、以降は繰り返さない） ──
    r = 1
    ws.merge_cells(f'A{r}:H{r}')
    ws.cell(r,1).value = '数    量    書'
    ws.cell(r,1).font  = Font(name='MS Gothic', size=11, bold=True)
    ws.cell(r,1).alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[r].height = 20
    r += 1

    h1 = r
    ws.cell(h1,1).border = thin_border()
    ws.cell(h1,2).value = '種　　別'
    ws.cell(h1,2).font  = Font(name='MS Gothic', size=9)
    ws.cell(h1,2).alignment = Alignment(horizontal='center', vertical='center')
    ws.cell(h1,2).border = thin_border()
    for col, h in zip([3,4,5,6,7,8], ['形状・寸法','数  量','単 位','単　価','金　額','摘      要']):
        ws.cell(h1,col).value = h
        ws.cell(h1,col).font  = Font(name='MS Gothic', size=9)
        ws.cell(h1,col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.cell(h1,col).border = thin_border()
    ws.row_dimensions[h1].height = 30
    r += 1

    h2 = r
    ws.cell(h2,1).border = thin_border()
    ws.cell(h2,2).value = '内　　訳'
    ws.cell(h2,2).font  = Font(name='MS Gothic', size=9)
    ws.cell(h2,2).alignment = Alignment(horizontal='left', vertical='center')
    ws.cell(h2,2).border = thin_border()
    for col in range(3, 9):
        ws.merge_cells(f'{chr(64+col)}{h1}:{chr(64+col)}{h2}')
    ws.row_dimensions[h2].height = 20
    r += 1

    # ── データ行（ページ区切りなし） ──
    subtotal_map      = {}
    current_big       = None
    current_sub       = None
    section_data_rows = []

    for vals in data:
        vals = (vals + [''] * 8)[:8]
        num, name, shape, qty, unit, price, amt, note = vals
        bold        = bool(num)
        is_subtotal = name in SUBTOTAL_NAMES

        if num.startswith('【'):
            current_big = num; current_sub = None; section_data_rows = []
        elif num:
            current_sub = num; section_data_rows = []

        apply_cell(ws.cell(r,1), num,   bold=bold, center=True)
        apply_cell(ws.cell(r,2), name,  bold=bold or is_subtotal)
        apply_cell(ws.cell(r,3), shape)
        apply_cell(ws.cell(r,4), clean_qty(qty) if qty else '', center=True)
        apply_cell(ws.cell(r,5), unit,  center=True)
        apply_cell(ws.cell(r,6), '',    center=True)

        if qty and unit and not is_subtotal and not bold:
            set_formula(ws.cell(r,7), f'=INT(F{r}*D{r})')
            section_data_rows.append(r)
        elif is_subtotal:
            if len(section_data_rows) == 1:
                set_formula(ws.cell(r,7), f'=G{section_data_rows[0]}')
            elif len(section_data_rows) > 1:
                set_formula(ws.cell(r,7), f'=SUM(G{section_data_rows[0]}:G{section_data_rows[-1]})')
            else:
                apply_cell(ws.cell(r,7), '', center=True)
            subtotal_map[(current_big, current_sub)] = r
            section_data_rows = []
        else:
            apply_cell(ws.cell(r,7), '', center=True)

        apply_cell(ws.cell(r,8), note)
        ws.row_dimensions[r].height = 25  # ★ 全データ行25pt均一
        r += 1

    return subtotal_map

# ---------- 工事別数量書 ----------

def build_category_sheet(ws, rows, subtotal_map):
    ws.title = '工事別数量書'
    ws.column_dimensions['A'].width = 6.6
    ws.column_dimensions['B'].width = 40.5
    ws.column_dimensions['C'].width = 24.5
    ws.column_dimensions['D'].width = 22.6
    ws.column_dimensions['E'].width = 26.6
    ws.column_dimensions['F'].width = 4.2

    data = [row for row in rows
            if len(row) >= 2 and not all(v == '' for v in row) and not is_page_num(row[0])]

    # ── ヘッダー（1ページ分のみ） ──
    r = 1
    ws.cell(r,1).border = thin_border()
    ws.merge_cells(f'B{r}:F{r}')
    ws.cell(r,2).value = '　　　　　　　　　　　　　　　　　　工 種 別 内 訳 書 '
    ws.cell(r,2).font  = Font(name='MS Gothic', size=10, bold=True)
    ws.cell(r,2).alignment = Alignment(horizontal='left', vertical='center')
    ws.cell(r,2).border = thin_border()
    ws.row_dimensions[r].height = 20
    r += 1

    h1 = r
    for col, h in [(1,None),(2,'工    種'),(3,'内  容 (数量)'),(4,'金　額 （円）'),(5,' 摘      要')]:
        ws.cell(h1,col).value = h
        ws.cell(h1,col).font  = Font(name='MS Gothic', size=9)
        ws.cell(h1,col).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(h1,col).border = thin_border()
    ws.row_dimensions[h1].height = 16
    r += 1

    h2 = r
    ws.cell(h2,1).border = thin_border()
    ws.cell(h2,2).value = '種    別'
    ws.cell(h2,2).font  = Font(name='MS Gothic', size=9)
    ws.cell(h2,2).alignment = Alignment(horizontal='center', vertical='center')
    ws.cell(h2,2).border = thin_border()
    ws.merge_cells(f'C{h1}:C{h2}')
    ws.merge_cells(f'D{h1}:D{h2}')
    ws.merge_cells(f'E{h1}:F{h2}')
    ws.cell(h2,3).border = thin_border()
    ws.cell(h2,4).border = thin_border()
    ws.cell(h2,5).border = thin_border()
    ws.row_dimensions[h2].height = 16
    r += 1

    # ── データ行（ページ区切りなし） ──
    current_big  = None
    section_rows = []   # 「計」(直接工事費小計)用：1-1〜1-n の金額行
    grand_rows   = []   # 「合 計」用：計＋共通仮設費＋現場管理費＋一般管理費等 の行

    for row in data:
        vals = (row + [''] * 5)[:5]
        num, name, qty, amt, note = vals
        bold        = bool(num)
        is_subtotal = name in SUBTOTAL_NAMES
        is_total    = name in {'合 計', '合　計', '合　　計'}

        if num.startswith('【'):
            current_big = num; section_rows = []; grand_rows = []

        apply_cell(ws.cell(r,1), num,  bold=bold, center=True)
        apply_cell(ws.cell(r,2), name, bold=bold or is_subtotal)
        apply_cell(ws.cell(r,3), qty,  center=True)

        if qty == '一式' and num and not num.startswith('【') and not is_subtotal:
            # 直接工事費の細目（1-1等）→数量書の小計を参照
            ref = subtotal_map.get((current_big, num))
            if ref:
                set_formula(ws.cell(r,4), f'=数量書!G{ref}', right=True)
                section_rows.append(r)
            else:
                apply_cell(ws.cell(r,4), '', right=True)
            ws.merge_cells(f'E{r}:F{r}')
            apply_cell(ws.cell(r,5), note)
            ws.cell(r,6).border = thin_border()
            # 共通仮設費(2)/現場管理費(3)/一般管理費等(4)/有価発生材(5)は合計対象
            if '-' not in num:
                grand_rows.append(r)
        elif is_total:
            # 「合 計」＝ 計＋共通仮設費＋現場管理費＋一般管理費等…（連続行をSUM）
            if grand_rows:
                set_formula(ws.cell(r,4),
                            f'=SUM(D{grand_rows[0]}:D{grand_rows[-1]})', right=True)
            else:
                apply_cell(ws.cell(r,4), '', right=True)
            ws.merge_cells(f'E{r}:F{r}')
            apply_cell(ws.cell(r,5), '')
            ws.cell(r,6).border = thin_border()
            grand_rows = []
        elif is_subtotal and section_rows:
            # 「計」＝ 直接工事費（1-1〜1-n）の合算
            formula = (f'=D{section_rows[0]}' if len(section_rows) == 1
                       else f'=SUM(D{section_rows[0]}:D{section_rows[-1]})')
            set_formula(ws.cell(r,4), formula, right=True)
            ws.merge_cells(f'E{r}:F{r}')
            apply_cell(ws.cell(r,5), '')
            ws.cell(r,6).border = thin_border()
            grand_rows.append(r)  # 「計」も合計対象に含める
            section_rows = []
        else:
            apply_cell(ws.cell(r,4), '', right=True)
            ws.merge_cells(f'E{r}:F{r}')
            apply_cell(ws.cell(r,5), note)
            ws.cell(r,6).border = thin_border()

        ws.row_dimensions[r].height = 25  # ★ 全データ行25pt均一
        r += 1

    return ws

# ---------- 総括数量書 ----------

def build_summary_sheet(ws, rows, ws_cat):
    ws.title = '総括数量書'
    ws.column_dimensions['A'].width = 6.6
    ws.column_dimensions['B'].width = 25.7
    ws.column_dimensions['C'].width = 16.7
    ws.column_dimensions['D'].width = 25.7
    ws.column_dimensions['E'].width = 20.7
    ws.column_dimensions['F'].width = 20.7
    ws.column_dimensions['G'].width = 4.2

    # ── ヘッダー ──
    ws.merge_cells('A1:G1')
    ws['A1'].value = '総　括　書'
    ws['A1'].font  = Font(name='MS Gothic', size=11)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 32

    ws.cell(2,1).border = thin_border()
    ws.merge_cells('B2:C2')
    ws.cell(2,2).value = '種     別'
    ws.cell(2,2).font  = Font(name='MS Gothic', size=9)
    ws.cell(2,2).alignment = Alignment(horizontal='center', vertical='center')
    ws.cell(2,2).border = thin_border()
    ws.cell(2,4).value = '  内  容 (数量)'
    ws.cell(2,4).font  = Font(name='MS Gothic', size=9)
    ws.cell(2,4).alignment = Alignment(horizontal='center', vertical='center')
    ws.cell(2,4).border = thin_border()
    ws.cell(2,5).value = '金　額 （円）'
    ws.cell(2,5).font  = Font(name='MS Gothic', size=9)
    ws.cell(2,5).alignment = Alignment(horizontal='center', vertical='center')
    ws.cell(2,5).border = thin_border()
    ws.row_dimensions[2].height = 16

    ws.cell(3,1).border = thin_border()
    ws.merge_cells('B3:C3')
    ws.cell(3,2).value = '内     訳'
    ws.cell(3,2).font  = Font(name='MS Gothic', size=9)
    ws.cell(3,2).alignment = Alignment(horizontal='center', vertical='center')
    ws.cell(3,2).border = thin_border()
    ws.cell(3,4).border = thin_border()
    ws.cell(3,5).border = thin_border()
    ws.row_dimensions[3].height = 16

    ws.merge_cells('D2:D3')
    ws.merge_cells('E2:E3')
    ws.merge_cells('F2:G3')
    ws.cell(2,6).value = '摘　　要'
    ws.cell(2,6).font  = Font(name='MS Gothic', size=9)
    ws.cell(2,6).alignment = Alignment(horizontal='center', vertical='center')
    ws.cell(2,6).border = thin_border()

    ws.merge_cells('B4:G4')
    ws.cell(4,1).border = thin_border()
    ws.cell(4,2).border = thin_border()
    ws.row_dimensions[4].height = 25

    # 工事別数量書の「合 計」行番号をマップ
    cat_total_map = {}
    big_code = None
    for i, row in enumerate(ws_cat.iter_rows(min_row=4, values_only=True), 4):
        if not row: continue
        code = str(row[0]).strip() if row[0] else ''
        name = str(row[1]).strip() if row[1] else ''
        if code.startswith('【'): big_code = code
        if name in {'合 計', '合　計', '合　　計'} and big_code:
            cat_total_map[big_code] = i

    r = 5
    item_rows = []
    data = [row for row in rows
            if len(row) >= 2 and not all(v == '' for v in row) and not is_page_num(row[0])]

    for vals in data:
        vals = (vals + [''] * 5)[:5]
        num, name, qty, amt, note = vals
        bold = bool(num)

        apply_cell(ws.cell(r,1), num, bold=bold, center=True)
        ws.merge_cells(f'B{r}:C{r}')
        apply_cell(ws.cell(r,2), name, bold=bold)
        ws.cell(r,3).border = thin_border()
        apply_cell(ws.cell(r,4), qty, center=True)

        # 工事番号は【1】形式。そのままキーとして工事別の合計行を参照
        if num.startswith('【') and qty == '一式':
            ref = cat_total_map.get(num)
            if ref: set_formula(ws.cell(r,5), f'=工事別数量書!D{ref}', right=True)
            else:   apply_cell(ws.cell(r,5), '', right=True)
            item_rows.append(r)
        elif name in {'工事価格 合計', '工事価格　合計'} and item_rows:
            set_formula(ws.cell(r,5), f'=SUM(E{item_rows[0]}:E{item_rows[-1]})', right=True)
        elif name == '消費税額':
            set_formula(ws.cell(r,5), f'=INT(E{r-1}*0.1)', right=True)
        elif name in {'総 合 計', '総　合　計'}:
            set_formula(ws.cell(r,5), f'=E{r-2}+E{r-1}', right=True)
        else:
            apply_cell(ws.cell(r,5), '', right=True)

        ws.merge_cells(f'F{r}:G{r}')
        apply_cell(ws.cell(r,6), note)
        ws.cell(r,7).border = thin_border()
        ws.row_dimensions[r].height = 25  # ★ 全データ行25pt均一
        r += 1

# ---------- Excel生成 ----------

def build_excel(summary_rows, category_rows, quantity_rows):
    wb = openpyxl.Workbook()
    ws_quantity  = wb.active
    subtotal_map = build_quantity_sheet(ws_quantity, quantity_rows)
    ws_category  = wb.create_sheet()
    build_category_sheet(ws_category, category_rows, subtotal_map)
    ws_summary   = wb.create_sheet()
    build_summary_sheet(ws_summary, summary_rows, ws_category)
    wb.move_sheet(ws_summary, offset=-wb.index(ws_summary))
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ============================================================
# Streamlit UI
# ============================================================

st.set_page_config(page_title='工事数量書 変換ツール', page_icon='📄', layout='centered')

st.title('📄 工事数量書 PDF → Excel 変換')
st.caption('JKK工事数量書のPDFをアップロードするとExcelに変換します')
st.divider()

uploaded = st.file_uploader('PDFファイルを選択', type='pdf', label_visibility='collapsed')

if uploaded:
    st.info(f'ファイル: **{uploaded.name}**　({uploaded.size / 1024:.0f} KB)')

    if st.button('変換開始', type='primary', use_container_width=True):
        progress = st.progress(0, text='準備中…')
        status   = st.empty()

        try:
            def cb(pct, msg):
                progress.progress(pct, text=msg)

            summary_rows, category_rows, quantity_rows = extract_all_pages(uploaded, cb)

            progress.progress(0.9, text='Excel生成中…')
            excel_buf = build_excel(summary_rows, category_rows, quantity_rows)

            progress.progress(1.0, text='完了！')
            status.success(
                f'変換完了 ✅　'
                f'総括: {len(summary_rows)}行 ／ '
                f'工事別: {len(category_rows)}行 ／ '
                f'数量: {len(quantity_rows)}行'
            )

            out_name = uploaded.name.replace('.pdf', '.xlsx').replace('.PDF', '.xlsx')
            st.download_button(
                label='📥 Excelをダウンロード',
                data=excel_buf,
                file_name=out_name,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True,
                type='primary',
            )

        except Exception as e:
            progress.empty()
            st.error(f'エラーが発生しました: {e}')

st.divider()
st.caption('対応フォーマット: 総括数量書 / 工事別数量書 / 数量書 の3シート構成PDF')
