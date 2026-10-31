# -*- coding: utf-8 -*-
"""Build the concise QR check-in handover manual."""

from pathlib import Path
import re

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "QR報到系統_使用說明書.docx"

BLUE = "2E74B5"
DARK_BLUE = "1F4D78"
NAVY = "0B2545"
MUTED = "5A6673"
LIGHT_BLUE = "E8EEF5"
LIGHT_GRAY = "F4F6F9"
BORDER = "D7E0EA"
WHITE = "FFFFFF"


def set_run_font(run, name="Calibri", size=None, color=None, bold=None, italic=None):
    run.font.name = name
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), name)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), name)
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), name)
    if size is not None:
        run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def set_style_font(style, name="Calibri", size=11, color="000000", bold=None):
    style.font.name = name
    style._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), name)
    style._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), name)
    style._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), name)
    style.font.size = Pt(size)
    style.font.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        style.font.bold = bold


def set_cell_shading(paragraph, fill):
    ppr = paragraph._p.get_or_add_pPr()
    shd = ppr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        ppr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def set_paragraph_border(paragraph, color=BORDER, size="8", space="6"):
    ppr = paragraph._p.get_or_add_pPr()
    borders = ppr.find(qn("w:pBdr"))
    if borders is None:
        borders = OxmlElement("w:pBdr")
        ppr.append(borders)
    for edge in ("top", "left", "bottom", "right"):
        tag = qn(f"w:{edge}")
        border = borders.find(tag)
        if border is None:
            border = OxmlElement(f"w:{edge}")
            borders.append(border)
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), size)
        border.set(qn("w:space"), space)
        border.set(qn("w:color"), color)


def add_callout(doc, label, text, fill=LIGHT_GRAY, label_color=NAVY):
    p = doc.add_paragraph(style="Callout")
    set_cell_shading(p, fill)
    set_paragraph_border(p)
    label_run = p.add_run(label + "  ")
    set_run_font(label_run, size=10.5, color=label_color, bold=True)
    body_run = p.add_run(text)
    set_run_font(body_run, size=10.5, color=NAVY)
    return p


def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    relationship_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), relationship_id)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), BLUE)
    rpr.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rpr.append(underline)
    rfonts = OxmlElement("w:rFonts")
    rfonts.set(qn("w:ascii"), "Calibri")
    rfonts.set(qn("w:hAnsi"), "Calibri")
    rfonts.set(qn("w:eastAsia"), "Calibri")
    rpr.append(rfonts)
    run.append(rpr)
    text_node = OxmlElement("w:t")
    text_node.text = text
    run.append(text_node)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def next_id(elements, attr):
    values = []
    for element in elements:
        value = element.get(qn(attr))
        if value and value.isdigit():
            values.append(int(value))
    return max(values, default=0) + 1


def add_numbering(doc, fmt, level_text, bullet_font=None):
    numbering = doc.part.numbering_part._element
    abstract_id = next_id(numbering.findall(qn("w:abstractNum")), "w:abstractNumId")
    num_id = next_id(numbering.findall(qn("w:num")), "w:numId")

    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "singleLevel")
    abstract.append(multi)

    level = OxmlElement("w:lvl")
    level.set(qn("w:ilvl"), "0")
    start = OxmlElement("w:start")
    start.set(qn("w:val"), "1")
    level.append(start)
    num_fmt = OxmlElement("w:numFmt")
    num_fmt.set(qn("w:val"), fmt)
    level.append(num_fmt)
    lvl_text = OxmlElement("w:lvlText")
    lvl_text.set(qn("w:val"), level_text)
    level.append(lvl_text)
    lvl_jc = OxmlElement("w:lvlJc")
    lvl_jc.set(qn("w:val"), "left")
    level.append(lvl_jc)
    ppr = OxmlElement("w:pPr")
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "num")
    tab.set(qn("w:pos"), "540")
    tabs.append(tab)
    ppr.append(tabs)
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "540")
    ind.set(qn("w:hanging"), "270")
    ppr.append(ind)
    level.append(ppr)
    if bullet_font:
        rpr = OxmlElement("w:rPr")
        rfonts = OxmlElement("w:rFonts")
        rfonts.set(qn("w:ascii"), bullet_font)
        rfonts.set(qn("w:hAnsi"), bullet_font)
        rpr.append(rfonts)
        level.append(rpr)
    abstract.append(level)
    numbering.append(abstract)

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abstract_ref = OxmlElement("w:abstractNumId")
    abstract_ref.set(qn("w:val"), str(abstract_id))
    num.append(abstract_ref)
    numbering.append(num)
    return num_id


def add_multilevel_numbering(doc):
    numbering = doc.part.numbering_part._element
    abstract_id = next_id(numbering.findall(qn("w:abstractNum")), "w:abstractNumId")
    num_id = next_id(numbering.findall(qn("w:num")), "w:numId")

    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "multilevel")
    abstract.append(multi)

    for ilvl, level_text, left, hanging in [
        (0, "%1.", 540, 270),
        (1, "%1.%2", 900, 360),
    ]:
        level = OxmlElement("w:lvl")
        level.set(qn("w:ilvl"), str(ilvl))
        start = OxmlElement("w:start")
        start.set(qn("w:val"), "1")
        level.append(start)
        num_fmt = OxmlElement("w:numFmt")
        num_fmt.set(qn("w:val"), "decimal")
        level.append(num_fmt)
        lvl_text = OxmlElement("w:lvlText")
        lvl_text.set(qn("w:val"), level_text)
        level.append(lvl_text)
        lvl_jc = OxmlElement("w:lvlJc")
        lvl_jc.set(qn("w:val"), "left")
        level.append(lvl_jc)
        if ilvl == 1:
            restart = OxmlElement("w:lvlRestart")
            restart.set(qn("w:val"), "1")
            level.append(restart)
        ppr = OxmlElement("w:pPr")
        tabs = OxmlElement("w:tabs")
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), "num")
        tab.set(qn("w:pos"), str(left))
        tabs.append(tab)
        ppr.append(tabs)
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), str(left))
        ind.set(qn("w:hanging"), str(hanging))
        ppr.append(ind)
        level.append(ppr)
        abstract.append(level)

    numbering.append(abstract)
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abstract_ref = OxmlElement("w:abstractNumId")
    abstract_ref.set(qn("w:val"), str(abstract_id))
    num.append(abstract_ref)
    numbering.append(num)
    return num_id


def clear_num(paragraph):
    ppr = paragraph._p.get_or_add_pPr()
    existing = ppr.find(qn("w:numPr"))
    if existing is not None:
        ppr.remove(existing)


def apply_num(paragraph, num_id, level=0):
    clear_num(paragraph)
    ppr = paragraph._p.get_or_add_pPr()
    num_pr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), str(level))
    num_id_el = OxmlElement("w:numId")
    num_id_el.set(qn("w:val"), str(num_id))
    num_pr.append(ilvl)
    num_pr.append(num_id_el)
    ppr.append(num_pr)


def add_list_item(doc, text, num_id, style="Guide List"):
    p = doc.add_paragraph(style=style)
    apply_num(p, num_id)
    run = p.add_run(text)
    set_run_font(run, size=10.5, color="000000")
    return p


def convert_section_numbering(doc, admin_num_id, volunteer_num_id):
    """Convert existing numbered paragraphs to 1 -> 1.1 multi-level lists."""
    active_num_id = None
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if paragraph.style.name == "Heading 1":
            if "給管理者" in text:
                active_num_id = admin_num_id
            elif "給志工" in text:
                active_num_id = volunteer_num_id
            else:
                active_num_id = None
            continue

        if active_num_id is None:
            continue

        if paragraph.style.name == "Heading 2":
            paragraph.text = re.sub(r"^\d+\.\s*", "", text)
            apply_num(paragraph, active_num_id, level=0)
        elif paragraph.style.name == "Guide List":
            apply_num(paragraph, active_num_id, level=1)


def add_label_paragraph(doc, label, text):
    p = doc.add_paragraph(style="Normal")
    label_run = p.add_run(label)
    set_run_font(label_run, size=10.5, color=NAVY, bold=True)
    text_run = p.add_run(text)
    set_run_font(text_run, size=10.5, color="000000")
    return p


def add_heading(doc, text, level=1):
    p = doc.add_paragraph(text, style=f"Heading {level}")
    p.paragraph_format.keep_with_next = True
    return p


def add_footer(section):
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run("第 ")
    set_run_font(run, size=8.5, color=MUTED)
    run = p.add_run()
    set_run_font(run, size=8.5, color=MUTED)
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr)
    run._r.append(fld_char2)
    run = p.add_run(" 頁")
    set_run_font(run, size=8.5, color=MUTED)


def build_document():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.header_distance = Inches(0.492)
    section.footer_distance = Inches(0.492)

    styles = doc.styles
    normal = styles["Normal"]
    set_style_font(normal, size=11, color="000000")
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.25

    for name, size, color, before, after in [
        ("Heading 1", 16, BLUE, 18, 10),
        ("Heading 2", 13, BLUE, 14, 7),
        ("Heading 3", 12, DARK_BLUE, 10, 5),
    ]:
        style = styles[name]
        set_style_font(style, size=size, color=color, bold=True)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.line_spacing = 1.15
        style.paragraph_format.keep_with_next = True

    for name in ("Guide List", "Callout"):
        if name not in styles:
            styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    list_style = styles["Guide List"]
    set_style_font(list_style, size=10.5, color="000000")
    list_style.paragraph_format.left_indent = Inches(0.375)
    list_style.paragraph_format.first_line_indent = Inches(-0.188)
    list_style.paragraph_format.space_after = Pt(4)
    list_style.paragraph_format.line_spacing = 1.25
    callout_style = styles["Callout"]
    set_style_font(callout_style, size=10.5, color=NAVY)
    callout_style.paragraph_format.left_indent = Inches(0.12)
    callout_style.paragraph_format.right_indent = Inches(0.12)
    callout_style.paragraph_format.space_before = Pt(4)
    callout_style.paragraph_format.space_after = Pt(8)
    callout_style.paragraph_format.line_spacing = 1.2

    add_footer(section)
    decimal_id = add_numbering(doc, "decimal", "%1.")
    admin_num_id = add_multilevel_numbering(doc)
    volunteer_num_id = add_multilevel_numbering(doc)
    bullet_id = add_numbering(doc, "bullet", "•", bullet_font="Symbol")

    # First-page editorial cover / quick start.
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(30)
    p.paragraph_format.space_after = Pt(10)
    run = p.add_run("NONPROFIT OPERATING GUIDE")
    set_run_font(run, size=9.5, color=BLUE, bold=True)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run("QR 報到系統｜使用說明")
    set_run_font(run, size=25, color=NAVY, bold=True)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(20)
    run = p.add_run("管理者與志工快速指南")
    set_run_font(run, size=12.5, color=MUTED)

    add_callout(
        doc,
        "一句話",
        "管理者建立活動與名單；志工使用活動連結掃描報到；活動後由管理者下載 CSV 或列印 / 存成 PDF。",
        fill=LIGHT_BLUE,
    )

    p = doc.add_paragraph(style="Normal")
    label = p.add_run("系統網址：")
    set_run_font(label, size=10.5, color=NAVY, bold=True)
    add_hyperlink(p, "開啟報到系統（qr-checkin-3iv.pages.dev）", "https://qr-checkin-3iv.pages.dev/")
    p.paragraph_format.space_after = Pt(12)

    add_heading(doc, "先記住 4 件事", 1)
    for text in [
        "使用 Google 帳號登入；只能使用已授權的帳號。",
        "每個活動有一個活動連結；同一活動可把同一連結發到志工群組。",
        "活動資料表會建立在「建立活動的管理者」Google Drive。",
        "目前報到時間固定以台灣時間（Asia/Taipei）記錄。",
    ]:
        add_list_item(doc, text, bullet_id)

    add_heading(doc, "角色分工", 1)
    add_label_paragraph(doc, "管理者：", "建立活動、匯入名單、加入志工、發送活動連結、產生 QR 與下載報告。")
    add_label_paragraph(doc, "志工：", "使用活動連結登入、掃描 QR、查詢總表，必要時進行手動報到。")

    doc.add_page_break()

    # Administrator section.
    add_heading(doc, "一、給管理者", 1)
    add_heading(doc, "1. 登入與進入行政", 2)
    for text in [
        "開啟系統網址，按「使用 Google 帳號登入」，使用已授權的管理者帳號。",
        "登入後預設進入「掃描」；請按下方「行政」進行活動設定。",
    ]:
        add_list_item(doc, text, decimal_id)

    add_heading(doc, "2. 建立活動", 2)
    for text in [
        "在「行政」輸入活動名稱，按「建立活動」。活動名稱請寫清楚，例如：1150909 金山活動。",
        "第一次建立活動時，Google 會要求 Drive 授權；請按允許。",
        "活動資料表會建立在這位管理者的 Google Drive。其他管理者可透過活動連結操作系統；若要直接開啟 Google Sheet，仍需另外在 Drive 共用。",
    ]:
        add_list_item(doc, text, decimal_id)

    add_heading(doc, "3. 加入志工帳號", 2)
    for text in [
        "在「行政 → 志工帳號」輸入志工登入 Google 使用的完整 email，按「加入」。",
        "請志工用同一個 email 登入；email 拼字不同會被視為不同帳號。",
    ]:
        add_list_item(doc, text, decimal_id)

    add_heading(doc, "4. 匯入活動名單", 2)
    for text in [
        "按「下載活動名單範本」，用 Excel 填寫後儲存。範本預留 1,000 位參與者，未使用的列留白即可。",
        "編號：必填、不可重複，請維持文字格式；產生 QR 後不要更改。",
        "姓名：必填。通訊地址、備註：依活動需要填寫，可留空。",
        "回到目前活動，選擇檔案後按「上傳活動名單」。",
    ]:
        add_list_item(doc, text, decimal_id)
    add_callout(
        doc,
        "重要",
        "重新上傳會取代目前活動名單。請確認 Excel 是完整名單，不要只上傳新增的幾列。",
        fill="FFF8E8",
        label_color="7A5A00",
    )

    add_heading(doc, "5. 發送活動連結", 2)
    for text in [
        "在目前活動區按「確認志工帳號名單」，確認志工 email 已加入。",
        "複製活動連結，發到志工群組；同一活動的所有志工使用同一個連結。每個活動的連結不同。",
        "連結頁面會顯示活動名稱，方便志工確認自己進入正確活動。",
    ]:
        add_list_item(doc, text, decimal_id)

    add_heading(doc, "6. 活動前、活動中與活動後", 2)
    for text in [
        "活動前：在行政頁按「產生 QR 碼卡片」或下載 Word 合併列印名單，製作並發放 QR。先請一位志工測試登入與掃描。",
        "活動中：管理者也可使用掃描、總表與手動報到；名單外的人可在掃描頁按「新增現場報到」。",
        "活動後：在「出席報告」按「下載 CSV」用 Excel 開啟，或按「列印報告」後在列印視窗選「另存為 PDF」。",
    ]:
        add_list_item(doc, text, decimal_id)

    doc.add_page_break()

    # Volunteer section.
    add_heading(doc, "二、給志工", 1)
    add_heading(doc, "1. 登入正確活動", 2)
    for text in [
        "向管理者取得這次活動的活動連結，開啟後按「使用 Google 帳號登入」。",
        "請使用管理者加入名單的同一個 Google email；不要用另一個帳號代替。",
        "登入後預設進入「掃描」。若看到的活動名稱不對，先回報管理者，不要繼續報到。",
    ]:
        add_list_item(doc, text, decimal_id)

    add_heading(doc, "2. 掃描 QR 報到", 2)
    for text in [
        "第一次使用時，允許瀏覽器使用手機相機。",
        "把參與者 QR Code 對準掃描框；畫面顯示「報到成功」即完成。",
        "成功畫面會顯示編號、姓名與報到時間；點一下結果區，或等待約 3 秒後可掃描下一位。",
        "若顯示已報到，不要重複掃描；可查看原本的報到時間。",
    ]:
        add_list_item(doc, text, decimal_id)

    add_heading(doc, "3. 查詢、手動報到與現場報到", 2)
    for text in [
        "按下方「總表」可搜尋姓名或編號；綠色列代表已報到，並會顯示報到時間。",
        "點選參與者，可進入手動報到頁；必要時可按「取消報到」。",
        "若參與者不在名單，從掃描頁按「新增現場報到」，填寫姓名及需要的資料後送出。",
        "其他手機完成報到後，總表通常會在約 5 秒內更新；也可以按右上角重新整理。",
    ]:
        add_list_item(doc, text, decimal_id)

    add_callout(
        doc,
        "注意",
        "不要直接修改 Google Sheet 的報到欄位；請使用系統操作，才能保留正確的報到人員與時間。",
        fill=LIGHT_BLUE,
    )

    add_heading(doc, "三、遇到問題時", 1)
    add_label_paragraph(doc, "顯示「此帳號未獲授權」：", "請確認登入的 Google email 是否與管理者加入的 email 完全相同；仍無法登入時，請管理者加入或修正帳號後重新登入。")
    add_label_paragraph(doc, "掃描器無法開啟：", "在瀏覽器設定允許相機，重新整理頁面；仍無法使用時，可改用總表的手動報到。")
    add_label_paragraph(doc, "查不到活動或參與者：", "確認使用的是本次活動連結，並檢查編號是否與名單完全一致。")
    add_label_paragraph(doc, "總表數字看起來沒更新：", "按右上角重新整理；若仍未更新，確認手機網路正常，再通知管理者。")

    add_callout(
        doc,
        "活動前 1 分鐘檢查",
        "管理者已建立正確活動 → 名單已完整上傳 → 志工 email 已加入 → 活動連結已發出 → 至少完成一次測試掃描。",
        fill="EAF6EE",
        label_color="1F5D3A",
    )

    doc.core_properties.title = "QR 報到系統｜使用說明"
    doc.core_properties.subject = "Nonprofit 管理者與志工快速指南"
    doc.core_properties.author = "QR 報到系統"
    convert_section_numbering(doc, admin_num_id, volunteer_num_id)
    doc.core_properties.comments = ""
    doc.save(OUTPUT)
    return OUTPUT


if __name__ == "__main__":
    path = build_document()
    print(path)
