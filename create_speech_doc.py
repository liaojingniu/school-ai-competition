# -*- coding: utf-8 -*-
"""
生成《小学AI应用场景比赛 - 一年级小朋友极简操作演说稿》Word文档
【超精简·零压力定制版】：全篇台词仅 5~6 句话，字大如画，像背儿歌一样简单，1分钟轻松演练！
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=140, bottom=140, left=160, right=160):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
    for edge in ('top', 'left', 'bottom', 'right'):
        edge_data = kwargs.get(edge)
        if edge_data:
            val, sz, space, color = edge_data
            tag = f'<w:{edge} {nsdecls("w")} w:val="{val}" w:sz="{sz}" w:space="{space}" w:color="{color}"/>'
            tcBorders.append(parse_xml(tag))
    tcPr.append(tcBorders)

def create_document():
    doc = docx.Document()
    
    # 页面边距设为较宽适中
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

    doc.styles['Normal'].font.name = '微软雅黑'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    doc.styles['Normal'].font.size = Pt(12)
    doc.styles['Normal'].font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    # ---------------- 标题区 ----------------
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(4)
    run_title = p_title.add_run("🏆 小学AI比赛 · 极简演说演示稿")
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1B, 0x4F, 0x72)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    run_sub = p_sub.add_run("（一年级小朋友专属 · 全篇仅5句话 · 1分钟零压力背诵）")
    run_sub.font.size = Pt(12)
    run_sub.font.bold = True
    run_sub.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)

    # 参赛选手基本信息
    t_info = doc.add_table(rows=1, cols=3)
    t_info.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_cells = t_info.rows[0].cells
    info_data = [
        ("参赛选手", "___________"),
        ("年级班级", "一年级 _____班"),
        ("指导老师/家长", "___________")
    ]
    for i, (label, val) in enumerate(info_data):
        c = info_cells[i]
        set_cell_background(c, "F4F6F6")
        set_cell_margins(c, 80, 80, 100, 100)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p.add_run(f"{label}：")
        r1.font.bold = True
        r1.font.size = Pt(10)
        r2 = p.add_run(val)
        r2.font.size = Pt(10)
        r2.font.color.rgb = RGBColor(0x1B, 0x4F, 0x72)

    # ---------------- 核心口诀 ----------------
    p_cue = doc.add_paragraph()
    p_cue.paragraph_format.space_before = Pt(12)
    p_cue.paragraph_format.space_after = Pt(12)
    p_cue.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_cue_box = p_cue.add_run("⭐ 小朋友速记口诀：按 1 举 2 开风扇，按 2 举 4 开小灯！ ⭐")
    r_cue_box.font.size = Pt(13)
    r_cue_box.font.bold = True
    r_cue_box.font.color.rgb = RGBColor(0xD3, 0x54, 0x00)

    # ---------------- 第一页：小朋友直接朗读背诵卡（超大字） ----------------
    p_sec1 = doc.add_paragraph()
    p_sec1.paragraph_format.space_before = Pt(6)
    p_sec1.paragraph_format.space_after = Pt(6)
    r_sec1 = p_sec1.add_run("一、 🌟 小朋友大字背诵卡（只需背这 5 句话！）")
    r_sec1.font.size = Pt(15)
    r_sec1.font.bold = True
    r_sec1.font.color.rgb = RGBColor(0x1B, 0x4F, 0x72)

    simple_steps = [
        (
            "① 开场问好",
            "【动作：鞠躬，微笑大声说】",
            "“评委老师好！我是一年级的【你的名字】。\n今天我带来了聪明的——《AI视觉小管家》！”"
        ),
        (
            "② 介绍原理",
            "【动作：指一指摄像头和小板子】",
            "“它有一只神奇的眼睛，只要看懂数字，就能帮我开关风扇和小灯！”"
        ),
        (
            "③ 开风扇",
            "【动作：按按键 1，双手举起数字 2 卡片】",
            "“天气好热呀，看我出示数字 2——\n哇！风扇转起来啦，好凉快呀！”"
        ),
        (
            "④ 开小灯",
            "【动作：按按键 2，双手举起数字 4 卡片】",
            "“天黑要开灯，看我出示数字 4——\n哇！小灯亮起来啦，好明亮呀！”"
        ),
        (
            "⑤ 收尾致谢",
            "【动作：放下卡片，立正鞠躬】",
            "“它是不是很神奇？我的演示完毕，谢谢评委老师！”"
        )
    ]

    for title, act, speech in simple_steps:
        p_card = doc.add_paragraph()
        p_card.paragraph_format.space_before = Pt(8)
        p_card.paragraph_format.space_after = Pt(4)
        p_card.paragraph_format.line_spacing = 1.35
        
        # 步骤标题
        r_t = p_card.add_run(f"{title} ")
        r_t.font.size = Pt(13.5)
        r_t.font.bold = True
        r_t.font.color.rgb = RGBColor(0x1F, 0x61, 0x8D)

        # 动作提示
        r_a = p_card.add_run(f"{act}\n")
        r_a.font.size = Pt(11)
        r_a.font.bold = True
        r_a.font.color.rgb = RGBColor(0x27, 0xAE, 0x60)

        # 真正说话台词（特大字加粗）
        r_s = p_card.add_run(speech)
        r_s.font.size = Pt(15)
        r_s.font.bold = True
        r_s.font.color.rgb = RGBColor(0x17, 0x20, 0x2A)

    doc.add_page_break()

    # ---------------- 第二页：现场动作对照表（家长/老师陪练用） ----------------
    p_sec2 = doc.add_paragraph()
    p_sec2.paragraph_format.space_before = Pt(6)
    p_sec2.paragraph_format.space_after = Pt(8)
    r_sec2 = p_sec2.add_run("二、 🎬 现场边讲边演示·动作流程对照表")
    r_sec2.font.size = Pt(14)
    r_sec2.font.bold = True
    r_sec2.font.color.rgb = RGBColor(0x1B, 0x4F, 0x72)

    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths = [Inches(1.0), Inches(1.8), Inches(2.3), Inches(1.7)]
    headers = ["环节", "🎬 小朋友动作", "🎤 小朋友说的话", "💻 电脑与实物反应"]

    for i, h in enumerate(headers):
        c = table.rows[0].cells[i]
        c.width = col_widths[i]
        set_cell_background(c, "2E86C1")
        set_cell_margins(c, 100, 100, 100, 100)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    table_data = [
        (
            "第 1 步\n打招呼",
            "立正站好，微笑鞠躬。",
            "“评委老师好！我是一年级的【某某】。今天我带来了聪明的《AI视觉小管家》！”",
            "电脑待命，摄像头画面开启。"
        ),
        (
            "第 2 步\n讲本领",
            "指一下摄像头，再指一下ESP32小板子。",
            "“它有一只神奇的眼睛，只要看懂数字，就能帮我开关风扇和小灯！”",
            "画面清晰识别准备完毕。"
        ),
        (
            "第 3 步\n吹风扇",
            "① 按键盘【1】键。\n② 双手拿【数字2】平举正对摄像头。\n③ 看到风扇转，指给评委看。",
            "“天气好热呀，看我出示数字 2——\n哇！风扇转起来啦，好凉快呀！”",
            "① 语音：『请拿出数字进行识别』\n② 语音：『识别到数字2成功，风扇启动』\n③ 风扇转动 7 秒后自动停止。"
        ),
        (
            "第 4 步\n开小灯",
            "① 按键盘【2】键。\n② 双手拿【数字4】平举正对摄像头。\n③ 看到灯亮，做开心表情。",
            "“天黑要开灯，看我出示数字 4——\n哇！小灯亮起来啦，好明亮呀！”",
            "① 语音：『请拿出数字进行识别』\n② 语音：『识别到数字4成功，指示灯启动』\n③ LED小灯亮起 7 秒后自动熄灭。"
        ),
        (
            "第 5 步\n致谢收尾",
            "放下卡片，立正站好，鞠躬致谢。",
            "“它是不是很神奇？我的演示完毕，谢谢评委老师！”",
            "演说结束，等待评委老师表扬提问。"
        )
    ]

    for row_idx, row_data in enumerate(table_data):
        row = table.add_row()
        bg_color = "FFFFFF" if row_idx % 2 == 0 else "F8F9FA"
        for col_idx, cell_text in enumerate(row_data):
            cell = row.cells[col_idx]
            cell.width = col_widths[col_idx]
            set_cell_background(cell, bg_color)
            set_cell_margins(cell, 100, 100, 90, 90)
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            if col_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(cell_text)
                r.font.size = Pt(9.5)
                r.font.bold = True
                r.font.color.rgb = RGBColor(0x1B, 0x4F, 0x72)
            elif col_idx == 1:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run(cell_text)
                r.font.size = Pt(9.5)
                r.font.color.rgb = RGBColor(0x27, 0xAE, 0x60)
            elif col_idx == 2:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run(cell_text)
                r.font.size = Pt(10)
                r.font.bold = True
                r.font.color.rgb = RGBColor(0x8E, 0x44, 0xAD)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run(cell_text)
                r.font.size = Pt(9)
                r.font.color.rgb = RGBColor(0x34, 0x49, 0x5E)

    for r in table.rows:
        for c in r.cells:
            set_cell_border(
                c,
                top=('single', '4', '0', 'D5D8DC'),
                bottom=('single', '4', '0', 'D5D8DC'),
                left=('single', '4', '0', 'D5D8DC'),
                right=('single', '4', '0', 'D5D8DC')
            )

    # ---------------- 第三部分：家长现场防紧张与救场小锦囊 ----------------
    p_sec3 = doc.add_paragraph()
    p_sec3.paragraph_format.space_before = Pt(14)
    p_sec3.paragraph_format.space_after = Pt(4)
    r_sec3 = p_sec3.add_run("三、 💡 家长与老师现场“零失误”小锦囊")
    r_sec3.font.size = Pt(13)
    r_sec3.font.bold = True
    r_sec3.font.color.rgb = RGBColor(0x1B, 0x4F, 0x72)

    tips = [
        "1. 拿卡技巧：双手拿在白纸左右两侧，手指不挡住中间数字，平举在胸前，离摄像头 10~150 厘米（近到贴脸10cm、远到1.5米）都能秒级识别！",
        "2. 键盘极简：让孩子只按【1】（风扇）和【2】（电灯），出示非目标数字时系统仅提醒一次重试，绝不反复催促打扰演说节奏。",
        "3. 救场神器：万一现场光线差或反光，指导老师坐在旁边敲键盘【Y】键，系统会立刻模拟成功触发动作，绝不卡壳！"
    ]
    for tip in tips:
        p_t = doc.add_paragraph()
        p_t.paragraph_format.space_before = Pt(2)
        p_t.paragraph_format.space_after = Pt(2)
        r = p_t.add_run(tip)
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "AI比赛操作演说稿_一年级小朋友专用.docx")
    doc.save(output_path)
    print(f"[OK] 极简版演说稿更新成功: {output_path}")

if __name__ == '__main__':
    create_document()
