# -*- coding: utf-8 -*-
"""
generate_ppt.py
生成：output/瑞幸中秋营销提案.pptx
依赖：python-pptx, Pillow

运行：python generate_ppt.py

说明：此脚本会：
- 生成17页 PPTX（中文），每页含标题、要点与讲稿备注
- 用 Pillow 生成三张原创数位手绘风格 PNG 视觉稿并嵌入幻灯片
- 最后一页包含可核查的数据来源占位文本

注意：脚本以最常见的 PPT 布局创建幻灯片，若需美术级调整请在 PowerPoint 中打开并微调。
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from PIL import Image, ImageDraw, ImageFont
import os

OUTPUT_DIR = 'output'
OUTPUT_FILE = os.path.join(OUTPUT_DIR, '瑞幸中秋营销提案.pptx')
IMAGE_DIR = 'assets'

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(IMAGE_DIR, exist_ok=True)

# 简单字体处理：使用系统常见字体路径或Pillow自带默认
try:
    FONT_PATH = '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
    ImageFont.truetype(FONT_PATH, 24)
except Exception:
    FONT_PATH = None

def create_sketch_1(path):
    # 月见杯：夜色城市剪影 + 咖啡杯 + 月亮
    W, H = 1200, 675
    bg = (10, 18, 40)
    img = Image.new('RGB', (W, H), bg)
    draw = ImageDraw.Draw(img)
    # moon
    draw.ellipse((820, 60, 980, 220), fill=(250, 230, 170))
    # city silhouette
    for i in range(0, 12):
        x = i * 100
        h = 120 + (i % 5) * 30
        draw.rectangle((x, H - h, x + 80, H), fill=(22, 34, 56))
    # coffee cup (simple)
    draw.rectangle((420, 320, 780, 500), fill=(245,245,245))
    draw.ellipse((420, 300, 780, 380), fill=(245,245,245))
    draw.ellipse((460, 310, 740, 390), fill=(255, 255, 220))
    # glow
    draw.ellipse((740, 120, 1040, 420), fill=(255,240,200,80))
    # text
    f = ImageFont.truetype(FONT_PATH, 36) if FONT_PATH else None
    draw.text((60, 40), '月见杯 — 月见山海', fill=(230,230,230), font=f)
    img.save(path, quality=90)


def create_sketch_2(path):
    # 门店打卡位草图
    W, H = 1200, 675
    bg = (8, 16, 34)
    img = Image.new('RGB', (W, H), bg)
    draw = ImageDraw.Draw(img)
    # moon backdrop
    draw.ellipse((900, 40, 1080, 220), fill=(255, 244, 200))
    # store facade
    draw.rectangle((240, 240, 960, 540), fill=(30, 40, 60))
    draw.rectangle((300, 280, 900, 420), fill=(14, 22, 36))
    # sign
    f = ImageFont.truetype(FONT_PATH, 28) if FONT_PATH else None
    draw.text((480, 250), '瑞幸', fill=(255,215,120), font=f)
    # person silhouette
    draw.ellipse((360, 420, 420, 480), fill=(60,70,90))
    draw.rectangle((384, 460, 396, 560), fill=(60,70,90))
    draw.text((60, 40), '门店打卡位草图', fill=(220,220,220), font=f)
    img.save(path, quality=90)


def create_sketch_3(path):
    # 城市地图 + 记忆轨迹
    W, H = 1200, 675
    bg = (7, 14, 28)
    img = Image.new('RGB', (W, H), bg)
    draw = ImageDraw.Draw(img)
    f = ImageFont.truetype(FONT_PATH, 28) if FONT_PATH else None
    # map-like lines
    points = [(120, 520), (240, 380), (360, 420), (520, 300), (700, 360), (880, 260), (980, 420), (1100, 360)]
    for i in range(len(points)-1):
        draw.line((points[i], points[i+1]), fill=(80,110,160), width=6)
    # moon at top center
    draw.ellipse((500, 40, 700, 240), fill=(252, 236, 180))
    draw.text((60, 40), '城市地图 · 记忆轨迹', fill=(220,220,220), font=f)
    img.save(path, quality=90)

# 生成草图
sk1 = os.path.join(IMAGE_DIR, 'sketch_moon_cup.png')
sk2 = os.path.join(IMAGE_DIR, 'sketch_store.png')
sk3 = os.path.join(IMAGE_DIR, 'sketch_map.png')

create_sketch_1(sk1)
create_sketch_2(sk2)
create_sketch_3(sk3)

# PPT 内容（17页）
slides_content = [
    { 'title': '月见山海｜瑞幸中秋节营销提案',
      'bullets': ['把中秋从“送礼”变成“城市记忆”'],
      'note': '本次提案的核心：体验化、故事化、社交化。' , 'image': sk1},
    { 'title': '提案概览：把中秋做成瑞幸的“城市生活仪式”',
      'bullets': ['目标：情绪连接 + 线下体验 + 社交传播','价值：城市故事，而非单一商品'],
      'note': '三个价值：提升门店、强化品牌、形成传播资产' , 'image': None},
    { 'title': '洞察一：中秋不再只是“送礼”',
      'bullets': ['消费趋向体验与纪念','年轻人偏好可打卡的消费体验'],
      'note': '传统礼盒的可替代性高，体验更能产生记忆' , 'image': None},
    { 'title': '洞察二：咖啡是年轻人中秋社交语言',
      'bullets': ['中秋包含家庭团聚与城市社交','咖啡具备情绪连接力'],
      'note': '瑞幸的优势在于日常消费的高频触点' , 'image': None},
    { 'title': '受众洞察：他们想要有故事的消费',
      'bullets': ['目标：18-35岁城市年轻人','高频咖啡用户与礼赠中间人'],
      'note': '用户心理：想要能被记住的惊喜' , 'image': None},
    { 'title': '品牌抓手：做“月见”，不是“月饼”',
      'bullets': ['主题：月见城市计划','口号：一杯咖啡，一场月见'],
      'note': '品牌要成为城市中的情绪陪伴者' , 'image': None},
    { 'title': '创意主案：月见城市计划',
      'bullets': ['门店为月见站点','用户共同创作城市故事','UGC引导传播'],
      'note': '三层结构：故事 / 门店 / 传播' , 'image': sk2},
    { 'title': '产品方案：可见的记忆载体',
      'bullets': ['限定饮品（如月见拿铁）','主题礼赠组合与会员权益'],
      'note': '商品设计强调记忆载体而非堆叠礼盒' , 'image': None},
    { 'title': '门店方案：月见站点设计',
      'bullets': ['月光专区、打卡区、城市故事墙','统一门店话术与视觉指引'],
      'note': '线下体验是情绪连接的核心' , 'image': sk2},
    { 'title': '传播方案：用户共创的城市故事',
      'bullets': ['城市短片（12城）','UGC挑战 #月见城市#','社媒联动：抖音/小红书/视频号'],
      'note': '传播以叙事为主，弱化广告感' , 'image': None},
    { 'title': '概念视频脚本（示例）',
      'bullets': ['夜色城市、门店、扫码留言、月光场景'],
      'note': '短片以“被记住的城市瞬间”为核心' , 'image': None},
    { 'title': '原创手绘草图与视觉稿',
      'bullets': ['月见杯 / 门店打卡位 / 城市地图轨迹'],
      'note': '三张手绘草图已生成并嵌入' , 'image': sk3},
    { 'title': '执行节奏：预热→进行→复购',
      'bullets': ['预热：海报+短片','进行：门店打卡+UGC','复购：会员回流与精选'],
      'note': '重点是传播后持续留存' , 'image': None},
    { 'title': 'KPI 与评估',
      'bullets': ['门店销售增长、会员新增与复购','社媒传播量、UGC参与率'],
      'note': '结合门店数据与社媒数据评估效果' , 'image': None},
    { 'title': '风险与应对',
      'bullets': ['被认为为普通促销→强化叙事','门店执行差异→统一SOP'],
      'note': '风险管理侧重叙事一致性与门店质量' , 'image': None},
    { 'title': '数据来源（公开可核查）',
      'bullets': ['国家统计局、艾媒咨询、阿里/天猫数据、QuestMobile、瑞幸官方披露'],
      'note': '每一页如引用具体数据，请在页脚标注来源' , 'image': None},
    { 'title': '结语：把中秋从“送礼”变成“被记住”',
      'bullets': ['让瑞幸成为城市记忆的一部分','一杯咖啡，一场月见'],
      'note': '结束语：呼应封面与提案主旨' , 'image': sk1},
]

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# helper to add title + bullets + image + notes
from pptx.util import Cm

def add_slide(prs, title, bullets, note=None, image_path=None):
    slide_layout = prs.slide_layouts[5]  # blank
    slide = prs.slides.add_slide(slide_layout)
    left = Inches(0.6)
    top = Inches(0.4)
    width = Inches(11)
    # Title
    tx = slide.shapes.add_textbox(left, top, width, Inches(1))
    tf = tx.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.name = 'Microsoft YaHei'
    p.font.color.rgb = RGBColor(255, 245, 230)
    # Bullets
    bx = slide.shapes.add_textbox(left, Inches(1.6), width * 0.6, Inches(4))
    btf = bx.text_frame
    btf.margin_left = Pt(6)
    btf.word_wrap = True
    for i, line in enumerate(bullets):
        p = btf.add_paragraph() if i>0 else btf.paragraphs[0]
        p.text = '• ' + line
        p.level = 0
        p.font.size = Pt(18)
        p.font.name = 'Microsoft YaHei'
        p.font.color.rgb = RGBColor(235, 235, 235)
    # Image
    if image_path and os.path.exists(image_path):
        img_left = Inches(8)
        img_top = Inches(1.6)
        slide.shapes.add_picture(image_path, img_left, img_top, width=Inches(4.6))
    # background dark rectangle
    from pptx.enum.shapes import MSO_SHAPE
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(10, 18, 40)
    bg.shadow = False
    bg.line.fill.background()
    # move bg to back
    slide.shapes._spTree.remove(bg._element)
    slide.shapes._spTree.insert(2, bg._element)
    # notes
    if note:
        slide.notes_slide.notes_text_frame.text = note

for s in slides_content:
    add_slide(prs, s['title'], s['bullets'], note=s.get('note'), image_path=s.get('image'))

# 保存文件
prs.save(OUTPUT_FILE)
print('PPT 已生成：', OUTPUT_FILE)
