from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


def load_font(size: int):
    for name in ("msyh.ttc", "simhei.ttf", "arial.ttf"):
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def box(draw, xy, text, fill, outline="#334155", text_fill="#111827", font=None):
    x1, y1, x2, y2 = xy
    draw.rounded_rectangle(xy, radius=10, fill=fill, outline=outline, width=2)
    lines = text.split("\n")
    line_height = font.size + 5 if hasattr(font, "size") else 18
    total = len(lines) * line_height
    y = y1 + ((y2 - y1) - total) / 2
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=font)
        draw.text((x1 + ((x2 - x1) - (bbox[2] - bbox[0])) / 2, y), line, fill=text_fill, font=font)
        y += line_height


def arrow(draw, start, end):
    draw.line([start, end], fill="#475569", width=3)
    x, y = end
    draw.polygon([(x, y), (x - 8, y - 5), (x - 8, y + 5)], fill="#475569")


def main() -> int:
    out = Path.home() / "Desktop" / "ecommerce-workflow-skill-structure.png"
    img = Image.new("RGB", (1500, 980), "#F8FAFC")
    draw = ImageDraw.Draw(img)
    title_font = load_font(34)
    font = load_font(22)
    small = load_font(18)

    draw.text((60, 40), "模块化电商数据分析 Workflow Skill", fill="#0F172A", font=title_font)
    draw.text((60, 86), "两个功能模块使用同一个源数据；总控脚本只负责调度顺序。", fill="#475569", font=font)

    box(draw, (560, 145, 940, 225), "ecommerce-workflow-skill\n新的组合 skill", "#DBEAFE", font=font)
    box(draw, (90, 325, 390, 405), "SKILL.md\nCodex 使用说明", "#E0F2FE", font=font)
    box(draw, (450, 325, 750, 405), "scripts\nrun_workflow.py 总入口", "#DCFCE7", font=font)
    box(draw, (810, 325, 1110, 405), "modules\n业务模块目录", "#FEF3C7", font=font)
    box(draw, (1170, 325, 1410, 405), "shared\n公共工具箱", "#FCE7F3", font=font)

    for x in (240, 600, 960, 1290):
        arrow(draw, (750, 225), (x, 325))

    box(draw, (120, 555, 430, 650), "同一个源数据 Excel", "#FFFFFF", font=font)
    box(draw, (550, 520, 840, 620), "product_research\n产品调研模块", "#FFEDD5", font=font)
    box(draw, (550, 690, 840, 790), "market_analysis\n市场分析模块", "#F3E8FF", font=font)
    box(draw, (970, 520, 1280, 620), "产品调研输出表\n+ 9 张图表 PNG", "#FFFFFF", font=small)
    box(draw, (970, 690, 1280, 790), "市场分析输出表\n+ 异常样本 sheet", "#FFFFFF", font=small)

    arrow(draw, (430, 600), (550, 570))
    arrow(draw, (430, 600), (550, 740))
    arrow(draw, (840, 570), (970, 570))
    arrow(draw, (840, 740), (970, 740))
    arrow(draw, (960, 405), (700, 520))
    arrow(draw, (960, 405), (700, 690))
    arrow(draw, (1290, 405), (700, 520))
    arrow(draw, (1290, 405), (700, 690))

    draw.text((1170, 430), "shared 里放：", fill="#0F172A", font=font)
    draw.text((1170, 468), "路径处理\nExcel/sheet 检查\nDISPIMG 风险检查\n子进程执行", fill="#475569", font=small, spacing=8)

    draw.text((60, 890), "以后新增功能：modules\\profit_analysis\\run.py + config.py，然后在 run_workflow.py 注册步骤。", fill="#334155", font=font)
    img.save(out)
    print(out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
