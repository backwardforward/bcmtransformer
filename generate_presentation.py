import argparse
import os
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Cm, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import logging

logging.basicConfig(level=logging.INFO)

def parse_args():
    parser = argparse.ArgumentParser(description="Generate Business Capability Map PPTX")
    parser.add_argument('--fontSizeLevel1', type=int, required=True)
    parser.add_argument('--fontSizeLevel2', type=int, required=True)
    parser.add_argument('--colorFillLevel1', type=str, required=True)
    parser.add_argument('--colorFillLevel2', type=str, required=True)
    parser.add_argument('--textColorLevel1', type=str, required=True)
    parser.add_argument('--textColorLevel2', type=str, required=True)
    parser.add_argument('--borderColor', type=str, required=True)
    parser.add_argument('--widthLevel2', type=float, required=True)
    parser.add_argument('--heightLevel2', type=float, required=True)
    parser.add_argument('--excelPath', type=str, required=False, default=None)
    return parser.parse_args()

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def add_colored_box(slide, left, top, width, height, text, fill_color, border_color, border_width, font_size, bold, text_color, align_left_top=False):
    try:
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, left, top, width, height
        )
        rgb_fill = RGBColor(*hex_to_rgb(fill_color))
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb_fill
        rgb_border = RGBColor(*hex_to_rgb(border_color))
        shape.line.color.rgb = rgb_border
        shape.line.width = Pt(border_width)
        shape.text = text
        text_frame = shape.text_frame
        for p in text_frame.paragraphs:
            for run in p.runs:
                run.font.size = Pt(font_size)
                run.font.bold = bold
                run.font.color.rgb = RGBColor(*hex_to_rgb(text_color))
        if align_left_top:
            from pptx.enum.text import PP_ALIGN
            from pptx.enum.text import MSO_ANCHOR
            text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
            text_frame.vertical_anchor = MSO_ANCHOR.TOP
        else:
            text_frame.paragraphs[0].alignment = 1  # center
        return shape
    except Exception as e:
        logging.error(f"Error in add_colored_box: {e}")
        return None

def add_business_capabilities(slide, df, args, prs_width, prs_height):
    level1_col = 'ID_1'
    level2_col = 'ID_2'
    level3_col = 'ID_3'
    capability_col = 'LEVEL_1_CAPABILITY'
    subcapability_col = 'LEVEL_2_CAPABILITY'
    numbering_col = 'FULL_ID'

    margin = Inches(0.3)
    l1_text_margin = Inches(0.2)
    l2_box_margin = Inches(0.2)
    l2_box_min_height = Inches(0.8)
    l2_box_side_margin = Inches(0.2)

    unique_level1s = df[level1_col].unique()
    n_level1 = len(unique_level1s)
    if n_level1 == 0:
        return
    l1_box_width = (prs_width - 2 * margin - (n_level1 - 1) * margin) / n_level1
    l1_box_lefts = [margin + i * (l1_box_width + margin) for i in range(n_level1)]
    l1_box_top = margin
    l1_box_bottom_margin = margin

    for i, l1_value in enumerate(unique_level1s):
        l1_left = l1_box_lefts[i]
        l1_row = df[df[level1_col] == l1_value].iloc[0]
        l1_num = str(int(l1_row[level1_col]))
        l1_text = f"{l1_num}. {l1_row[capability_col]}"
        # Level 2 boxes for this Level 1
        level2s = df[(df[level1_col] == l1_value) & (df[level2_col] != 1)]
        n_level2 = len(level2s)
        # Calculate available height for Level 2 boxes
        available_height = prs_height - l1_box_top - l1_box_bottom_margin
        # Reserve space for Level 1 text at the top
        l1_text_height = Inches(0.7)
        l2_area_top = l1_box_top + l1_text_height + l2_box_margin
        l2_area_height = available_height - l1_text_height - l2_box_margin
        if n_level2 > 0:
            l2_box_height = max(l2_box_min_height, (l2_area_height - (n_level2 - 1) * l2_box_margin) / n_level2)
        else:
            l2_box_height = 0
        # Draw Level 1 background box (full height for this column)
        add_colored_box(
            slide, l1_left, l1_box_top, l1_box_width, available_height,
            '', args.colorFillLevel1, args.borderColor, 2, args.fontSizeLevel1, True, args.textColorLevel1, align_left_top=False
        )
        # Draw Level 1 text at the top (inside the Level 1 box)
        add_colored_box(
            slide, l1_left + l1_text_margin, l1_box_top + l1_text_margin, l1_box_width - 2 * l1_text_margin, l1_text_height - l1_text_margin,
            l1_text, args.colorFillLevel1, args.borderColor, 0, args.fontSizeLevel1, True, args.textColorLevel1, align_left_top=True
        )
        # Draw Level 2 boxes on top
        for j, (_, l2_row) in enumerate(level2s.iterrows()):
            l2_left = l1_left + l2_box_side_margin
            l2_top = l2_area_top + j * (l2_box_height + l2_box_margin)
            l2_num = l2_row[numbering_col] if numbering_col in l2_row else f"{l1_num}.{j+1}"
            l2_text = f"{l2_num} {l2_row[subcapability_col]}"
            add_colored_box(
                slide, l2_left, l2_top, l1_box_width - 2 * l2_box_side_margin, l2_box_height,
                l2_text, args.colorFillLevel2, args.borderColor, 1.5, args.fontSizeLevel2, False, args.textColorLevel2, align_left_top=True
            )

def generate_from_dataframe(df, args):
    # Spaltennamen im DataFrame normalisieren!
    df.columns = [c.strip().upper().replace(' ', '_') for c in df.columns]
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_business_capabilities(slide, df, args, prs.slide_width, prs.slide_height)
    output_dir = os.path.join(os.path.dirname(__file__), 'output')
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, 'Business_Capability_Map.pptx')
    prs.save(output_path)
    logging.info(f"Saved presentation to {output_path}")
    print(f"Saved presentation to {output_path}")
    return output_path

def main():
    args = parse_args()
    excel_path = args.excelPath or os.path.join(os.path.dirname(__file__), 'excel_data', 'bcm_test_source.xlsx')
    if not os.path.exists(excel_path):
        logging.error(f"Excel file not found: {excel_path}")
        return
    df = pd.read_excel(excel_path)
    generate_from_dataframe(df, args)

if __name__ == "__main__":
    main()
