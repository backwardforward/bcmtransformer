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
    parser.add_argument('--outputPath', type=str, required=False, default=None)
    return parser.parse_args()

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def add_colored_box(slide, left, top, width, height, text, fill_color, border_color, border_width, font_size, bold, text_color, align_left_top=False):
    try:
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
        )
        rgb_fill = RGBColor(*hex_to_rgb(fill_color))
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb_fill
        rgb_border = RGBColor(*hex_to_rgb(border_color))
        shape.line.color.rgb = rgb_border
        shape.line.width = Pt(border_width)
        # Make border invisible if border_color is None or border_width is 0
        if not border_color or border_width == 0:
            shape.line.fill.background()
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
        # Make corners a little more rounded
        try:
            shape.adjustments[0] = 0.03
        except Exception:
            pass
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
    subsubcapability_col = 'LEVEL_3_CAPABILITY'
    numbering_col = 'FULL_ID'

    # Layout constants
    margin = Inches(0.04)  # minimal gap between L1 columns (about 3pt)
    l1_text_margin = Inches(0.15)
    l2_box_margin = Inches(0.1)
    l2_box_side_margin = Inches(0.15)
    l3_box_margin = Inches(0.1)
    l1_box_min_width = Inches(2.0)
    l2_box_min_width = Inches(1.4)
    l3_box_min_width = Inches(1.0)
    l1_box_min_height = Inches(0.8)
    l2_box_min_height = Inches(0.5)
    l3_box_min_height = Inches(0.35)
    l1_box_text_height = Inches(0.7)
    l2_box_text_height = Inches(0.5)
    l3_box_text_height = Inches(0.4)
    box_padding = Inches(0.15)
    child_left_pad = Inches(0.15)
    child_top_pad = Inches(0.15)
    vertical_spacing = Inches(0.1)
    min_box_height = Inches(0.2)
    min_spacing = Inches(0.05)
    min_padding = Inches(0.07)

    # Ensure all margin and padding values are defined as numbers (never None)
    l1_box_text_height = float(l1_box_text_height) if l1_box_text_height is not None else 0.0
    l2_box_text_height = float(l2_box_text_height) if l2_box_text_height is not None else 0.0
    l3_box_text_height = float(l3_box_text_height) if l3_box_text_height is not None else 0.0
    l2_box_side_margin = float(l2_box_side_margin) if l2_box_side_margin is not None else 0.0
    l2_box_margin = float(l2_box_margin) if l2_box_margin is not None else 0.0
    l3_box_margin = float(l3_box_margin) if l3_box_margin is not None else 0.0
    box_padding = float(box_padding) if box_padding is not None else 0.0
    child_left_pad = float(child_left_pad) if child_left_pad is not None else 0.0
    child_top_pad = float(child_top_pad) if child_top_pad is not None else 0.0
    vertical_spacing = float(vertical_spacing) if vertical_spacing is not None else 0.0
    min_box_height = float(min_box_height)
    min_spacing = float(min_spacing)
    min_padding = float(min_padding)

    # 1. Build tree: {L1_ID: {row, children: {L2_ID: {row, children: {L3_ID: {row}}}}}}
    tree = {}
    for _, row in df.iterrows():
        l1 = row[level1_col]
        l2 = row[level2_col] if level2_col in row and pd.notnull(row[level2_col]) else None
        l3 = row[level3_col] if level3_col in row and pd.notnull(row[level3_col]) else None
        # Force all IDs to str for tree keys
        l1_key = str(l1)
        l2_key = str(l2) if l2 is not None else None
        l3_key = str(l3) if l3 is not None else None
        if l1_key not in tree:
            tree[l1_key] = {'row': row, 'children': {}}
        if l2_key and l2_key != "1":
            if l2_key not in tree[l1_key]['children']:
                tree[l1_key]['children'][l2_key] = {'row': row, 'children': {}}
            if l3_key and l3_key != "1":
                if l3_key not in tree[l1_key]['children'][l2_key]['children']:
                    tree[l1_key]['children'][l2_key]['children'][l3_key] = {'row': row}

    # 2. Recursively compute heights for each node
    def compute_natural_height(node, level):
        if level == 3:
            return (l3_box_text_height or 0) + 2*(box_padding or 0)
        elif level == 2:
            children = node.get('children', {})
            if not children:
                return (l2_box_text_height or 0) + 2*(box_padding or 0)
            h = 0
            for c in children.values():
                h += (compute_natural_height(c, 3) or 0) + (vertical_spacing or 0)
            h -= (vertical_spacing or 0) if children else 0
            return h + 2*(box_padding or 0) + (l2_box_text_height or 0)
        elif level == 1:
            children = node.get('children', {})
            if not children:
                return (l1_box_text_height or 0) + 2*(box_padding or 0)
            h = 0
            for c in children.values():
                h += (compute_natural_height(c, 2) or 0) + (vertical_spacing or 0)
            h -= (vertical_spacing or 0) if children else 0
            return h + 2*(box_padding or 0) + (l1_box_text_height or 0)

    # 3. Recursively draw boxes
    def draw_node_scaled(node, level, left, top, width, scaling):
        # All verticals are scaled, but not below minimums
        if level == 1:
            l1_row = node['row']
            l1_num = str(l1_row[level1_col])
            l1_text = f"{l1_num}. {l1_row[capability_col]}"
            # Compute scaled heights
            height = max(l1_box_min_height or 0, scaling * (compute_natural_height(node, 1) or 0))
            pad = max(min_padding or 0, scaling * (box_padding or 0))
            header_h = l1_box_text_height or 0
            add_colored_box(
                slide, left, top, width, height,
                l1_text, args.colorFillLevel1, args.borderColor, 2, args.fontSizeLevel1, True, args.textColorLevel1, align_left_top=True
            )
            y = top + header_h + max(min_padding or 0, scaling * (child_top_pad or 0))
            for l2_id, l2_node in node.get('children', {}).items():
                l2_natural = compute_natural_height(l2_node, 2) or 0
                l2_height = max(l2_box_min_height or 0, scaling * l2_natural)
                draw_node_scaled(l2_node, 2, left + max(min_padding or 0, scaling * (child_left_pad or 0)), y, width - 2*max(min_padding or 0, scaling * (child_left_pad or 0)), scaling)
                y += l2_height + max(min_spacing or 0, scaling * (vertical_spacing or 0))
        elif level == 2:
            l2_row = node['row']
            l1_num = str(l2_row[level1_col])
            l2_num = str(l2_row[level2_col])
            l2_text = f"{l1_num}.{l2_num} {l2_row[subcapability_col]}"
            height = max(l2_box_min_height or 0, scaling * (compute_natural_height(node, 2) or 0))
            pad = max(min_padding or 0, scaling * (box_padding or 0))
            header_h = l2_box_text_height or 0
            add_colored_box(
                slide, left, top, width, height,
                l2_text, args.colorFillLevel2, args.borderColor, 1.5, args.fontSizeLevel2 + 1, True, args.textColorLevel2, align_left_top=True
            )
            y = top + header_h + max(min_padding or 0, scaling * (child_top_pad or 0))
            for l3_id, l3_node in node.get('children', {}).items():
                l3_natural = compute_natural_height(l3_node, 3) or 0
                l3_height = max(l3_box_min_height or 0, scaling * l3_natural)
                draw_node_scaled(l3_node, 3, left + max(min_padding or 0, scaling * (child_left_pad or 0)), y, width - 2*max(min_padding or 0, scaling * (child_left_pad or 0)), scaling)
                y += l3_height + max(min_spacing or 0, scaling * (vertical_spacing or 0))
        elif level == 3:
            l3_row = node['row']
            l1_num = str(l3_row[level1_col])
            l2_num = str(l3_row[level2_col])
            l3_num = str(l3_row[level3_col])
            l3_text = f"{l1_num}.{l2_num}.{l3_num} {l3_row[subsubcapability_col]}"
            height = max(l3_box_min_height or 0, scaling * ((l3_box_text_height or 0) + 2*(box_padding or 0)))
            add_colored_box(
                slide, left, top, width, height,
                l3_text, args.colorFillLevel2, args.borderColor, 1, args.fontSizeLevel2 - 2, False, args.textColorLevel2, align_left_top=True
            )

    # 4. Layout Level 1 boxes horizontally, fit all on slide
    unique_level1s = list(tree.keys())
    n_level1 = len(unique_level1s)
    total_width = prs_width - 2*margin - (n_level1-1)*margin
    l1_width = max(l1_box_min_width, total_width / n_level1)
    l1_width = float(l1_width)
    margin_f = float(margin)
    prs_width_f = float(prs_width)
    # If too many columns, shrink below min width to fit all
    if l1_width * n_level1 + margin_f * (n_level1 - 1) > prs_width_f - 2*margin_f:
        l1_width = (prs_width_f - 2*margin_f - margin_f*(n_level1-1)) / n_level1
    available_height = prs_height - 2*margin
    l1_lefts = [margin + i*(l1_width + margin) for i in range(n_level1)]
    for i, l1_id in enumerate(unique_level1s):
        l1_node = tree[l1_id]
        natural_h = compute_natural_height(l1_node, 1)
        scaling = min(1.0, available_height / natural_h) if natural_h > 0 else 1.0
        draw_node_scaled(l1_node, 1, l1_lefts[i], margin, l1_width, scaling)

def generate_from_dataframe(df, args, output_path=None):
    # Spaltennamen im DataFrame normalisieren!
    df.columns = [c.strip().upper().replace(' ', '_') for c in df.columns]
    # Mapping für Spaltennamen aus dem Frontend
    col_map = {
        "ID1": "ID_1",
        "ID2": "ID_2",
        "ID3": "ID_3",
        "LEVEL1_CAPABILITY": "LEVEL_1_CAPABILITY",
        "LEVEL2_CAPABILITY": "LEVEL_2_CAPABILITY",
        "LEVEL3_CAPABILITY": "LEVEL_3_CAPABILITY"
    }
    df = df.rename(columns=col_map)
    # Ensure all required columns exist (for manual table input)
    required_cols = [
        'ID_1', 'ID_2', 'ID_3',
        'LEVEL_1_CAPABILITY', 'LEVEL_2_CAPABILITY', 'LEVEL_3_CAPABILITY'
    ]
    for col in required_cols:
        if col not in df.columns:
            df[col] = ""
    # Convert ID columns to string for robust handling, fill NaN/None/empty with '1'
    for col in ['ID_1', 'ID_2', 'ID_3']:
        df[col] = df[col].fillna("1").replace("", "1").astype(str)
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_business_capabilities(slide, df, args, prs.slide_width, prs.slide_height)
    if output_path is None:
        output_dir = os.path.join(os.path.dirname(__file__), 'output')
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, 'Business_Capability_Map.pptx')
    else:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
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
    output_path = args.outputPath
    if output_path is None:
        output_dir = os.path.join(os.path.dirname(__file__), 'output')
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, 'Business_Capability_Map.pptx')
    generate_from_dataframe(df, args, output_path=output_path)

if __name__ == "__main__":
    main()
