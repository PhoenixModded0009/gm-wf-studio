import os
import sys
import re
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

def parse_markdown(filepath):
    """Parses Markdown and captures all text/formatting."""
    with open(filepath, 'r', encoding='utf-8') as file:
        content = file.read()

    parts = content.split('---')
    if len(parts) < 3:
        config = {'Theme': 'Light', 'Mode': 'Standard'}
        body_text = content
    else:
        frontmatter_text = parts[1].strip()
        body_text = parts[2].strip()
        config = {line.split(':', 1)[0].strip(): line.split(':', 1)[1].strip() 
                  for line in frontmatter_text.split('\n') if ':' in line}

    raw_slides = re.split(r'\n#+\s+', '\n' + body_text)
    if raw_slides and not raw_slides[0].strip():
        raw_slides.pop(0)
    
    slides = []
    for raw_slide in raw_slides:
        if not raw_slide.strip(): continue
        lines = raw_slide.split('\n')
        slide_data = {
            'title': lines[0].strip().replace('**', ''),
            'layout': 'Single', 'footer': '', 'bullets': [], 'placeholder': '', 'notes': ''
        }
        for line in lines[1:]:
            line_s = line.strip()
            if not line_s: continue
            if line_s.startswith('Layout:'): slide_data['layout'] = line.split(':', 1)[1].strip()
            elif line_s.startswith('Footer:'): slide_data['footer'] = line.split(':', 1)[1].strip()
            elif line_s.startswith('> Notes:'): slide_data['notes'] = line.split(':', 1)[1].strip()
            elif line_s.startswith('[PLACEHOLDER'): slide_data['placeholder'] = line_s
            else:
                level = 0 if len(line) - len(line.lstrip()) < 3 else 1
                clean_text = re.sub(r'^[-*+]\s+|^d+\.\s+', '', line.lstrip()).replace('**', '')
                if clean_text: slide_data['bullets'].append({'level': level, 'text': clean_text})
        slides.append(slide_data)
    return config, slides

def build_presentation(config, slides, output_filename="output.pptx"):
    """Builds PPTX with dynamic themes and red-alert placeholders."""
    theme = config.get('Theme', 'Light').lower()
    # Flexibility: Look for 'blue_template.pptx', 'red_template.pptx', etc.
    template_file = f"{theme}_template.pptx" if os.path.exists(f"{theme}_template.pptx") else 'light_template.pptx'
    
    try:
        prs = Presentation(template_file)
    except:
        raise ValueError(f"Template '{template_file}' not found. Upload it to GitHub.")

    layout_map = {'Title': 0, 'Single': 1, 'Divider': 2, 'Split': 3, 'Comparison': 3, 'Algorithm': 5}

    for slide_data in slides:
        idx = layout_map.get(slide_data['layout'], 1)
        slide = prs.slides.add_slide(prs.slide_layouts[idx])
        
        if slide.shapes.title:
            slide.shapes.title.text = slide_data['title']

        if slide_data['bullets'] and idx in [1, 3]:
            tf = slide.placeholders[1].text_frame
            tf.clear()
            for b in slide_data['bullets']:
                p = tf.add_paragraph()
                p.text, p.level = b['text'], b['level']

        # THE RED ALERT STAMP
        if slide_data['placeholder']:
            box = slide.shapes.add_textbox(Inches(1), Inches(3.5), Inches(8), Inches(1))
            p = box.text_frame.paragraphs[0]
            p.text = f"🚨 ACTION: INSERT {slide_data['placeholder'].upper()} 🚨"
            p.font.size, p.font.bold = Pt(24), True
            p.font.color.rgb = RGBColor(255, 0, 0)
            p.alignment = PP_ALIGN.CENTER

        if slide_data['footer']:
            tx = slide.shapes.add_textbox(Inches(5.5), Inches(6.8), Inches(4), Inches(0.5))
            tx.text_frame.text = slide_data['footer']
        
        if slide_data['notes']:
            slide.notes_slide.notes_text_frame.text = slide_data['notes']

    prs.save(output_filename)
