import os
import sys
import re
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

def parse_markdown(filepath):
    """Parses the Markdown file and safely captures ALL text."""
    with open(filepath, 'r', encoding='utf-8') as file:
        content = file.read()

    parts = content.split('---')
    if len(parts) < 3:
        frontmatter_text = ""
        body_text = content
        config = {'Theme': 'Light', 'Mode': 'Standard'}
    else:
        frontmatter_text = parts[1].strip()
        body_text = parts[2].strip()
        config = {}
        for line in frontmatter_text.split('\n'):
            if ':' in line:
                key, val = line.split(':', 1)
                config[key.strip()] = val.strip()

    # Safely split slides even if AI uses ## or ###
    raw_slides = re.split(r'\n#+\s+', '\n' + body_text)
    if raw_slides and not raw_slides[0].strip():
        raw_slides.pop(0)
    
    slides = []
    for raw_slide in raw_slides:
        if not raw_slide.strip():
            continue
            
        lines = raw_slide.split('\n')
        title = lines[0].strip().replace('**', '') # Clean bolding from titles
        
        slide_data = {
            'title': title,
            'layout': 'Single', 
            'footer': '',
            'bullets': [],
            'placeholder': '',
            'notes': ''
        }

        for line in lines[1:]:
            line_s = line.strip()
            if not line_s:
                continue
                
            if line_s.startswith('Layout:'):
                slide_data['layout'] = line.split(':', 1)[1].strip()
            elif line_s.startswith('Footer:'):
                slide_data['footer'] = line.split(':', 1)[1].strip()
            elif line_s.startswith('> Notes:'):
                slide_data['notes'] = line.split(':', 1)[1].strip()
            elif line_s.startswith('[PLACEHOLDER') or line_s.startswith('[Placeholder'):
                slide_data['placeholder'] = line_s
            elif line_s.startswith('---'):
                continue 
            else:
                # NEW: Capture ALL text, even if AI forgets the bullet symbol
                indent_level = len(line) - len(line.lstrip())
                level = 0 if indent_level < 3 else 1 
                
                # Strip out any weird markdown formatting
                clean_text = re.sub(r'^[-*+]\s+', '', line.lstrip())
                clean_text = re.sub(r'^\d+\.\s+', '', clean_text)
                clean_text = clean_text.replace('**', '')
                
                if clean_text:
                    slide_data['bullets'].append({'level': level, 'text': clean_text})

        slides.append(slide_data)

    return config, slides

def build_presentation(config, slides, output_filename="output.pptx"):
    """Builds the PowerPoint file and protects against corrupted templates."""
    theme = config.get('Theme', 'Light').lower()
    template_file = 'dark_template.pptx' if 'dark' in theme else 'light_template.pptx'
    
    try:
        prs = Presentation(template_file)
    except Exception as e:
        raise ValueError(f"CRITICAL ERROR: Could not open {template_file}. Ensure it is a valid PowerPoint file, not an empty document.")

    layout_map = {
        'Title': 0, 'Single': 1, 'Data_Heavy': 1, 'Step_by_Step': 1,
        'Divider': 2, 'Split': 3, 'Comparison': 3, 'PICO': 3,
        'Vitals_Grid': 3, 'Algorithm': 5 
    }

    for slide_data in slides:
        layout_name = slide_data['layout']
        layout_idx = layout_map.get(layout_name, 1) 
        
        try:
            slide_layout = prs.slide_layouts[layout_idx]
            slide = prs.slides.add_slide(slide_layout)
        except IndexError:
            # Fallback if your template is missing layouts
            slide = prs.slides.add_slide(prs.slide_layouts[1])

        if slide.shapes.title:
            slide.shapes.title.text = slide_data['title']

        if slide_data['bullets'] and layout_idx in [1, 3]: 
            try:
                body_shape = slide.placeholders[1]
                tf = body_shape.text_frame
                tf.clear() 
                
                for i, bullet in enumerate(slide_data['bullets']):
                    p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
                    p.text = bullet['text']
                    p.level = bullet['level']
            except (IndexError, KeyError):
                pass # Ignores the error if the slide master is missing the text box

        if slide_data['footer']:
            left, top, width, height = Inches(5.5), Inches(6.8), Inches(4), Inches(0.5)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            p = txBox.text_frame.paragraphs[0]
            p.text = slide_data['footer']
            p.font.size = Pt(12)
            p.font.italic = True
            p.alignment = PP_ALIGN.RIGHT

        if slide_data['notes']:
            slide.notes_slide.notes_text_frame.text = slide_data['notes']

    prs.save(output_filename)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(1)
    md_file = sys.argv[1]
    config, slides = parse_markdown(md_file)
    build_presentation(config, slides, md_file.replace('.md', '.pptx'))
