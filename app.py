import streamlit as st
import google.generativeai as genai
import os
from generate_ppt import parse_markdown, build_presentation 

# --- UI Setup ---
st.set_page_config(page_title="GM-WF Studio", page_icon="🩺", layout="wide")
st.title("🩺 GM-WF Studio v4.1")
st.markdown("Automated Pediatric Presentation Engine")

# --- Control Panel (Sidebar) ---
with st.sidebar:
    st.header("⚙️ Settings")
    density = st.selectbox("Density", ["Standard", "Minimalist", "Detailed"])
    audience = st.selectbox("Audience", ["Interns", "Co-Residents", "Attendings", "Thesis Committee"])
    slide_limit = st.slider("Max Slides", 3, 25, 10)
    mode = st.selectbox("Mode", ["Research_Update", "Theory_Topic", "Case_Presentation", "Practical_Approach", "Journal_Club", "Morbidity_Mortality"])
    theme = st.radio("Theme", ["Light", "Dark"])
    
# --- Main Input ---
raw_notes = st.text_area("Paste your raw clinical brain dump here:", height=250)

if st.button("Generate Presentation", type="primary"):
    if not raw_notes:
        st.error("Please enter your clinical notes first.")
    else:
        with st.spinner("🧠 AI is structuring your clinical logic..."):
            try:
                # 1. Authenticate with Google Gemini
                api_key = st.secrets["GEMINI_API_KEY"]
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel('gemini-2.5-flash')
                
                # 2. Construct the Master Prompt
                prompt = f"""
                Act as an Expert Medical Presentation Architect. Convert these notes to a strict Markdown script for a PowerPoint engine.
                
                CONTROL PANEL:
                Density: {density} (Strictly enforce bullet limits based on this: Minimalist=3, Standard=5, Detailed=7 nested).
                Audience: {audience}
                Slide Limit: {slide_limit}
                
                CRITICAL RULES:
                1. Start with this exact YAML block:
                ---
                Ratio: 16:9
                Theme: {theme}
                Accent: PICU Blue
                Mode: {mode}
                ---
                2. Use the exact slide syntax:
                # Slide Title
                Layout: [Title, Single, Split, Comparison, Algorithm, Vitals_Grid, PICO, Data_Heavy, Step_by_Step, Divider]
                Footer: [Citation or blank]
                
                - Bullet 1
                
                [PLACEHOLDER: description]
                > Notes: [Speaker notes]
                
                RAW NOTES:
                {raw_notes}
                """
                
                # 3. Get AI Response
                response = model.generate_content(prompt)
                markdown_output = response.text
                
                # 4. Save Temp File & Run Your Engine
                with open("temp.md", "w", encoding="utf-8") as f:
                    f.write(markdown_output)
                    
                config, slides = parse_markdown("temp.md")
                output_file = "final_presentation.pptx"
                build_presentation(config, slides, output_file)
                
                st.success("✅ Presentation built successfully!")
                
                # 5. Provide Download Button
                with open(output_file, "rb") as file:
                    st.download_button(
                        label="📥 Download PowerPoint File",
                        data=file,
                        file_name="GM_WF_Presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                
                # Cleanup temp files
                os.remove("temp.md")
                os.remove(output_file)
                
            except Exception as e:

                st.error(f"An error occurred: {e}")
