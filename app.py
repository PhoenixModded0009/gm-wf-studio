import streamlit as st
import google.generativeai as genai
import os
from generate_ppt import parse_markdown, build_presentation 

st.set_page_config(page_title="GM-WF Studio", page_icon="🩺", layout="wide")
st.title("🩺 GM-WF Studio v4.2")

# --- Sidebar Controls ---
with st.sidebar:
    st.header("⚙️ Settings")
    density = st.selectbox("Density", ["Standard", "Minimalist", "Detailed"])
    mode = st.selectbox("Mode", ["Research_Update", "Theory_Topic", "Case_Presentation", "Practical_Approach"])
    theme_choice = st.text_input("Theme (e.g., Light, Dark, Blue, Red)", value="Light")

# --- Step 1: Brain Dump ---
raw_notes = st.text_area("Paste your clinical notes here:", height=200)

if st.button("1. Generate Draft Markdown", type="primary"):
    if not raw_notes:
        st.error("Enter notes first!")
    else:
        with st.spinner("AI is drafting..."):
            api_key = st.secrets["GEMINI_API_KEY"]
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-2.0-flash') # Updated to latest stable
            
            prompt = f"Act as an Academic Architect. Convert these notes to Markdown for a PPT engine. Theme: {theme_choice}, Mode: {mode}, Density: {density}. Rules: No asterisks, use '# Slide Title', 'Layout: [Type]', and '- [Content]'. Notes: {raw_notes}"
            
            response = model.generate_content(prompt)
            st.session_state['draft'] = response.text.replace("```markdown", "").replace("```", "").strip()

# --- Step 2: Review & Edit ---
if 'draft' in st.session_state:
    st.markdown("### ✍️ Edit Your Slides Before Compiling")
    edited_md = st.text_area("Live Slide Code:", value=st.session_state['draft'], height=400)
    
    if st.button("2. Compile to PowerPoint"):
        with open("temp.md", "w", encoding="utf-8") as f:
            f.write(edited_md)
        
        config, slides = parse_markdown("temp.md")
        build_presentation(config, slides, "final.pptx")
        
        with open("final.pptx", "rb") as f:
            st.session_state['ppt_ready'] = f.read()
        st.success("PowerPoint Built!")

# --- Step 3: Download ---
if 'ppt_ready' in st.session_state:
    st.download_button("📥 Download Presentation", data=st.session_state['ppt_ready'], file_name="Presentation.pptx")
