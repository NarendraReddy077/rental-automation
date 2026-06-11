import streamlit as st
import os
import base64

def load_template(filename, directory="templates"):
    """Reads raw file contents from templates directory for dynamic HTML/CSS injection."""
    # Since load_template can be called from ui/ directory, resolve templates directory relative to project root
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(root_dir, directory, filename)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return f.read()
        except Exception:
            return ""
    return ""

def inject_styles_and_bg():
    # Inject css file from templates
    style_css = load_template("style.css")
    if style_css:
        st.markdown(f"<style>{style_css}</style>", unsafe_allow_html=True)

    # Inject background image
    bg_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "bg.png")
    if os.path.exists(bg_path):
        try:
            with open(bg_path, "rb") as f:
                bg_data = f.read()
            bg_base64 = base64.b64encode(bg_data).decode()
            bg_css = f"""
            <style>
            .stApp {{
                background-image: url("data:image/png;base64,{bg_base64}");
                background-size: cover;
                background-position: center;
                background-repeat: no-repeat;
                background-attachment: fixed;
            }}
            </style>
            """
            st.markdown(bg_css, unsafe_allow_html=True)
        except Exception:
            pass
