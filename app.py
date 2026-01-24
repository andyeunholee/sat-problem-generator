import streamlit as st
import google.generativeai as genai
import os
from pathlib import Path
import json
import re
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Page Configuration
st.set_page_config(
    page_title="Elite Prep | SAT Analysis & Adaptive Practice",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Brand Colors & Styles
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
    
    html, body, [class*="css"]  {
        font-family: 'Inter', sans-serif;
        color: #1f2937;
        background-color: #f3f4f6;
    }
    
    /* Main Container */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 5rem;
        max-width: 1000px;
    }

    /* Header Styling */
    .main-header {
        font-family: 'Inter', sans-serif;
        color: #0C1E41; /* Elite Navy */
        text-align: center;
        font-weight: 800;
        font-size: 2.2rem;
        margin-bottom: 0.5rem;
        letter-spacing: -0.02em;
    }
    .sub-header {
        color: #6b7280;
        text-align: center;
        margin-bottom: 3rem;
        font-size: 1.1rem;
        font-weight: 400;
    }
    
    /* Card Styling */
    .stCard {
        background-color: white;
        padding: 2rem;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        margin-bottom: 2rem;
        border: 1px solid #e5e7eb;
    }

    /* Button Styling */
    div.stButton > button {
        background-color: #0C1E41;
        color: white;
        font-weight: 600;
        border-radius: 8px;
        padding: 0.75rem 1.5rem;
        border: none;
        transition: all 0.2s ease;
        width: 100%;
        box-shadow: 0 4px 6px rgba(12, 30, 65, 0.2);
    }
    div.stButton > button:hover {
        background-color: #1a3a6e;
        transform: translateY(-1px);
        box-shadow: 0 6px 8px rgba(12, 30, 65, 0.3);
    }

    /* Practice Topic Buttons (Secondary) */
    .topic-btn {
        border: 1px solid #0C1E41 !important;
        background-color: white !important;
        color: #0C1E41 !important;
    }
    
    /* Status Indicators */
    .status-success { color: #166534; font-weight: 500;}
    .status-error { color: #991b1b; font-weight: 500;}

</style>
""", unsafe_allow_html=True)

# Helper: Upload file to Gemini (File API)
def upload_to_gemini(path, mime_type=None):
    try:
        file = genai.upload_file(path, mime_type=mime_type)
        return file
    except Exception as e:
        print(f"Error uploading {path}: {e}")
        return None

# Helper: Load Gemini Model
def get_gemini_response(input_prompt, content_parts):
    try:
        api_key = os.getenv("GOOGLE_API_KEY")
        if not api_key:
            if 'api_key' in st.session_state and st.session_state.api_key:
                api_key = st.session_state.api_key
            else:
                st.error("Google API Key not found.")
                return None
        
        genai.configure(api_key=api_key)
        
        # Using Gemini 3.0 Pro Preview as requested
        model = genai.GenerativeModel('gemini-3-pro-preview')
        
        full_payload = [input_prompt] + content_parts
        
        generation_config = genai.types.GenerationConfig(
            temperature=0.2,
            max_output_tokens=8192,
        )

        response = model.generate_content(full_payload, generation_config=generation_config)
        
        # Safe access to text
        try:
            return response.text
        except ValueError:
            # Handle cases where response is blocked or empty
            finish_reason = response.candidates[0].finish_reason if response.candidates else "UNKNOWN"
            st.error(f"Generation stopped. Finish Reason: {finish_reason}")
            # If MAX_TOKENS (2), try to return what we have
            if finish_reason == 2 and response.candidates and response.candidates[0].content.parts:
                return response.candidates[0].content.parts[0].text
            return None
    except Exception as e:
        st.error(f"Error calling Gemini API: {str(e)}")
        return None

# Helper: Extract JSON from text
def extract_json_topics(text):
    try:
        # Find JSON block using regex
        match = re.search(r'```json\s*(\{.*?\})\s*```', text, re.DOTALL)
        if match:
            return json.loads(match.group(1))
        else:
            # Try finding just brace content if no markdown tags
            match = re.search(r'\{.*\}', text, re.DOTALL)
            if match:
                return json.loads(match.group(0))
    except Exception as e:
        print(f"JSON Parse Error: {e}")
    return None

# Helper: Force Formatting for Options
def ensure_formatting(text):
    if not text: return ""
    # Regex to find A) B) C) D) that are NOT at start of line
    # We look for " A)", " B)" etc and force newline
    # Pattern: Space + [A-D] + ) + Space
    pattern = r"(\s)([A-D]\)) "
    
    # Replace with \n\n<Option>
    formatted_text = re.sub(pattern, r"\n\n\2 ", text)
    return formatted_text

# Helper: Convert Markdown to Simple HTML for Word
def convert_markdown_to_html(md_text):
    if not md_text: return ""
    html = md_text
    
    # Basic Headers
    html = re.sub(r'### (.*)', r'<h3>\1</h3>', html)
    html = re.sub(r'## (.*)', r'<h2>\1</h2>', html)
    html = re.sub(r'# (.*)', r'<h1>\1</h1>', html)
    # Bold
    html = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', html)
    # Newlines
    html = html.replace('\n', '<br>')
    # Horizontal Rule
    html = html.replace('---', '<hr>')
    
    # Wrapper for clean Word import + MathJax for LaTeX
    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <script>
        window.MathJax = {{
          tex: {{
            inlineMath: [['$', '$'], ['\\\\(', '\\\\)']]
          }},
          svg: {{
            fontCache: 'global'
          }}
        }};
        </script>
        <script type="text/javascript" id="MathJax-script" async
          src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js">
        </script>
        <style>
            body {{ font-family: 'Arial', sans-serif; line-height: 1.6; padding: 20px; max-width: 800px; margin: 0 auto; }}
            h1, h2, h3 {{ color: #0C1E41; }}
            .question-block {{ margin-bottom: 20px; }}
        </style>
    </head>
    <body>
        {html}
    </body>
    </html>
    """
    return full_html

# Helper: Load Local Resources (JPGs)
def load_local_resources():
    resources = []
    content_parts = []
    current_dir = Path(".")
    
    # Core Textbooks
    file_map = {
        "SAT Math Textbook": "SAT-Math.jpg",
        "SAT English Textbook": "SAT-English.jpg",
        "Vocabulary List": ["Vocabulary.jpg", "Vocabulary.pdf"]
    }

    # DSAT Test Packets
    dsat_files = sorted(list(current_dir.glob("DSAT*.jpg")))
    
    # Load Core Files
    for label, filenames in file_map.items():
        if isinstance(filenames, str):
            filenames = [filenames]
        for fname in filenames:
            file_path = current_dir / fname
            if file_path.exists():
                try:
                    # Check session state/cache here normally, but simple upload for now
                    # (In prod, cache the File API URIs)
                    uploaded_file = upload_to_gemini(str(file_path))
                    if uploaded_file:
                        content_parts.append(uploaded_file)
                        resources.append(f"✅ {label} Loaded")
                        break
                except Exception:
                     resources.append(f"❌ Error loading {fname}")

    # Load Test Packets
    if dsat_files:
        resources.append(f"📚 {len(dsat_files)} Test Packets Detected")
        for dsat_path in dsat_files:
            try:
                uploaded_file = upload_to_gemini(str(dsat_path))
                if uploaded_file:
                    content_parts.append(uploaded_file)
            except Exception:
                pass

    return resources, content_parts, len(list(current_dir.glob("SAT*.jpg"))) # Return count for core

# --- Sidebar ---
with st.sidebar:
    # Logo & Brand (Centered)
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        if Path("Elite-Logo.png").exists():
            st.image("Elite-Logo.png", use_container_width=True)
    
    st.markdown("<h2 style='text-align: center; color: #0C1E41; margin-top: -10px;'>Elite Prep</h2>", unsafe_allow_html=True)
    
    st.markdown("---")
    
    # API Key (Hidden Label)
    env_key = os.getenv("GOOGLE_API_KEY")
    if env_key:
        st.success("✅ API Key Loaded")
    else:
        api_key_input = st.text_input("Enter Google API Key", type="password")
        if api_key_input:
            st.session_state.api_key = api_key_input
    
    st.markdown("---")
    st.markdown("### 📂 Upload Student Report")
    uploaded_files = st.file_uploader(
        "Upload Diagnostic PDF(s)", 
        type=['pdf', 'jpg', 'png'], 
        accept_multiple_files=True
    )
    
    st.markdown("---")
    st.markdown("### 📚 System Resources")
    
    # Quick UI Status check
    current_dir = Path(".")
    core_files = ["SAT-Math.jpg", "SAT-English.jpg", "Vocabulary.jpg"]
    found_core = 0
    for f in core_files:
        if (current_dir / f).exists():
            st.caption(f"✅ {f} detected")
            found_core += 1
        else:
            st.caption(f"⚪ {f} not found")
    
    dsat_count = len(list(current_dir.glob("DSAT*.jpg")))
    if dsat_count > 0:
        st.caption(f"📚 {dsat_count} Test Packets detected")

# --- Main Content ---
st.markdown("<h1 class='main-header'>Elite Prep | SAT Analysis & Improvement Plan | SAT Adaptive Practice Questions</h1>", unsafe_allow_html=True)


# Initialization
if 'analysis_done' not in st.session_state:
    st.session_state.analysis_done = False
if 'full_report' not in st.session_state:
    st.session_state.full_report = ""
if 'weak_topics' not in st.session_state:
    st.session_state.weak_topics = None
if 'practice_result' not in st.session_state:
    st.session_state.practice_result = ""
if 'all_topics' not in st.session_state:
    st.session_state.all_topics = None
if 'manual_practice_result' not in st.session_state:
    st.session_state.manual_practice_result = ""

# Main Tabs
main_tab1, main_tab2 = st.tabs(["📊 Student Analysis Report", "📚 Topic Bank (Manual Gen)"])

# ==========================================
# TAB 1: Student Analysis & Adaptive Practice
# ==========================================
with main_tab1:
    # 1. Report Generation Phase
    if not st.session_state.analysis_done:
        if uploaded_files:
            st.success(f"📂 {len(uploaded_files)} Report(s) Ready.")
            
            if st.button("🚀 Generate Analysis Report"):
                with st.spinner("Analyzing Student Data & Textbooks... (This may take 1-2 mins)"):
                    
                    # Load Assets
                    res_status, context_parts, _ = load_local_resources()
                    
                    # Load User Files
                    student_parts = []
                    for up_file in uploaded_files:
                        student_parts.append({
                            "mime_type": up_file.type,
                            "data": up_file.getvalue()
                        })
                    
                    all_content = context_parts + student_parts
                    
                    # Report Prompt
                    prompt = """
                    You are 'Elite Academy's Senior SAT Consultant.
                    
                    **Inputs provided:**
                    1. **Elite Academy Textbooks & Test Packets** (Context). 
                    2. **Student Diagnostic Test Results** (Target).
                    
                    **TASK:** Create a "SAT Analysis & Improvement Plan" report.
                    
                    **SECTIONS (Output strictly in Markdown):**
                    
                    1.  **Key Weaknesses Analysis**: Detailed breakdown of Reading/Writing and Math gaps.
                    2.  **Curriculum Mapping**: Map weaknesses to **specific Elite Academy Textbook Chapters**. Create a table.
                    3.  **Action Plan**: Target score, tutoring hours/frequency, weekly schedule.
                    4.  **Vocabulary List**: 50 customized words (numbered list).
                    
                    **IMPORTANT:** 
                    Do NOT generate practice questions yet. We will do that in a separate step.
                    
                    **CRITICAL OUTPUT INSTRUCTION:**
                    At the very end of your response, output a **RAW JSON block** Listing the specific "Weak Lesson Subjects" identified in the Curriculum Mapping.
                    The JSON must look exactly like this:
                    ```json
                    {
                        "English": ["Chapter 3: Grammar", "Chapter 5: Inference"],
                        "Math": ["Heart of Algebra", "Problem Solving"]
                    }
                    ```
                    """
                    
                    response_text = get_gemini_response(prompt, all_content)
                    
                    if response_text:
                        st.session_state.full_report = response_text
                        st.session_state.analysis_done = True
                        
                        # Parse JSON
                        topics = extract_json_topics(response_text)
                        if topics:
                            st.session_state.weak_topics = topics
                        else:
                            st.warning("Could not auto-detect weak topics for the practice menu. (JSON Parsing Failed)")
                            # Fallback default
                            st.session_state.weak_topics = {"English": ["General Practice"], "Math": ["General Practice"]}
                        
                        st.rerun() # Reload to show report
                    else:
                        st.error("Analysis Failed. Please try again.")
        else:
            st.info("👋 Please upload Student Reports in the Sidebar to start Analysis.")

    # 2. Display Report & Interactive Practice
    if st.session_state.analysis_done:
        st.markdown("### 📊 Analysis Report")
        
        # Remove the JSON block from display if possible for cleanliness (Optional, keeping it simple for now)
        clean_report = re.sub(r'```json\s*\{.*?\}\s*```', '', st.session_state.full_report, flags=re.DOTALL)
        
        with st.expander("📄 View Full Analysis Report", expanded=True):
            st.markdown(clean_report)
            st.markdown("---")
            st.info("Click the **Copy icon (📄)** top-right of the code block to copy for Word.")
            st.code(clean_report, language="markdown")

        st.markdown("---")
        st.markdown("## 🎯 Adaptive Practice Station")
        st.markdown("Select a subject and topic below to generate **10 targeted practice questions** instantly.")

        if st.session_state.weak_topics:
            # Layout
            tab1, tab2 = st.tabs(["📘 SAT English Practice", "📐 SAT Math Practice"])
            
            # English Tab
            with tab1:
                st.subheader("Weak English Topics")
                topics = st.session_state.weak_topics.get("English", [])
                cols = st.columns(3)
                for idx, topic in enumerate(topics):
                    if cols[idx % 3].button(f"Generate: {topic}", key=f"eng_{idx}"):
                        # Generate Logic
                        with st.spinner(f"Generating 10 questions for {topic}..."):
                            res_status, context_parts, _ = load_local_resources()
                            
                            practice_prompt = f"""
                            Create **10 SAT Practice Questions** for the topic: **'{topic}'**.
                            
                            **Style Manual:**
                            - Browse the provided "Elite Academy Textbooks" and "Test Packets".
                            - Mimic the difficulty and style of the questions found there.
                            
                            **CRITICAL FORMATTING:**
                            - **Multiple Choice Options**: MUST be on separate lines.
                              Example:
                              A) Option 1
                              B) Option 2
                              C) Option 3
                              D) Option 4
                            
                            **Answer Key:**
                            Provide the Answer Key and brief explanations at the very bottom.
                            
                            Output in Markdown.
                            """
                            q_response = get_gemini_response(practice_prompt, context_parts)
                            if q_response:
                                st.session_state.practice_result = f"### 📘 Practice Set: {topic}\n\n" + q_response
            
            # Math Tab
            with tab2:
                st.subheader("Weak Math Topics")
                topics = st.session_state.weak_topics.get("Math", [])
                cols = st.columns(3)
                for idx, topic in enumerate(topics):
                    if cols[idx % 3].button(f"Generate: {topic}", key=f"math_{idx}"):
                         with st.spinner(f"Generating 10 questions for {topic}..."):
                            res_status, context_parts, _ = load_local_resources()
                            
                            practice_prompt = f"""
                            Create **10 SAT Math Practice Questions** for the topic: **'{topic}'**.
                            
                            **Instructions:**
                            - Use LaTeX for match equations.
                            - Browse the provided "Elite Academy Textbooks" (Math) and "Test Packets".
                            - Mimic the difficulty and style.
                            
                            **CRITICAL FORMATTING:**
                            - **Multiple Choice Options**: MUST be on separate lines.
                              Example:
                              A) Option 1
                              B) Option 2
                              C) Option 3
                              D) Option 4
                            
                            **Answer Key:**
                            Provide the Answer Key and brief explanations at the very bottom.
                            
                            Output in Markdown.
                            """
                            q_response = get_gemini_response(practice_prompt, context_parts)
                            if q_response:
                                st.session_state.practice_result = f"### 📐 Practice Set: {topic}\n\n" + q_response

        # Display Generated Questions
        if st.session_state.practice_result:
            # Force Formatting
            st.session_state.practice_result = ensure_formatting(st.session_state.practice_result)
            
            st.markdown("---")
            st.markdown(st.session_state.practice_result)
            st.markdown("### 📋 Copy Practice Set")
            st.info("Select the text above to copy.")
            
            if st.button("Clear Practice Result"):
                st.session_state.practice_result = ""
                st.rerun() # Reload

# ==========================================
# TAB 2: Topic Bank (Manual)
# ==========================================
with main_tab2:
    st.markdown("### 📚 Elite Topic Bank")
    st.markdown("Select any Lesson Subject from the textbooks to generate new practice questions manually.")
    
    # 1. Hardcoded Topic Lists (Optimized)
    math_topics = [
        "Chapter 2: Solving Linear Equations", "Chapter 3: Systems of Linear Equations", "Chapter 4: Systems of Linear Equations in Context",
        "Chapter 5: Linear Functions", "Chapter 6: Graphs of Lines", "Chapter 7: Linear Models", "Chapter 8: Linear Inequalities",
        "Chapter 9: Absolute Value", "Chapter 10: Ratios", "Chapter 11: Units", "Chapter 12: Percentages", "Chapter 13: Percent Change",
        "Chapter 14: Rates", "Chapter 15: Lines and Angles", "Chapter 16: Similar Triangles", "Chapter 17: Measures of Center",
        "Chapter 18: Measures of Spread", "Chapter 19: Tables", "Chapter 20: Bar Graphs", "Chapter 21: Time-Series Graphs",
        "Chapter 22: Probability", "Chapter 23: Scatterplots", "Chapter 24: Inferences from Sample Statistics", "Chapter 25: Statistical Claims",
        "Chapter 26: Right Triangles", "Chapter 27: Trigonometry", "Chapter 28: Area and Volume", "Chapter 29: Circles",
        "Chapter 30: Completing the Square", "Chapter 31: Equations of Circles", "Chapter 32: Quadratic Equations", "Chapter 33: Rational Functions",
        "Chapter 34: Exponents and Radicals", "Chapter 35: Polynomials", "Chapter 36: Graphs of Functions", "Chapter 37: Nonlinear Equations in Context",
        "Chapter 38: Composite Functions", "Chapter 39: Linear and Exponential Models", "Chapter 40: Equivalent Expressions"
    ]

    english_topics = [
        "Chapter 2: Central Ideas", "Chapter 3: Parts of Speech", "Chapter 4: Phrases", "Chapter 5: Active Reading",
        "Chapter 6: Clauses", "Chapter 7: Appositives", "Chapter 8: Command of Evidence (Textual)", "Chapter 9: Subject-Verb Agreement",
        "Chapter 10: Inferences", "Chapter 11: Verb Tense and Time Reference", "Chapter 12: Words in Context",
        "Chapter 13: Possessive Nouns and Possessive Determiners", "Chapter 14: Parentheticals", "Chapter 15: Modifier Placement",
        "Chapter 16: Text Structure and Purpose", "Chapter 17: Transitions", "Chapter 18: Informational Graphics",
        "Chapter 19: Rhetorical Synthesis", "Chapter 20: Cross-Text Connections", "Chapter 21: Punctuation"
    ]

    # Pre-load topics if not present
    if not st.session_state.all_topics:
        st.session_state.all_topics = {
            "Math": math_topics,
            "English": english_topics
        }

    # 2. Selection & Generation
    col_m, col_e = st.columns(2)
    
    with col_m:
        st.subheader("📐 Math Topics")
        math_list = st.session_state.all_topics.get("Math", math_topics) # Fallback to hardcoded
        selected_math = st.selectbox("Select Math Topic", ["-- Select --"] + math_list)
        if st.button("Generate Math Questions", disabled=(selected_math=="-- Select --")):
            with st.spinner(f"Generating 10 Math Questions for {selected_math}..."):
                res_status, context_parts, _ = load_local_resources()
                prompt = f"""
                Create **10 SAT Math Practice Questions** for the topic: **'{selected_math}'**.
                
                **INSTRUCTIONS:**
                - Mimic the exact difficulty and style of the "Elite Academy Textbooks".
                - Use LaTeX for all math equations.
                
                **CRITICAL STRUCTURE:**
                1. **Questions 1-10**: List the questions clearly.
                   - **Multiple Choice**: Put each option (A, B, C, D) on a NEW LINE.
                2. **--- PAGE BREAK ---**
                3. **Answer Key & Explanations**: You MUST provide this section at the very end.
                   - Format: "1. A - Explanation..."
                
                Output in clean Markdown.
                """
                res = get_gemini_response(prompt, context_parts)
                if res:
                    st.session_state.manual_practice_result = f"### 📐 Manual Set: {selected_math}\n\n" + res
    
    with col_e:
        st.subheader("📘 English Topics")
        eng_list = st.session_state.all_topics.get("English", english_topics) # Fallback to hardcoded
        selected_eng = st.selectbox("Select English Topic", ["-- Select --"] + eng_list)
        if st.button("Generate English Questions", disabled=(selected_eng=="-- Select --")):
                with st.spinner(f"Generating 10 English Questions for {selected_eng}..."):
                    res_status, context_parts, _ = load_local_resources()
                    prompt = f"""
                    Create **10 SAT English Practice Questions** for the topic: **'{selected_eng}'**.
                    
                    **INSTRUCTIONS:**
                    - Mimic the exact passage length and question style of the "Elite Academy Textbooks" / DSAT.
                    
                    **CRITICAL STRUCTURE:**
                    1. **Questions 1-10**: List the questions clearly.
                       - **Multiple Choice**: Put each option (A, B, C, D) on a NEW LINE.
                    2. **--- PAGE BREAK ---**
                    3. **Answer Key & Explanations**: You MUST provide this section at the very end.
                       - Format: "1. A - Explanation..."
                    
                    Output in clean Markdown.
                    """
                    res = get_gemini_response(prompt, context_parts)
                    if res:
                        st.session_state.manual_practice_result = f"### 📘 Manual Set: {selected_eng}\n\n" + res

    # Display Result
    if st.session_state.manual_practice_result:
        # Force Formatting
        st.session_state.manual_practice_result = ensure_formatting(st.session_state.manual_practice_result)
        
        st.markdown("---")
        
        # Header + Copy Button Layout
        r_col1, r_col2 = st.columns([5, 2])
        with r_col1:
            st.markdown("### 📝 Generated Practice Set")
        with r_col2:
            # Download Button for Word
            html_data = convert_markdown_to_html(st.session_state.manual_practice_result)
            st.download_button(
                label="💾 Download for HTML",
                data=html_data,
                file_name="Elite_Practice_Set.html",
                mime="text/html"
            )

        st.markdown(st.session_state.manual_practice_result)
        
        if st.button("Clear Manual Result"):
            st.session_state.manual_practice_result = ""
            st.rerun()

