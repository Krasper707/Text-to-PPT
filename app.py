# app.py

import streamlit as st
from openai import OpenAI
import json
import requests  # <-- NEW: Our tool for making HTTP requests
from pptx import Presentation  # <-- Add this
from pptx.util import Inches   # <-- Add this (optional, but good for later)
import io                      # <-- Add this


# --- Page Configuration ---
st.set_page_config(
    page_title="Your Text, Your Style",
    page_icon="✨",
    layout="centered"
)

# --- Session State Initialization ---
# We need to store the list of models so we don't lose it on every interaction.
if 'models' not in st.session_state:
    st.session_state.models = []
if 'selected_model' not in st.session_state:
    st.session_state.selected_model = None

# --- LLM and Logic Functions ---

def analyze_template(template_file) -> tuple[Presentation, dict]:
    """
    Analyzes an uploaded PowerPoint template to identify key slide layouts.

    Args:
        template_file: The file-like object from Streamlit's file uploader.

    Returns:
        A tuple containing the Presentation object and a dictionary mapping
        layout names ('title_slide', 'content_slide') to their layout objects.
    """
    # The python-pptx library needs a file-like object, which Streamlit provides.
    prs = Presentation(template_file)
    
    layouts = {}
    
    # We will try to identify the two most important layouts:
    # 1. The Title Slide Layout
    # 2. The Title and Content Layout
    
    # Heuristic for finding the 'Title and Content' layout (usually has a title and a large body placeholder)
    for i, layout in enumerate(prs.slide_layouts):
        has_title = False
        has_body = False
        for ph in layout.placeholders:
            if 'Title' in ph.name or ph.placeholder_format.type == 1: # 1 is Title type
                has_title = True
            if 'Content' in ph.name or 'Body' in ph.name or ph.placeholder_format.type == 2: # 2 is Body type
                has_body = True
        if has_title and has_body:
            layouts['content_slide'] = layout
            break # We found one, let's stop here
    
    # If we couldn't find a content slide, fall back to a common default (layout index 1)
    if 'content_slide' not in layouts and len(prs.slide_layouts) > 1:
        layouts['content_slide'] = prs.slide_layouts[1]
        
    # Heuristic for 'Title Slide' layout (often the first one, index 0)
    # Or one that has a title and maybe a subtitle, but no main body.
    if len(prs.slide_layouts) > 0:
        layouts['title_slide'] = prs.slide_layouts[0]

    # A final check to ensure we have *something* to work with
    if 'content_slide' not in layouts:
        st.warning("Could not definitively identify a 'Title and Content' layout. Using the first available layout.")
        layouts['content_slide'] = prs.slide_layouts[0]
        
    return prs, layouts

@st.cache_data(show_spinner="Fetching available models...")
def get_available_models(aipipe_token: str) -> list[str]:
    """
    Fetches the list of available models from the AI Pipe API.
    Returns a list of model ID strings.
    """
    if not aipipe_token:
        return []
    
    url = "https://aipipe.org/openrouter/v1/models"
    headers = {
        "Authorization": f"Bearer {aipipe_token}",
    }
    
    try:
        response = requests.get(url, headers=headers)
        # If the request was successful
        if response.status_code == 200:
            data = response.json()
            # The model data is in the 'data' key, we extract the 'id' from each entry
            model_ids = [model['id'] for model in data.get('data', [])]
            return sorted(model_ids) # Return a sorted list
        else:
            # If the token is bad or something else goes wrong
            st.error(f"Failed to fetch models. Status code: {response.status_code} - {response.text}")
            return []
    except Exception as e:
        st.error(f"An error occurred while fetching models: {e}")
        return []

# def generate_slide_content(text: str, guidance: str, aipipe_token: str, model_name: str) -> dict | None:
#     # This function remains the same as before, no changes needed here.
#     try:
#         client = OpenAI(api_key=aipipe_token, base_url="https://aipipe.org/openai/v1")
#     except Exception as e:
#         st.error(f"Failed to initialize the AI client. Error: {e}")
#         return None

#     system_prompt = """
#     You are an expert presentation creator... [Your previous system prompt here] ...
#     {"slides": [{"title": "Slide Title", "content": ["Bullet point 1."], "speaker_notes": "Notes."}]}
#     """
#     user_content = f"### USER GUIDANCE:\n{guidance}\n\n### SOURCE TEXT:\n{text}"
    
#     st.info(f"Sending request to model: {model_name}...")
#     try:
#         response = client.chat.completions.create(
#             model=model_name,
#             messages=[
#                 {"role": "system", "content": system_prompt},
#                 {"role": "user", "content": user_content}
#             ],
#             response_format={"type": "json_object"}
#         )
#         response_content = response.choices[0].message.content
#         return json.loads(response_content)
#     except Exception as e:
#         st.error(f"An error occurred while communicating with the AI model: {e}")
#         return None

def generate_slide_content(text: str, guidance: str, aipipe_token: str, model_name: str) -> dict | None:
    """
    Uses an LLM via AI Pipe to structure text into slide content.
    This version builds the HTTP request manually to bypass any SDK issues.
    """
    st.info(f"Sending request to model: {model_name} via manual POST request...")

    # The endpoint for chat completions on the OpenRouter path
    url = "https://aipipe.org/openrouter/v1/chat/completions"

    # Define the headers, exactly like in our working get_models function
    headers = {
        "Authorization": f"Bearer {aipipe_token}",
        "Content-Type": "application/json",
        # (Optional but good practice) Add a Referer header
        "HTTP-Referer": "https://github.com/mshakir-io/your-text-your-style-app" # Change to your repo
    }

    # This is our master instruction to the AI.
    system_prompt = """
    You are an expert presentation creator. Your task is to analyze the following text and user guidance to structure it into a series of presentation slides.
    - Break down the content logically.
    - The number of slides should be reasonable for the amount of text provided.
    - Ensure each slide's content is concise and formatted as bullet points.
    - Create a title for each slide.
    - Generate brief, helpful speaker notes for each slide.
    - You MUST output your response ONLY as a single, valid JSON object following this exact schema:
    {"slides": [{"title": "Slide Title", "content": ["Bullet point 1.", "Bullet point 2."], "speaker_notes": "Notes for the speaker."}]}
    Do not include any other text, explanations, or markdown formatting like ```json before or after the JSON object.
    """
    
    # Combine the user's text and guidance into a single prompt.
    user_content = f"### USER GUIDANCE:\n{guidance}\n\n### SOURCE TEXT:\n{text}"

    # The body of our request, formatted as a Python dictionary
    payload = {
        "model": model_name,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_content}
        ],
        "response_format": {"type": "json_object"}
    }
    
    try:
        # Make the POST request
        response = requests.post(url, headers=headers, json=payload)

        # Check if the request was successful
        if response.status_code == 200:
            # Parse the JSON response from the server
            response_data = response.json()
            # Extract the text content from the response
            response_content = response_data['choices'][0]['message']['content']
            # Convert the JSON string inside the content into a Python dictionary
            parsed_json = json.loads(response_content)
            return parsed_json
        else:
            # If we get an error, display it clearly
            st.error(f"API Error: Status Code {response.status_code}")
            st.error(f"Response Body: {response.text}")
            return None

    except Exception as e:
        st.error(f"An exception occurred while making the API request: {e}")
        return None

# --- UI Components ---
st.title("✨ Your Text, Your Style")
st.subheader("Auto-Generate a Presentation from Text")

st.info("Get your free token at [aipipe.org/login](https://aipipe.org/login).")

# 1. Text Input
with st.expander("Step 1: Provide Your Source Text", expanded=True):
    source_text = st.text_area("Paste text here.", height=300, label_visibility="collapsed")

# 2. Guidance and AI Configuration (REFORGED UI)
with st.expander("Step 2: Configure the AI", expanded=True):
    guidance = st.text_input("Optional: Provide guidance", placeholder="e.g., 'Turn into a 5-slide investor pitch deck'")
    
    aipipe_token = st.text_input("Enter your AI Pipe Token", type="password")

    # Button to trigger fetching models
    if st.button("Load Available Models"):
        st.session_state.models = get_available_models(aipipe_token)
        if not st.session_state.models:
            st.warning("Could not load models. Please check your AI Pipe Token.")

    # Only show the select box if we have a list of models
    if st.session_state.models:
        # Use a selectbox for the user to choose a model
        st.session_state.selected_model = st.selectbox(
            "Choose a model:",
            options=st.session_state.models,
            # Try to pre-select a good default model if it exists
            index=st.session_state.models.index("anthropic/claude-3-5-sonnet-20240620") if "anthropic/claude-3-5-sonnet-20240620" in st.session_state.models else 0
        )

# 3. Template Upload
with st.expander("Step 3: Upload Your Template", expanded=True):
    uploaded_template = st.file_uploader("Upload .pptx or .potx", type=['pptx', 'potx'], label_visibility="collapsed")

# 4. Generate Button
st.divider()
if st.button("🚀 Generate Presentation", type="primary", use_container_width=True):
    # --- Updated Validation ---
    if not source_text:
        st.warning("Please provide the source text in Step 1.")
    elif not aipipe_token:
        st.warning("Please enter your AI Pipe Token in Step 2.")
    elif not st.session_state.selected_model:
        st.warning("Please load and select a model in Step 2.")
    elif not uploaded_template:
        st.warning("Please upload a PowerPoint template in Step 3.")
    else:
        # --- Generation Logic ---
        with st.spinner("🤖 The AI is thinking..."):
            slide_data = generate_slide_content(
                text=source_text,
                guidance=guidance,
                aipipe_token=aipipe_token,
                model_name=st.session_state.selected_model # <-- Use the selected model
            )

        if slide_data:
            st.success("✅ AI has successfully structured the content!")
            st.json(slide_data)