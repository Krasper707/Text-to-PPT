# # app.py

# import streamlit as st
# from openai import OpenAI
# import json
# import requests  # <-- NEW: Our tool for making HTTP requests
# from pptx import Presentation  # <-- Add this
# from pptx.util import Inches   # <-- Add this (optional, but good for later)
# import io                      # <-- Add this


# # --- Page Configuration ---
# st.set_page_config(
#     page_title="Your Text, Your Style",
#     page_icon="✨",
#     layout="centered"
# )

# # --- Session State Initialization ---
# # We need to store the list of models so we don't lose it on every interaction.
# if 'models' not in st.session_state:
#     st.session_state.models = []
# if 'selected_model' not in st.session_state:
#     st.session_state.selected_model = None

# # --- LLM and Logic Functions ---

# def analyze_template(template_file) -> tuple[Presentation, dict]:
#     """
#     Analyzes an uploaded PowerPoint template to identify key slide layouts.

#     Args:
#         template_file: The file-like object from Streamlit's file uploader.

#     Returns:
#         A tuple containing the Presentation object and a dictionary mapping
#         layout names ('title_slide', 'content_slide') to their layout objects.
#     """
#     # The python-pptx library needs a file-like object, which Streamlit provides.
#     prs = Presentation(template_file)
    
#     layouts = {}
    
#     # We will try to identify the two most important layouts:
#     # 1. The Title Slide Layout
#     # 2. The Title and Content Layout
    
#     # Heuristic for finding the 'Title and Content' layout (usually has a title and a large body placeholder)
#     for i, layout in enumerate(prs.slide_layouts):
#         has_title = False
#         has_body = False
#         for ph in layout.placeholders:
#             if 'Title' in ph.name or ph.placeholder_format.type == 1: # 1 is Title type
#                 has_title = True
#             if 'Content' in ph.name or 'Body' in ph.name or ph.placeholder_format.type == 2: # 2 is Body type
#                 has_body = True
#         if has_title and has_body:
#             layouts['content_slide'] = layout
#             break # We found one, let's stop here
    
#     # If we couldn't find a content slide, fall back to a common default (layout index 1)
#     if 'content_slide' not in layouts and len(prs.slide_layouts) > 1:
#         layouts['content_slide'] = prs.slide_layouts[1]
        
#     # Heuristic for 'Title Slide' layout (often the first one, index 0)
#     # Or one that has a title and maybe a subtitle, but no main body.
#     if len(prs.slide_layouts) > 0:
#         layouts['title_slide'] = prs.slide_layouts[0]

#     # A final check to ensure we have *something* to work with
#     if 'content_slide' not in layouts:
#         st.warning("Could not definitively identify a 'Title and Content' layout. Using the first available layout.")
#         layouts['content_slide'] = prs.slide_layouts[0]
        
#     return prs, layouts

# @st.cache_data(show_spinner="Fetching available models...")
# def get_available_models(aipipe_token: str) -> list[str]:
#     """
#     Fetches the list of available models from the AI Pipe API.
#     Returns a list of model ID strings.
#     """
#     if not aipipe_token:
#         return []
    
#     url = "https://aipipe.org/openrouter/v1/models"
#     headers = {
#         "Authorization": f"Bearer {aipipe_token}",
#     }
    
#     try:
#         response = requests.get(url, headers=headers)
#         # If the request was successful
#         if response.status_code == 200:
#             data = response.json()
#             # The model data is in the 'data' key, we extract the 'id' from each entry
#             model_ids = [model['id'] for model in data.get('data', [])]
#             return sorted(model_ids) # Return a sorted list
#         else:
#             # If the token is bad or something else goes wrong
#             st.error(f"Failed to fetch models. Status code: {response.status_code} - {response.text}")
#             return []
#     except Exception as e:
#         st.error(f"An error occurred while fetching models: {e}")
#         return []

# # def generate_slide_content(text: str, guidance: str, aipipe_token: str, model_name: str) -> dict | None:
# #     # This function remains the same as before, no changes needed here.
# #     try:
# #         client = OpenAI(api_key=aipipe_token, base_url="https://aipipe.org/openai/v1")
# #     except Exception as e:
# #         st.error(f"Failed to initialize the AI client. Error: {e}")
# #         return None

# #     system_prompt = """
# #     You are an expert presentation creator... [Your previous system prompt here] ...
# #     {"slides": [{"title": "Slide Title", "content": ["Bullet point 1."], "speaker_notes": "Notes."}]}
# #     """
# #     user_content = f"### USER GUIDANCE:\n{guidance}\n\n### SOURCE TEXT:\n{text}"
    
# #     st.info(f"Sending request to model: {model_name}...")
# #     try:
# #         response = client.chat.completions.create(
# #             model=model_name,
# #             messages=[
# #                 {"role": "system", "content": system_prompt},
# #                 {"role": "user", "content": user_content}
# #             ],
# #             response_format={"type": "json_object"}
# #         )
# #         response_content = response.choices[0].message.content
# #         return json.loads(response_content)
# #     except Exception as e:
# #         st.error(f"An error occurred while communicating with the AI model: {e}")
# #         return None

# def generate_slide_content(text: str, guidance: str, aipipe_token: str, model_name: str) -> dict | None:
#     """
#     Uses an LLM via AI Pipe to structure text into slide content.
#     This version builds the HTTP request manually to bypass any SDK issues.
#     """
#     st.info(f"Sending request to model: {model_name} via manual POST request...")

#     # The endpoint for chat completions on the OpenRouter path
#     url = "https://aipipe.org/openrouter/v1/chat/completions"

#     # Define the headers, exactly like in our working get_models function
#     headers = {
#         "Authorization": f"Bearer {aipipe_token}",
#         "Content-Type": "application/json",
#         # (Optional but good practice) Add a Referer header
#         "HTTP-Referer": "https://github.com/mshakir-io/your-text-your-style-app" # Change to your repo
#     }

#     # This is our master instruction to the AI.
#     system_prompt = """
#     You are an expert presentation creator. Your task is to analyze the following text and user guidance to structure it into a series of presentation slides.
#     - Break down the content logically.
#     - The number of slides should be reasonable for the amount of text provided.
#     - Ensure each slide's content is concise and formatted as bullet points.
#     - Create a title for each slide.
#     - Generate brief, helpful speaker notes for each slide.
#     - You MUST output your response ONLY as a single, valid JSON object following this exact schema:
#     {"slides": [{"title": "Slide Title", "content": ["Bullet point 1.", "Bullet point 2."], "speaker_notes": "Notes for the speaker."}]}
#     Do not include any other text, explanations, or markdown formatting like ```json before or after the JSON object.
#     """
    
#     # Combine the user's text and guidance into a single prompt.
#     user_content = f"### USER GUIDANCE:\n{guidance}\n\n### SOURCE TEXT:\n{text}"

#     # The body of our request, formatted as a Python dictionary
#     payload = {
#         "model": model_name,
#         "messages": [
#             {"role": "system", "content": system_prompt},
#             {"role": "user", "content": user_content}
#         ],
#         "response_format": {"type": "json_object"}
#     }
    
#     try:
#         # Make the POST request
#         response = requests.post(url, headers=headers, json=payload)

#         # Check if the request was successful
#         if response.status_code == 200:
#             # Parse the JSON response from the server
#             response_data = response.json()
#             # Extract the text content from the response
#             response_content = response_data['choices'][0]['message']['content']
#             # Convert the JSON string inside the content into a Python dictionary
#             parsed_json = json.loads(response_content)
#             return parsed_json
#         else:
#             # If we get an error, display it clearly
#             st.error(f"API Error: Status Code {response.status_code}")
#             st.error(f"Response Body: {response.text}")
#             return None

#     except Exception as e:
#         st.error(f"An exception occurred while making the API request: {e}")
#         return None


# def create_presentation(slide_data: dict, template_file) -> io.BytesIO | None:
#     """
#     Creates a new PowerPoint presentation. It loads the template for styles,
#     deletes the original slides, and then adds the new AI-generated slides.
#     This version includes robust data sanitization and placeholder finding.
#     """
#     # --- DATA CORRECTION STEP for the main 'slides' list ---
#     slides_list = slide_data.get('slides', [])
#     if slides_list and isinstance(slides_list, list) and len(slides_list) > 0 and isinstance(slides_list[0], dict):
#         if all(isinstance(k, str) and k.isdigit() for k in slides_list[0].keys()):
#             st.warning("AI returned a dictionary with string numbers. Attempting to fix.")
#             corrected_list = [v for k, v in sorted(slides_list[0].items(), key=lambda item: int(item[0]))]
#             slide_data['slides'] = corrected_list
#         elif all(isinstance(k, int) for k in slides_list[0].keys()):
#             st.warning("AI returned a dictionary instead of a list for slides. Attempting to fix.")
#             slide_data['slides'] = list(slides_list[0].values())
            
#     # --- PRESENTATION BUILDING ---
#     prs, layouts = analyze_template(template_file)
#     if not prs or not layouts.get('title_slide') or not layouts.get('content_slide'):
#         return None

#     # Delete original slides from the template to start clean
#     for i in range(len(prs.slides) - 1, -1, -1):
#         rId = prs.slides._sldIdLst[i].rId
#         prs.part.drop_rel(rId)
#         del prs.slides._sldIdLst[i]

#     try:
#         # --- Create Title Slide ---
#         first_slide_info = slide_data['slides'][0]
#         slide = prs.slides.add_slide(layouts['title_slide'])
#         if slide.shapes.title:
#             slide.shapes.title.text = first_slide_info.get('title', 'Presentation Title')
#         for shape in slide.placeholders:
#              if shape.placeholder_format.type != 1:
#                 shape.text = "Generated by 'Your Text, Your Style'"
#                 break
#         if 'speaker_notes' in first_slide_info:
#             slide.notes_slide.notes_text_frame.text = first_slide_info['speaker_notes']

#         # --- Create Content Slides ---
#         for slide_info in slide_data['slides'][1:]:
#             slide = prs.slides.add_slide(layouts['content_slide'])
#             if slide.shapes.title:
#                 slide.shapes.title.text = slide_info.get('title', '')
            
#             # --- ROBUSTLY FIND BODY SHAPE ---
#             # We find the placeholder that is NOT the title. This is a very reliable
#             # way to get the main content area in a 'Title and Content' layout.
#             body_shape = None
#             for shape in slide.placeholders:
#                 if shape.placeholder_format.type != 1: # Not a title
#                     body_shape = shape
#                     break
            
#             if body_shape:
#                 tf = body_shape.text_frame
#                 tf.clear()

#                 # --- DATA SANITIZATION for 'content' ---
#                 # Check if the AI gave us a dictionary instead of a list for bullet points
#                 content_points = slide_info.get('content', [])
#                 if isinstance(content_points, dict):
#                     content_points = list(content_points.values()) # Fix it
                
#                 # Ensure it's a list before we loop
#                 if isinstance(content_points, list):
#                     for point in content_points:
#                         p = tf.add_paragraph()
#                         p.text = str(point) # Ensure it's a string
#                         p.level = 0
            
#             if 'speaker_notes' in slide_info:
#                 slide.notes_slide.notes_text_frame.text = slide_info['speaker_notes']
        
#         powerpoint_stream = io.BytesIO()
#         prs.save(powerpoint_stream)
#         powerpoint_stream.seek(0)
#         return powerpoint_stream
        
#     except Exception as e:
#         st.error(f"An error occurred while building the presentation slides: {e}")
#         return None

# # --- UI Components ---
# st.title("✨ Your Text, Your Style")
# st.subheader("Auto-Generate a Presentation from Text")

# st.info("Get your free token at [aipipe.org/login](https://aipipe.org/login).")

# # 1. Text Input
# with st.expander("Step 1: Provide Your Source Text", expanded=True):
#     source_text = st.text_area("Paste text here.", height=300, label_visibility="collapsed")

# # 2. Guidance and AI Configuration (REFORGED UI)
# with st.expander("Step 2: Configure the AI", expanded=True):
#     guidance = st.text_input("Optional: Provide guidance", placeholder="e.g., 'Turn into a 5-slide investor pitch deck'")
    
#     aipipe_token = st.text_input("Enter your AI Pipe Token", type="password")

#     # Button to trigger fetching models
#     if st.button("Load Available Models"):
#         st.session_state.models = get_available_models(aipipe_token)
#         if not st.session_state.models:
#             st.warning("Could not load models. Please check your AI Pipe Token.")

#     # Only show the select box if we have a list of models
#     if st.session_state.models:
#         # Use a selectbox for the user to choose a model
#         st.session_state.selected_model = st.selectbox(
#             "Choose a model:",
#             options=st.session_state.models,
#             # Try to pre-select a good default model if it exists
#             index=st.session_state.models.index("anthropic/claude-3-5-sonnet-20240620") if "anthropic/claude-3-5-sonnet-20240620" in st.session_state.models else 0
#         )

# # 3. Template Upload
# with st.expander("Step 3: Upload Your Template", expanded=True):
#     uploaded_template = st.file_uploader("Upload .pptx or .potx", type=['pptx', 'potx'], label_visibility="collapsed")

# # 4. Generate Button
# st.divider()
# if st.button("🚀 Generate Presentation", type="primary", use_container_width=True):
#     # --- Input Validation ---
#     if not source_text:
#         st.warning("Please provide source text in Step 1.")
#     elif not aipipe_token:
#         st.warning("Please enter your AI Pipe Token in Step 2.")
#     elif not st.session_state.selected_model:
#         st.warning("Please load and select a model in Step 2.")
#     elif not uploaded_template:
#         st.warning("Please upload a template in Step 3.")
#     else:
#         # --- Processing Pipeline ---
#         with st.spinner("🤖 AI is structuring your content..."):
#             slide_data = generate_slide_content(
#                 source_text, guidance, aipipe_token, st.session_state.selected_model
#             )

#         if slide_data:
#             st.success("✅ AI content generated successfully!")
#             with st.spinner("🎨 Applying your template's style..."):
#                 powerpoint_file = create_presentation(slide_data, uploaded_template)
            
#             if powerpoint_file:
#                 st.success("🎉 Your presentation is ready!")
#                 st.download_button(
#                     label="📥 Download Presentation (.pptx)",
#                     data=powerpoint_file,
#                     file_name="generated_presentation.pptx",
#                     mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
#                     use_container_width=True
#                 )
#             else:
#                 st.error("Failed to build the PowerPoint file from the AI content.")
#         else:
#             st.error("Failed to get a valid structure from the AI.")

# app.py (reloaded)

import streamlit as st
import json
import requests
from pptx import Presentation
import io
from typing import Union # <-- IMPORT THIS FOR OLDER PYTHON VERSIONS
from pptx.enum.text import MSO_AUTO_SIZE


# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Your Text, Your Style", page_icon="✨", layout="centered")

# --- SESSION STATE ---
if 'models' not in st.session_state: st.session_state.models = []
if 'selected_model' not in st.session_state: st.session_state.selected_model = None

# =====================================================================================
# LOGIC FUNCTIONS (WITH COMPATIBLE TYPE HINTS)
# =====================================================================================

@st.cache_data(show_spinner="Fetching available models...")
def get_available_models(aipipe_token: str) -> list[str]:
    # This function remains the same.
    if not aipipe_token: return []
    url = "https://aipipe.org/openrouter/v1/models"
    headers = {"Authorization": f"Bearer {aipipe_token}"}
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return sorted([model['id'] for model in response.json().get('data', [])])
        else:
            st.error(f"Failed to fetch models. Status: {response.status_code} - {response.text}")
            return []
    except Exception as e:
        st.error(f"An error occurred while fetching models: {e}")
        return []

def generate_slide_content(text: str, guidance: str, aipipe_token: str, model_name: str) -> Union[dict, None]: # <-- FIX
    """Uses an LLM to structure text and suggest visuals."""
    st.info(f"Sending request to model: {model_name}...")
    url = "https://aipipe.org/openrouter/v1/chat/completions"
    headers = {"Authorization": f"Bearer {aipipe_token}", "Content-Type": "application/json"}
    
    # system_prompt = """
    # You are an expert presentation creator. Your task is to analyze the following text and user guidance to structure it into a series of presentation slides.
    # - Break down the content logically. The number of slides should be reasonable for the text provided.
    # - For each slide, provide a `title`, `content` as bullet points, and `speaker_notes`.
    # - For each slide, also provide a `visual_suggestion`. This should be a brief, one-sentence description of an ideal image.
    # - If the slide content is purely abstract or does not need an image, the value for `visual_suggestion` MUST be the exact string "none".
    # - You MUST output your response ONLY as a single, valid JSON object following this exact schema:
    # {"slides": [{"title": "Slide Title", "content": ["Bullet point 1."], "speaker_notes": "Notes.", "visual_suggestion": "A picture of a computer."}]}
    # Do not include any other text, explanations, or markdown.
    # """
    system_prompt = """
    You are an expert presentation creator. Your task is to analyze the following text and user guidance to structure it into a series of presentation slides.
    - Break down the content logically.
    - For each slide, provide a `title`, `content`, `speaker_notes`, and a `visual_suggestion`.
    - For the `content` field, create a list of strings. To create nested bullet points, start a string with '- ' for a sub-bullet.
    - For `visual_suggestion`, describe an ideal image. If no image is needed, use the exact string "none".
    - You MUST output your response ONLY as a single, valid JSON object following this exact schema:
    {"slides": [{"title": "Main Point", "content": ["First bullet.", "- Sub-bullet.", "Second bullet."], "speaker_notes": "Notes.", "visual_suggestion": "A relevant image."}]}
    """

    user_content = f"### USER GUIDANCE:\n{guidance}\n\n### SOURCE TEXT:\n{text}"
    payload = { "model": model_name, "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": user_content}], "response_format": {"type": "json_object"} }
    
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=180)
        if response.status_code == 200:
            return json.loads(response.json()['choices'][0]['message']['content'])
        else:
            st.error(f"API Error: Status {response.status_code} - Response: {response.text}")
            return None
    except Exception as e:
        st.error(f"An exception occurred making the API request: {e}")
        return None

# def analyze_template(template_file) -> tuple[Union[Presentation, None], dict, list]: # <-- FIX
#     """
#     Analyzes a template to find layouts and extract an image bank.
#     Identifies 'title', 'content_only', and 'content_with_image' layouts.
#     """
#     try:
#         prs = Presentation(template_file)
#         layouts = {}
#         image_bank = []

#         for layout in prs.slide_layouts:
#             placeholders = {p.placeholder_format.type for p in layout.placeholders}
#             if 1 in placeholders and 2 in placeholders and 18 in placeholders:
#                 layouts['content_with_image'] = layout
#             elif 1 in placeholders and 2 in placeholders and 18 not in placeholders:
#                 layouts['content_only'] = layout
        
#         layouts['title'] = prs.slide_layouts[0]
#         if 'content_only' not in layouts: layouts['content_only'] = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]
#         if 'content_with_image' not in layouts: layouts['content_with_image'] = layouts.get('content_only')

#         for slide in prs.slides:
#             for shape in slide.shapes:
#                 if hasattr(shape, 'image'):
#                     image_bank.append(io.BytesIO(shape.image.blob))

#         st.success(f"Template analyzed: Found {len(image_bank)} reusable images.")
#         return prs, layouts, image_bank
#     except Exception as e:
#         st.error(f"Could not read the PowerPoint file. It may be corrupt. Error: {e}")
#         return None, {}, []

# # def create_presentation(slide_data: dict, template_file) -> Union[io.BytesIO, None]: # <-- FIX
# #     """Builds the presentation using classified layouts and the image bank."""
# #     prs, layouts, image_bank = analyze_template(template_file)
# #     if not prs: return None

# #     for i in range(len(prs.slides) - 1, -1, -1):
# #         rId = prs.slides._sldIdLst[i].rId
# #         prs.part.drop_rel(rId)
# #         del prs.slides._sldIdLst[i]

# #     slides_list = slide_data.get('slides', [])
# #     image_idx = 0
    
# #     try:
# #         if not slides_list:
# #              st.error("AI returned an empty list of slides.")
# #              return None

# #         # Create Title Slide
# #         slide_info = slides_list[0]
# #         slide = prs.slides.add_slide(layouts['title'])
# #         if slide.shapes.title:
# #             slide.shapes.title.text = slide_info.get('title', 'Presentation Title')
        
# #         for shape in slide.placeholders:
# #              if shape.placeholder_format.type != 1:
# #                 shape.text = "Generated by 'Your Text, Your Style'"
# #                 break
        
# #         # Create Content Slides
# #         for slide_info in slides_list[1:]:
# #             visual_suggestion = slide_info.get('visual_suggestion', 'none').lower()
# #             use_image_layout = (visual_suggestion != 'none' and image_bank)

# #             layout = layouts['content_with_image'] if use_image_layout else layouts['content_only']
# #             slide = prs.slides.add_slide(layout)
            
# #             if slide.shapes.title:
# #                 slide.shapes.title.text = slide_info.get('title', '')
            
# #             body_shape = next((s for s in slide.placeholders if s.placeholder_format.type == 2), None)
# #             if body_shape:
# #                 tf = body_shape.text_frame
# #                 tf.clear()
# #                 content = slide_info.get('content', [])
# #                 if isinstance(content, dict): content = list(content.values())
# #                 if isinstance(content, list):
# #                     for point in content:
# #                         p = tf.add_paragraph()
# #                         p.text = str(point)
# #                         p.level = 0
            
# #             if use_image_layout:
# #                 picture_placeholder = next((s for s in slide.placeholders if s.placeholder_format.type == 18), None)
# #                 if picture_placeholder:
# #                     img_stream = image_bank[image_idx % len(image_bank)]
# #                     img_stream.seek(0)
# #                     picture_placeholder.insert_picture(img_stream)
# #                     image_idx += 1
            
# #             if 'speaker_notes' in slide_info:
# #                 slide.notes_slide.notes_text_frame.text = slide_info['speaker_notes']

# #         powerpoint_stream = io.BytesIO()
# #         prs.save(powerpoint_stream)
# #         powerpoint_stream.seek(0)
# #         return powerpoint_stream
# #     except Exception as e:
# #         st.error(f"An error occurred while building slides: {e}")
# #         return None

# def create_presentation(slide_data: dict, template_file) -> Union[io.BytesIO, None]:
#     """
#     Builds the presentation using a robust, "Zen" approach that makes fewer
#     assumptions about the template's structure.
#     """
#     prs, layouts, image_bank = analyze_template(template_file)
#     if not prs: return None

#     # Delete original slides from the template to start fresh
#     for i in range(len(prs.slides) - 1, -1, -1):
#         rId = prs.slides._sldIdLst[i].rId
#         prs.part.drop_rel(rId)
#         del prs.slides._sldIdLst[i]

#     slides_list = slide_data.get('slides', [])
#     image_idx = 0
    
#     try:
#         if not slides_list:
#              st.error("AI returned an empty list of slides.")
#              return None

#         # --- "ZEN" TITLE SLIDE CREATION ---
#         slide_info = slides_list[0]
#         slide = prs.slides.add_slide(layouts['title'])
        
#         # Explicitly find title and subtitle placeholders
#         title_shape = slide.shapes.title if slide.shapes.title else None
#         subtitle_shape = None
#         for shape in slide.placeholders:
#             if shape != title_shape:
#                 subtitle_shape = shape
#                 break
        
#         # Populate them safely
#         if title_shape:
#             title_shape.text = slide_info.get('title', 'Presentation Title')
#         if subtitle_shape:
#             subtitle_shape.text = "Generated by 'PPT-To-Text'"
        
#         # --- "ZEN" CONTENT SLIDE CREATION ---
#         for slide_info in slides_list[1:]:
#             visual_suggestion = slide_info.get('visual_suggestion', 'none').lower()
#             use_image_layout = (visual_suggestion != 'none' and image_bank)
#             layout = layouts['content_with_image'] if use_image_layout else layouts['content_only']
#             slide = prs.slides.add_slide(layout)
            
#             # Populate the title (this is usually reliable)
#             if slide.shapes.title:
#                 slide.shapes.title.text = slide_info.get('title', '')
            
#             # --- ROBUSTLY FIND THE BODY SHAPE ---
#             # Find the main content placeholder by excluding the title. This works on
#             # almost any "Title and Content"-style layout, regardless of the placeholder type ID.
#             body_shape = None
#             title_shape = slide.shapes.title
#             for shape in slide.placeholders:
#                 if shape != title_shape:
#                     body_shape = shape
#                     break # Found it
            
#             # Now, populate the body if we found it

#             if body_shape:
#                 tf = body_shape.text_frame
#                 tf.clear()
#                 # UPGRADE 1: Auto-fit text to the shape
#                 tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                
#                 content = slide_info.get('content', [])
#                 if isinstance(content, dict): content = list(content.values())
                
#                 if isinstance(content, list):
#                     for point in content:
#                         p = tf.add_paragraph()
#                         # UPGRADE 2: Check for and apply nesting
#                         if point.strip().startswith('- '):
#                             p.text = point.strip().lstrip('- ').strip()
#                             p.level = 1  # Indent to the second level
#                         else:
#                             p.text = point
#                             p.level = 0  # Keep as a main bullet point
#             # Populate image if needed
#             if use_image_layout:
#                 picture_placeholder = next((s for s in slide.placeholders if s.placeholder_format.type == 18), None)
#                 if picture_placeholder:
#                     img_stream = image_bank[image_idx % len(image_bank)]
#                     img_stream.seek(0)
#                     picture_placeholder.insert_picture(img_stream)
#                     image_idx += 1
            
#             # Add speaker notes
#             if 'speaker_notes' in slide_info:
#                 slide.notes_slide.notes_text_frame.text = slide_info['speaker_notes']

#         # Save the final presentation
#         powerpoint_stream = io.BytesIO()
#         prs.save(powerpoint_stream)
#         powerpoint_stream.seek(0)
#         return powerpoint_stream
#     except Exception as e:
#         st.error(f"An error occurred while building slides: {e}")
#         return None

def analyze_template(template_file) -> tuple[Union[Presentation, None], dict, list]:
    """
    Analyzes a template to find layouts and extract a *clean* image bank.
    Only extracts images from actual Picture placeholders to avoid background art.
    """
    try:
        prs = Presentation(template_file)
        layouts, image_bank = {}, []
        
        # --- MORE ROBUST LAYOUT CLASSIFICATION ---
        for layout in prs.slide_layouts:
            placeholders = {p.placeholder_format.type for p in layout.placeholders}
            has_title = any(p.placeholder_format.type == 1 for p in layout.placeholders)
            has_body = any(p.placeholder_format.type == 2 for p in layout.placeholders)
            has_picture = any(p.placeholder_format.type == 18 for p in layout.placeholders)
            
            if has_title and has_body and has_picture: layouts['content_with_image'] = layout
            elif has_title and has_body: layouts['content_only'] = layout

        layouts['title'] = prs.slide_layouts[0]
        if 'content_only' not in layouts: layouts['content_only'] = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]
        if 'content_with_image' not in layouts: layouts['content_with_image'] = layouts.get('content_only')

        # --- CRITICAL FIX: Only extract images from Picture Placeholders ---
        for slide in prs.slides:
            for shape in slide.shapes:
                # We check if the shape is a placeholder AND its type is PICTURE (18)
                if shape.is_placeholder and shape.placeholder_format.type == 18 and hasattr(shape, 'image'):
                    image_bank.append(io.BytesIO(shape.image.blob))
        
        if image_bank:
            st.success(f"Template analyzed: Found {len(image_bank)} reusable content images.")
        else:
            st.info("Template analyzed: No reusable content images found. Will create a text-only presentation.")

        return prs, layouts, image_bank
    except Exception as e:
        st.error(f"Could not read the PowerPoint file. It may be corrupt. Error: {e}")
        return None, {}, []


def create_presentation(slide_data: dict, template_file) -> Union[io.BytesIO, None]:
    """
    Builds the presentation with a definitive, multi-step hierarchical method
    for finding the correct content placeholder.
    """
    prs, layouts, image_bank = analyze_template(template_file)
    if not prs: return None

    # Delete original slides to ensure a clean slate
    for i in range(len(prs.slides) - 1, -1, -1):
        rId = prs.slides._sldIdLst[i].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[i]

    slides_list = slide_data.get('slides', [])
    image_idx = 0
    
    try:
        if not slides_list:
             st.error("AI returned an empty list of slides.")
             return None

        # --- CLEAN TITLE SLIDE CREATION ---
        slide_info = slides_list[0]
        slide = prs.slides.add_slide(layouts['title'])
        if slide.shapes.title:
            slide.shapes.title.text = slide_info.get('title', 'Presentation Title')

        # --- DEFINITIVE CONTENT SLIDE CREATION ---
        for slide_info in slides_list[1:]:
            visual_suggestion = slide_info.get('visual_suggestion', 'none').lower()
            use_image_layout = (visual_suggestion != 'none' and image_bank)
            layout = layouts['content_with_image'] if use_image_layout else layouts['content_only']
            slide = prs.slides.add_slide(layout)
            
            if slide.shapes.title:
                slide.shapes.title.text = slide_info.get('title', '')
            
            # --- HIERARCHICAL BODY SHAPE FINDING (v3 - Definitive) ---
            body_shape = None
            
            # 1. Ideal Case: Find the official "Body" placeholder (type 2)
            for shape in slide.placeholders:
                if shape.placeholder_format.type == 2:
                    body_shape = shape
                    break
            
            # 2. Good Case: If no official body, find a generic content placeholder by elimination.
            if not body_shape:
                for shape in slide.placeholders:
                    # Exclude Title (1), Footer (14), Date (15), Slide Number (13)
                    if shape.placeholder_format.type not in [1, 14, 15, 13]:
                        if hasattr(shape, 'text_frame'): # Make sure it can hold text
                            body_shape = shape
                            break
            
            # 3. Last Resort Fallback: If still nothing, find the largest text box that isn't the title.
            if not body_shape:
                potential_shapes = [s for s in slide.placeholders if s.has_text_frame and s != slide.shapes.title]
                if potential_shapes:
                    body_shape = max(potential_shapes, key=lambda s: s.width * s.height)

            # --- POPULATE THE FOUND BODY SHAPE ---
            if body_shape:
                tf = body_shape.text_frame
                tf.clear()
                tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                tf.word_wrap = True
                content = slide_info.get('content', [])
                if isinstance(content, dict): content = list(content.values())
                if isinstance(content, list):
                    for point in content:
                        p = tf.add_paragraph()
                        if point.strip().startswith('- '):
                            p.text = point.strip().lstrip('- ').strip()
                            p.level = 1
                        else:
                            p.text = point
                            p.level = 0
            else:
                st.warning(f"Could not find a suitable content placeholder on slide titled '{slide_info.get('title', '')}'. This slide will be empty.")
            
            # --- IMAGE AND NOTES LOGIC (remains the same) ---
            if use_image_layout:
                pic_placeholder = next((s for s in slide.placeholders if s.placeholder_format.type == 18), None)
                if pic_placeholder:
                    img_stream = image_bank[image_idx % len(image_bank)]
                    img_stream.seek(0)
                    pic_placeholder.insert_picture(img_stream)
                    image_idx += 1
            
            if 'speaker_notes' in slide_info:
                slide.notes_slide.notes_text_frame.text = slide_info['speaker_notes']

        powerpoint_stream = io.BytesIO()
        prs.save(powerpoint_stream)
        powerpoint_stream.seek(0)
        return powerpoint_stream
    except Exception as e:
        st.error(f"An error occurred while building slides: {e}")
        return None


# =====================================================================================
# UI COMPONENTS & WORKFLOW
# =====================================================================================
st.title("✨ Your Text, Your Style")
st.subheader("Auto-Generate a Presentation from Text")

with st.expander("Step 1: Provide Your Source Text", expanded=True):
    source_text = st.text_area("Paste text here.", height=300, label_visibility="collapsed")

with st.expander("Step 2: Configure the AI", expanded=True):
    guidance = st.text_input("Optional guidance", placeholder="e.g., 'Turn into a 5-slide investor pitch deck'")
    aipipe_token = st.text_input("Enter your AI Pipe Token", type="password")
    
    if st.button("Load Available Models"):
        st.session_state.models = get_available_models(aipipe_token)
        if not st.session_state.models:
            st.warning("Could not load models. Check your token.")

    if st.session_state.models:
        st.session_state.selected_model = st.selectbox(
            "Choose a model:",
            options=st.session_state.models,
            index=st.session_state.models.index("anthropic/claude-3-5-sonnet-20240620") if "anthropic/claude-3-5-sonnet-20240620" in st.session_state.models else 0
        )

with st.expander("Step 3: Upload Your Template", expanded=True):
    uploaded_template = st.file_uploader("Upload .pptx or .potx", type=['pptx', 'potx'], label_visibility="collapsed")

st.divider()
if st.button("🚀 Generate Presentation", type="primary", use_container_width=True):
    if not source_text or not aipipe_token or not st.session_state.selected_model or not uploaded_template:
        st.warning("Please complete all steps before generating.")
    else:
        with st.spinner("🤖 AI is structuring content and suggesting visuals..."):
            slide_data = generate_slide_content(source_text, guidance, aipipe_token, st.session_state.selected_model)
        if slide_data:
            st.success("✅ AI content generated!")
            with st.spinner("🎨 Applying styles and reusing images..."):
                powerpoint_file = create_presentation(slide_data, uploaded_template)
            if powerpoint_file:
                st.success("🎉 Your presentation is ready!")
                st.download_button(
                    label="📥 Download Presentation (.pptx)",
                    data=powerpoint_file,
                    file_name="generated_presentation.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )