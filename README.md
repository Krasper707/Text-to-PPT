# Your Text, Your Style ✨

A publicly accessible web application that automatically generates a fully-formatted PowerPoint or a web-native Reveal.js presentation from bulk text, using a provided template for styling and imagery.

**Live Demo:** [View the Live App](https://text-to-ppt.streamlit.app/)

---

## How It Works

This tool leverages a Large Language Model (LLM) to intelligently structure raw text into a presentation format. It can then generate two types of outputs: a `.pptx` file that perfectly matches a user's template, or a self-contained `.html` Reveal.js slideshow.

1.  **Input:** The user provides a block of text and an AI Pipe token (a universal key for various LLMs). For PowerPoint generation, a `.pptx` template is required. For Reveal.js, the template is optional and only used to extract images.

2.  **AI Structuring:** The application sends the text to the chosen LLM with a detailed prompt, requesting a structured JSON output. This JSON serves as a blueprint, defining the `title`, `content`, `speaker_notes`, and a `visual_suggestion` for each slide.

3.  **Asset & Style Analysis:**
    *   **For PowerPoint:** The app analyzes the uploaded `.pptx` template to identify master slide layouts (e.g., "Title and Content", "Content with Image").
    *   **For Both:** If a template is provided, it scans for images within content placeholders, creating a bank of reusable visuals.

4.  **Assembly:**
    *   **PowerPoint:** It programmatically builds a new presentation, creating each slide using the appropriate layout from the template. This ensures all fonts, colors, and logos are inherited perfectly.
    *   **Reveal.js:** It weaves an HTML file, converting the AI's blueprint into `<section>` tags and embedding any reused images directly into the file.

5.  **Output:** The final, professionally styled `.pptx` file or a self-contained `.html` slideshow is provided for immediate download.

## Local Setup and Usage

To run this application on your own machine, follow these steps:

1.  **Prerequisites:**
    *   Python 3.8+
    *   `uv` (or `pip` and `venv`) installed.

2.  **Clone the repository:**
    ```bash
    git clone https://github.com/Krasper707/Text-to-PPT.git
    cd Text-to-PPT
    ```

3.  **Set up the environment and install dependencies:**
    ```bash
    # Create a virtual environment
    uv venv

    # Activate the environment
    # On macOS/Linux:
    source .venv/bin/activate
    # On Windows (PowerShell):
    .venv\Scripts\Activate.ps1

    # Install requirements
    uv pip install -r requirements.txt
    ```

4.  **Run the Streamlit app:**
    ```bash
    streamlit run app.py
    ```

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## Technical Write-up

### Text-to-Slide Mapping

The core of the text analysis is delegated to a Large Language Model. The application constructs a detailed system prompt that instructs the AI to act as an expert presentation creator. It specifies a strict JSON schema as the required output format, which includes fields for `title`, `content` (as a list of strings supporting markdown for nesting), `speaker_notes`, and a `visual_suggestion`. By offloading the semantic structuring to the AI and demanding a machine-readable format, the application can handle any form of input text and convert it into a predictable, logical plan for presentation assembly.

### Visual Style and Asset Application

The application achieves perfect style replication by treating the user's uploaded `.pptx` file as a style guide. Using the `python-pptx` library, it first deconstructs the template to identify its master slide layouts, categorizing them into types like "Title", "Content Only", and "Content with Image". It also scans the template's existing slides for images within content placeholders, creating a bank of reusable visuals. When building the new presentation, it *deletes* the template's original slides and then programmatically adds new slides, applying the appropriate pre-categorized layout. This method ensures that every new slide automatically inherits the template's fonts, color schemes, logos, and placeholder positioning, resulting in a seamlessly integrated and professionally styled final product.