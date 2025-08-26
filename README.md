# Your Text, Your Style ✨

A publicly accessible web application that automatically generates a fully-formatted PowerPoint presentation from bulk text, using a provided template for styling, layout, and imagery.


**Live Demo:**[[Link]](https://text-to-ppt.streamlit.app/)

Roll number: 24f2001293
Name: Karthik Murali M
---

## How It Works

This tool leverages a Large Language Model (LLM) to intelligently structure raw text into a presentation format, and then uses `python-pptx` to apply the visual styling from a user-provided template.

1.  **Input:** The user provides a block of text, an AI Pipe token (a universal key for various LLMs), and their own `.pptx` template.
2.  **AI Structuring:** The application sends the text to the chosen LLM with a detailed prompt, requesting a structured JSON output. This JSON serves as a blueprint, defining the title, content, speaker notes, and even a visual suggestion for each slide.
3.  **Style Analysis:** The app analyzes the uploaded `.pptx` template to identify master slide layouts (e.g., "Title and Content", "Content with Image") and extracts a bank of reusable images from content placeholders.
4.  **Assembly:** It then programmatically builds a new presentation. Each new slide is created using the appropriate layout from the template, ensuring all fonts, colors, and logos are inherited perfectly. If the AI suggests a visual and images are available, it places a reused image. Text is automatically formatted, wrapped, and resized to fit.
5.  **Output:** The final, professionally styled `.pptx` file is provided for immediate download.

## Local Setup and Usage

To run this application on your own machine, follow these steps:

1.  **Prerequisites:**
    *   Python 3.8+
    *   `uv` (or `pip` and `venv`) installed.

2.  **Clone the repository:**
    ```bash
    git clone https://github.com/your-username/your-repo-name.git
    cd your-repo-name
    ```

3.  **Set up the environment and install dependencies:**
    ```bash
    # Create a virtual environment
    uv venv

    # Activate the environment (macOS/Linux)
    source .venv/bin/activate

    # Install requirements
    uv pip install -r requirements.txt
    ```

4.  **Run the Streamlit app:**
    ```bash
    streamlit run app.py
    ```

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
