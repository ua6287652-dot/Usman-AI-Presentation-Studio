# Usman AI Presentation Studio

An AI-powered Streamlit app that turns presentation requirements into structured, editable PowerPoint presentations.

## Features
- Presentation topic
- Language
- Number of slides
- Presentation type
- Target audience
- Visual theme
- Additional instructions
- AI-generated slide outline and speaker notes
- Editable `.pptx` download
- No database required for this basic version

## Stack
- Python
- Streamlit
- Groq API
- `openai/gpt-oss-20b`
- python-pptx

## GitHub → Streamlit Cloud
1. Create a GitHub repository.
2. Upload `app.py`, `requirements.txt`, `README.md`, and `.gitignore`.
3. Commit changes.
4. In Streamlit Community Cloud, create a new app from the repository and select `app.py` as the main file.
5. Open **Manage app → Settings → Secrets** and add:

```toml
GROQ_API_KEY = "your_groq_api_key"
```

6. Save and reboot the app.
7. Enter a topic and click **Generate Presentation**.
8. Download the editable `.pptx`.

## Example
Topic: `Artificial Intelligence in Education`
- Language: English
- Slides: 8
- Type: Educational
- Audience: Students
- Theme: Midnight Blue

## Security
Do not hard-code your Groq API key or commit it to GitHub. Store it in Streamlit Secrets.
