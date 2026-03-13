# PPT AI Tool

This tool generates PowerPoint using AI and a template.

Noticefication
------------
You can use your own template and setting the placeholder name by your own.
This is just a simple test tool.
You should prepare your own OpenAI API key to set the enviornment.

Requirements
------------
1. Windows
2. Python 3.10+
3. Microsoft PowerPoint installed
4. Python module:
  streamlit
  pywin32
  python-pptx
  pydantic
  openai
  tqdm

Installation
------------
Double click:

install.bat

Run
------------
Double click:

run_ui.bat

The UI will open in your browser.

Default address:
http://localhost:8501

How to set API_KEY
------------
Open PowerShell
setx OPENAI_API_KEY "sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"

You can test whether is ready by:
echo $env:OPENAI_API_KEY

------------

