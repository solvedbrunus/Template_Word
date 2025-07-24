# Copilot Instructions for Word Template Filler

## Project Overview
This is a Word template filling application with three interface implementations:
- `word RV1.0.py`: Desktop GUI using Tkinter (production version)
- `streamlit_app.py`: Web interface using Streamlit (main web version)
- `test/streamlitTest.py`: Enhanced Streamlit version with improved UI (development/testing)

## Core Architecture Patterns

### Placeholder Detection
All versions use identical `extract_placeholders_in_order(doc)` function:
- Regex pattern: `r'{{[^}]+}}'` for placeholder format `{{fieldname}}`
- Searches both paragraphs and table cells
- Maintains document order for consistent user experience
- Returns deduplicated list preserving first occurrence order

### Template Processing
Standard pattern across all versions:
```python
# 1. Extract placeholders
placeholders = extract_placeholders_in_order(doc)

# 2. Collect user input (UI-specific)
data = {placeholder: user_input for placeholder in placeholders}

# 3. Fill template using string replacement
for paragraph in doc.paragraphs:
    for key, value in data.items():
        paragraph.text = paragraph.text.replace(key, value)
```

### File Structure Conventions
- `/data/` - Contains template files (e.g., `template - Mediacao.docx`)
- `/test/` - Development/experimental versions
- Root level - Production versions

## Implementation-Specific Patterns

### Streamlit Applications (`streamlit_app.py`, `test/streamlitTest.py`)
- Always check `docx_imported` before proceeding with Document operations
- Use `BytesIO()` buffer for in-memory document generation
- Implement bilingual support (en/pt) via `get_ui_text(language)` function
- File upload → auto-extract → fill → download workflow

### Tkinter Application (`word RV1.0.py`)
- Manual workflow: select → extract → fill → save
- Uses `ScrollableFrame` class for handling many placeholders
- Implements status bar updates and message boxes for feedback
- Portuguese-only interface with hardcoded strings

## Critical Development Workflows

### Running Applications
```bash
# Streamlit web version
streamlit run streamlit_app.py

# Desktop GUI version  
python "word RV1.0.py"

# Test/development version
streamlit run test/streamlitTest.py
```

### Dependencies
Always ensure `python-docx` is available:
```python
try:
    from docx import Document
    docx_imported = True
except ImportError:
    docx_imported = False
    # Handle gracefully in Streamlit with st.error()
```

## Known Issues to Address
- `test/streamlitTest.py` has undefined `replace_placeholders()` function (line 165)
- Should use `fill_template()` function like other implementations
- Template processing logic should be consistent across all versions

## Extension Guidelines
- Maintain the core `extract_placeholders_in_order()` function signature
- Preserve placeholder order for consistent user experience
- Keep bilingual support pattern for new Streamlit features
- Use `{{placeholder}}` format consistently across templates
