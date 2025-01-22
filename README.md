# Word Template Filler Application

This application provides a GUI interface for filling Word document templates with user-provided values. It supports placeholders in the format `{{placeholder}}` and handles both text in paragraphs and tables.

## Features

- **Template Selection**: Select a Word document template.
- **Placeholder Identification**: Automatically identify and list placeholders in the template.
- **User Input Collection**: Provide values for each placeholder.
- **Template Filling**: Fill the template with the provided values.
- **Save Filled Document**: Save the filled template as a new Word document.
- **Clear Entries**: Clear all input fields.

## Dependencies

- `tkinter`: For the GUI interface
- `python-docx`: For handling Word documents
- `re`: For regular expression operations

## Installation

To install the required dependencies, run:
```bash
pip install python-docx
```

## Usage

1. **Run the Application**: Execute the script to open the GUI.
    ```bash
    python word_RV1.0.py
    ```
2. **Select Template**: Click on "Abrir o Template" to select a Word document template.
3. **Identify Placeholders**: Click on "Extrair os Dados a Preencher" to extract placeholders from the template.
4. **Fill Placeholders**: Enter values for each placeholder in the provided fields.
5. **Save Filled Template**: Click on "Salvar o Template Preenchido" to save the filled template as a new document.
6. **Clear Entries**: Click on "Limpar os Campos" to clear all input fields.

## Screenshots

![Main Interface](screenshots/main_interface.png)
*Main interface of the Word Template Filler Application.*

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Acknowledgements

- [python-docx](https://python-docx.readthedocs.io/en/latest/) for handling Word documents.
- [tkinter](https://docs.python.org/3/library/tkinter.html) for the GUI interface.