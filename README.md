
# 📝 DOCX Comments Extractor 

This script is especially useful for those working with documents for fine-tuning LLMs (Large Language Models). 

It allows users to extract comments and their associated text from a DOCX file. The output can be presented in multiple formats, and the script provides the option to include additional details such as the author's name and timestamp.

## 🌟 Features

- 📌 Extract comments and the text they are associated with. It handles parent-child relationships of comments too.
- 🎨 Display in original, enhanced, or JSON format.
- 📜 Option to include the author's name and timestamp of the comment.

## 🔧 Requirements

- Python 3.x
- Libraries: `zipfile`, `lxml`, `argparse`, `json`

## 🛠 Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/yourusername/docx-comments-extractor.git
   cd docx-comments-extractor
   ```

2. Ensure you have the required libraries. You can install them using `pip`:

   ```bash
   pip install lxml argparse
   ```

## 🚀 Usage & Sample Outputs

1. **Original Format (default)**:

   ```bash
   python extract_comments.py /path/to/your/docx/file.docx
   ```

   **Sample Output**:
   ```
   Text Segment 1: Hello world
     Comment: Add exclamation.
   --------------------------------------------------
   Text Segment 2: can this
     Comment: Another comment
     Comment: Sure, would be done.
   --------------------------------------------------
   ```

   To include the author's name and timestamp:

   ```bash
   python extract_comments.py /path/to/your/docx/file.docx --include-details
   ```

   **Sample Output**:
   ```
   Text Segment 1: Hello world
     Comment (John Doe at 2023-10-27T17:50:45Z): Add exclamation.
   --------------------------------------------------
   ```

2. **Enhanced Format**:

   ```bash
   python extract_comments.py /path/to/your/docx/file.docx --display-format enhanced
   ```

   **Sample Output**:
   ```
   ==================================================
         Extracted Comments and Associated Text      
   ==================================================

   Text Segment 1:
   ---------------
   "Hello world"
   ---------------
   Comment: Add exclamation.
   --------------------------------------------------

   Text Segment 2:
   ---------------
   "can this"
   ---------------
   Comment: Another comment
   Comment: Sure, would be done.
   --------------------------------------------------
   ```

   To include the author's name and timestamp:

   ```bash
   python extract_comments.py /path/to/your/docx/file.docx --display-format enhanced --include-details
   ```

   **Sample Output**:
   ```
   Text Segment 1:
   ---------------
   "Hello world"
   ---------------
   Comment (John Doe at 2023-10-27T17:50:45Z): Add exclamation.
   --------------------------------------------------
   ```

3. **JSON Format**:

   ```bash
   python extract_comments.py /path/to/your/docx/file.docx --display-format json
   ```

   **Sample Output**:
   ```json
   [
       {
           "associated_text": "Hello world",
           "comments": ["Add exclamation."]
       },
       {
           "associated_text": "can this",
           "comments": ["Another comment", "Sure, would be done."]
       }
   ]
   ```

Replace `/path/to/your/docx/file.docx` with the actual path to your DOCX file.

## 📜 License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.

## 👏 Acknowledgments

- [OpenAI](https://www.openai.com/) for guidance and code snippets.
