# DSGL Document Processing - Handles .docx, .pdf, .epub formats

import json
import re
import os
from pathlib import Path
from openai import OpenAI
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Initialize OpenAI client with API key from environment variable
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# Document readers for different formats
def load_docx_text(docx_file):
    """Extract text from .docx file"""
    try:
        from docx import Document
        doc = Document(docx_file)
        return "\n".join([paragraph.text for paragraph in doc.paragraphs if paragraph.text.strip()])
    except ImportError:
        raise ImportError("python-docx not installed. Run: pip install python-docx")

def load_pdf_text(pdf_file):
    """Extract text from .pdf file"""
    try:
        import PyPDF2
        with open(pdf_file, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text() + "\n"
        return text
    except ImportError:
        raise ImportError("PyPDF2 not installed. Run: pip install PyPDF2")

def load_epub_text(epub_file):
    """Extract text from .epub file"""
    try:
        import ebooklib
        from ebooklib import epub
        from bs4 import BeautifulSoup
        
        book = epub.read_epub(epub_file)
        text = ""
        
        for item in book.get_items():
            if item.get_type() == ebooklib.ITEM_DOCUMENT:
                soup = BeautifulSoup(item.get_content(), 'html.parser')
                text += soup.get_text() + "\n"
        
        return text
    except ImportError:
        raise ImportError("ebooklib and beautifulsoup4 not installed. Run: pip install ebooklib beautifulsoup4")

def load_document_text(file_path):
    """Load text from document based on file extension"""
    file_path = Path(file_path)
    
    if file_path.suffix.lower() == '.docx':
        return load_docx_text(file_path)
    elif file_path.suffix.lower() == '.pdf':
        return load_pdf_text(file_path)
    elif file_path.suffix.lower() == '.epub':
        return load_epub_text(file_path)
    else:
        raise ValueError(f"Unsupported file format: {file_path.suffix}")

# DSGL-specific text parsing
def split_dsgl_categories(text):
    """
    Split DSGL text into categories based on DSGL formatting patterns.
    Adapt this regex pattern based on actual DSGL document structure.
    """
    # This pattern may need adjustment based on actual DSGL format
    # Common patterns might be: "Category A", "Part I", "Section 1", etc.
    parts = re.split(r'((?:Category|Part|Section)\s+[A-Z0-9]+[:\-\s][^\n]+)', text, flags=re.IGNORECASE)
    categories = []

    for i in range(1, len(parts), 2):
        header = parts[i].strip()
        body = parts[i+1].strip() if i+1 < len(parts) else ""

        # Extract category identifier and title from header
        match = re.match(r'(Category|Part|Section)\s+([A-Z0-9]+)[:\-\s](.+)', header, re.IGNORECASE)
        if match:
            cat_type = match.group(1)
            cat_num = match.group(2)
            cat_title = match.group(3).strip()
        else:
            cat_type, cat_num, cat_title = "Unknown", "Unknown", header

        categories.append({
            "category": f"{cat_type} {cat_num}",
            "title": cat_title,
            "raw_text": body
        })

    return categories

# AI processing with OpenAI

def truncate_text(text, max_chars=5000):
    """Truncate text to stay within token limits while preserving structure"""
    if len(text) <= max_chars:
        return text
    
    # Try to cut at a reasonable point (paragraph break)
    truncated = text[:max_chars]
    last_paragraph = truncated.rfind('\n\n')
    if last_paragraph > max_chars * 0.5:  # If we can cut at least halfway through
        return text[:last_paragraph] + "\n\n[... content truncated for processing ...]"
    else:
        return text[:max_chars] + "\n[... content truncated for processing ...]"

def convert_dsgl_to_json(category):
    """Convert DSGL category to hierarchical JSON structure with optimizations"""
    # Truncate the raw text to avoid token limits
    truncated_text = truncate_text(category['raw_text'])
    
    prompt = f"""
Parse this DSGL section into JSON:

{{
  "Label": "string",
  "Description": "string", 
  "SubStructures": [
    {{
      "Label": "string",
      "Description": "string",
      "SubStructures": [...]
    }}
  ]
}}

Category: {category['category']}
Title: {category['title']}

Text:
{truncated_text}

Return only valid JSON.
"""

    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",  # Using cheaper, faster model
            messages=[{"role": "user", "content": prompt}],
            max_tokens=1500,  # Limit response tokens to save costs
            temperature=0.1   # Lower temperature for consistency
        )

        content = response.choices[0].message.content
        print(f"API Response: {content[:100]}...")
        
        # Try to extract JSON from response
        json_start = content.find('{')
        json_end = content.rfind('}') + 1
        if json_start >= 0 and json_end > json_start:
            try:
                return json.loads(content[json_start:json_end])
            except:
                pass
        
        return {
            "Label": category['category'],
            "Description": category['title'],
            "SubStructures": []
        }
    
    except Exception as e:
        print(f"Error processing {category['category']}: {str(e)[:100]}")
        return {
            "Label": category['category'],
            "Description": category['title'],
            "SubStructures": []
        }

def process_dsgl_document(file_path, output_file="dsgl.json"):
    """Process a DSGL document and convert to structured JSON"""
    print(f"Processing DSGL document: {file_path}")
    
    # Load document text
    try:
        text = load_document_text(file_path)
        print(f"Successfully loaded document with {len(text)} characters")
    except Exception as e:
        print(f"Error loading document: {e}")
        return
    
    # Split into categories
    categories = split_dsgl_categories(text)
    print(f"Found {len(categories)} categories to process")
    
    if not categories:
        print("No categories found. You may need to adjust the parsing pattern in split_dsgl_categories()")
        return
    
    # Convert each category to structured JSON
    structured = []
    for i, cat in enumerate(categories, 1):
        print(f"Processing category {i}/{len(categories)}: {cat['category']}")
        result = convert_dsgl_to_json(cat)
        structured.append(result)
    
    # Save to file
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(structured, f, indent=2)
    
    print(f"DSGL processing complete. Output saved to: {output_file}")

# Example usage
if __name__ == "__main__":
    # Specify your DSGL document path here
    # dsgl_file = "path/to/your/dsgl_document.docx"  # or .pdf or .epub
    # process_dsgl_document(dsgl_file)
    
    print("DSGL Processor ready. Call process_dsgl_document('path/to/file') to process a document.")
    print("Supported formats: .docx, .pdf, .epub")