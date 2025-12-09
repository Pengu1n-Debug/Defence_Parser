# Handles File Text Representation

from bs4 import BeautifulSoup

def load_extract_texts(html_file):
    with open(html_file, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f, "html.parser")

    # Get only the 'extract' divs
    extracts = soup.find_all(class_="extract")

    # Combine into plain text
    texts = [e.get_text("\n", strip=True) for e in extracts]

    return "\n\n".join(texts)

#Splitter for Categories

import re
def split_into_categories(text):
    # Split on category headers in plain text format (after HTML tags are stripped)
    parts = re.split(r'(Category\s+[IVXLC]+[—-][^\n]+)', text)
    categories = []

    for i in range(1, len(parts), 2):
        header = parts[i].strip()
        body = parts[i+1].strip() if i+1 < len(parts) else ""

        # Extract category number and title from plain text header
        match = re.match(r'Category\s+([IVXLC]+)[—-](.+)', header)
        if match:
            cat_num = match.group(1)
            cat_title = match.group(2).strip()
        else:
            cat_num, cat_title = "Unknown", ""

        categories.append({
            "category": cat_num,
            "title": cat_title,
            "raw_text": body
        })

    return categories

# AI to Structure

import os
from openai import OpenAI
import json
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Initialize OpenAI client with API key from environment variable
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

def convert_category_to_json(category):
    prompt = f"""
Parse this section of the United States Munitions List into hierarchical JSON following this structure:

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

Rules:
- Use "Label" for identifiers like "(a)", "(b)", "(i)", "(ii)" or category numbers
- Use "Description" for the actual text content
- Use "SubStructures" array for nested items (e.g., sub-points under main entries)
- Keep the hierarchy: Category -> Main entries (a,b,c) -> Sub-entries (i,ii,iii)

Category: {category['category']}
Title: {category['title']}

Text:
{category['raw_text']}

Return only valid JSON following the exact structure above.
"""

    response = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt}]
    )

    content = response.choices[0].message.content
    print(f"API Response: {content[:200]}...")  # Debug output
    
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        # If response isn't valid JSON, return a basic structure
        return {
            "Label": f"Category {category['category']}",
            "Description": category['title'],
            "SubStructures": []
        }


#Put Together

# Load only the relevant USML extract content
text = load_extract_texts("HTML/title-22.html")


# Break into categories
categories = split_into_categories(text)
print(f"Found {len(categories)} categories to process")

# Convert each category into structured JSON
structured = [convert_category_to_json(cat) for cat in categories]

# Save to file
with open("usml.json", "w", encoding="utf-8") as f:
    json.dump(structured, f, indent=2)

