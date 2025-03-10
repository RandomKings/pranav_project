import os
import re
import json
import pandas as pd
import win32com.client
from typing import Sequence
from dotenv import load_dotenv
from docx import Document
from langchain_openai import OpenAI
from langchain.output_parsers import PydanticOutputParser
from langchain.prompts import PromptTemplate
from pydantic import BaseModel
from tkinter import filedialog, messagebox

# Load environment variables
load_dotenv("Final\\LLM_CON\\example.env")
token = os.getenv("GITHUB_TOKEN")
endpoint = "https://models.inference.ai.azure.com"
model_name = "gpt-4o"

# Load configuration
with open("Final\\LLM_CON\\config.json", "r") as f:
    config = json.load(f)
column_fields = config["column_fields"]

# Define schema for extracted data
attributes = {field: str for field in column_fields} 
Columns = type("Columns", (BaseModel,), {"__annotations__": attributes})

class Information(BaseModel):
    information: Sequence[Columns]

# Initialize parser
parser = PydanticOutputParser(pydantic_object=Information)

# Create prompt template
prompt = PromptTemplate(
    template="Ekstrak informasi berikut. Anda harus selalu mengembalikan JSON yang valid yang dipagari oleh blok kode markdown. Jangan kembalikan teks tambahan apa pun.:\n{format_instructions}\n{text}\n",
    input_variables=["text"],
    partial_variables={"format_instructions": parser.get_format_instructions()},
)

def clean_text(text):
    """Cleans text by removing extra spaces and newlines."""
    return re.sub(r"\s+", " ", text).strip()

def process_text(text):
    """Processes text using the AI model to extract structured information."""
    if not text.strip():
        print("Error: Received empty or None text input!")
        return pd.DataFrame()

    text = clean_text(text)  
    print("Processed text:", text[:500])  
    
    _input = prompt.format(text=text)
    model = OpenAI(temperature=0, base_url=endpoint, api_key=token, model=model_name)

    try:
        output = model(_input)
        result = parser.parse(output)
        data = result.model_dump()

        extracted_data = []
        for entry in data['information']:
            row = {field: entry.get(field, "").strip() for field in column_fields}

            # Clean numerical-only fields
            for field in ["diagnosa_sekunder", "prosedur_utama", "prosedur_sekunder"]:
                if row[field].replace(".", "").replace(" ", "").replace("(", "").replace(")", "").isdigit():
                    row[field] = ""

            extracted_data.append(row)

        return pd.DataFrame(extracted_data)

    except Exception as e:
        print("Error in processing:", e)
        return pd.DataFrame()

def read_docx(file_path):
    """Reads text from a DOCX file."""
    doc = Document(file_path)
    return "\n".join([para.text.strip() for para in doc.paragraphs if para.text.strip()])

def read_doc(file_path):
    """Reads text from a DOC file using win32com."""
    abs_path = os.path.abspath(file_path)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Run in background
    doc = word.Documents.Open(abs_path)
    text = doc.Content.Text
    doc.Close()
    word.Quit()
    return text

def select_files():
    """Opens a file dialog for users to select DOCX or DOC files."""
    return filedialog.askopenfilenames(title="Select Documents", filetypes=[("Word Documents", "*.docx;*.doc")])

def process_selected_files(selected_files):
    """Processes selected files and extracts structured data, returning DataFrame."""
    df = pd.DataFrame()

    if not selected_files:
        messagebox.showwarning("No Files Selected", "Please select at least one file.")
        return None

    for file_path in selected_files:
        print("Processing file:", file_path)

        if file_path.endswith(".docx"):
            text = read_docx(file_path)
        elif file_path.endswith(".doc"):
            text = read_doc(file_path)
        else:
            print(f"Unsupported file type: {file_path}")
            continue

        extracted_df = process_text(text)
        df = pd.concat([df, extracted_df], ignore_index=True)

    if df.empty:
        return None  # Ensure we don't return empty data

    return df  # Return DataFrame instead of saving


def save_to_excel(df, output_path):
    """Saves the extracted data to an Excel file."""
    if not df.empty:
        df.to_excel(output_path, index=False)
        print(f"Data saved to {output_path}")
        messagebox.showinfo("Success", f"Data saved to {output_path}")
    else:
        print("No data to save.")
        messagebox.showwarning("No Data", "No data was extracted to save.")

# Main execution flow

def extract_llm(selected_files):
    """Processes selected files using LLM and saves the extracted data."""
    if selected_files:
        df = process_selected_files(selected_files)
        if df is not None:
            output_file = filedialog.asksaveasfilename(
                defaultextension=".xlsx", 
                filetypes=[("Excel Files", "*.xlsx")], 
                title="Save Extracted Data", 
                initialfile="extracted_data.xlsx"
            )
            if output_file:
                save_to_excel(df, output_file)
                return output_file
    return None
