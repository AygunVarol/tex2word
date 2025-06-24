# TeX Section to Word Converter
This project provides a Python script to convert LaTeX section fragments (not full documents) into standalone PDF files, wrapping them in a minimal LaTeX template and supporting BibTeX citations.

## Features

- Automatically wraps each section fragment as a full LaTeX document
- Compiles each section as a standalone Word Document
- Processes BibTeX bibliography

## Usage

1. **Put your `.tex` section files and a `.bib` file in a folder**
2. **Edit `tex2word.py` if your `.bib` file has a different name or location.**
3. **Run the script:**

```bash
python tex2word.py
```
4. **Enter the path to the folder containing .tex files**
