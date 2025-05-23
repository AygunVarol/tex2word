import os
import subprocess

# SET THIS TO YOUR .bib FILENAME (must be in the folder or use full path)
BIBFILE = "references.bib"  # Change this as needed
BIBSTYLE = "plain"          # Or another style like "unsrt", "ieeetr", etc.

LATEX_TEMPLATE = r"""
\documentclass{article}
\usepackage[utf8]{inputenc}
\begin{document}
%s

\bibliographystyle{%s}
\bibliography{%s}
\end{document}
""" % ("%s", BIBSTYLE, os.path.splitext(BIBFILE)[0])

def compile_with_bibtex(folder_path):
    for filename in os.listdir(folder_path):
        if filename.endswith('.tex'):
            original_path = os.path.join(folder_path, filename)
            with open(original_path, 'r', encoding='utf-8') as f:
                content = f.read()

            temp_basename = os.path.splitext(filename)[0] + '_wrapped'
            temp_tex_path = os.path.join(folder_path, temp_basename + '.tex')

            with open(temp_tex_path, 'w', encoding='utf-8') as f:
                f.write(LATEX_TEMPLATE % content)

            print(f"Compiling {filename} as {temp_basename}.pdf with bibliography...")

            # Step 1: Run pdflatex (creates .aux)
            subprocess.run(
                ['pdflatex', '-interaction=nonstopmode', temp_tex_path],
                cwd=folder_path, stdout=subprocess.PIPE, stderr=subprocess.PIPE
            )
            # Step 2: Run bibtex (processes bibliography)
            subprocess.run(
                ['bibtex', temp_basename],
                cwd=folder_path, stdout=subprocess.PIPE, stderr=subprocess.PIPE
            )
            # Step 3 & 4: Run pdflatex twice more for cross-references
            subprocess.run(
                ['pdflatex', '-interaction=nonstopmode', temp_tex_path],
                cwd=folder_path, stdout=subprocess.PIPE, stderr=subprocess.PIPE
            )
            subprocess.run(
                ['pdflatex', '-interaction=nonstopmode', temp_tex_path],
                cwd=folder_path, stdout=subprocess.PIPE, stderr=subprocess.PIPE
            )

            print(f"Successfully compiled: {temp_basename}.pdf")

            # Clean up (optional): remove .aux, .log, .bbl, .blg, temp .tex
            for ext in ['.aux', '.log', '.bbl', '.blg', '.tex']:
                try:
                    os.remove(os.path.join(folder_path, temp_basename + ext))
                except FileNotFoundError:
                    pass

if __name__ == "__main__":
    folder = input("Enter the path to the folder containing .tex files: ").strip()
    compile_with_bibtex(folder)
