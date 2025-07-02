import os
import subprocess
import sys
import re
from pathlib import Path

# SET THIS TO YOUR .bib FILENAME (must be in the folder or use full path)
BIBFILE = "references.bib"  # Change this as needed

def check_pandoc():
    """Check if pandoc is installed and available."""
    # Common pandoc installation paths
    common_paths = [
        'pandoc',  # Standard PATH
        '/usr/local/bin/pandoc',  # macOS Homebrew
        '/opt/homebrew/bin/pandoc',  # macOS Apple Silicon Homebrew  
        'C:\\Program Files\\Pandoc\\pandoc.exe',  # Windows default
        'C:\\Program Files (x86)\\Pandoc\\pandoc.exe',  # Windows x86
        os.path.expanduser('~/.local/bin/pandoc'),  # Linux user install
        os.path.expanduser('~/AppData/Local/Pandoc/pandoc.exe'),  # Windows user
    ]
    
    for pandoc_path in common_paths:
        try:
            result = subprocess.run([pandoc_path, '--version'], 
                                  capture_output=True, text=True, check=True)
            print("Pandoc found:", result.stdout.split('\n')[0])
            print(f"Using pandoc at: {pandoc_path}")
            return pandoc_path
        except (subprocess.CalledProcessError, FileNotFoundError):
            continue
    
    # If not found, provide troubleshooting info
    print("ERROR: Pandoc is not installed or not found in PATH.")
    print("\nTroubleshooting steps:")
    print("1. Verify pandoc is installed:")
    print("   - Windows: Check 'Add or Remove Programs'")
    print("   - macOS: Run 'brew list pandoc' or check Applications")
    print("   - Linux: Run 'which pandoc' or 'dpkg -l | grep pandoc'")
    print("\n2. If installed, try restarting your terminal/command prompt")
    print("\n3. Add pandoc to PATH manually:")
    print("   - Windows: Add pandoc folder to System Environment Variables")
    print("   - macOS/Linux: Add 'export PATH=$PATH:/path/to/pandoc' to your shell config")
    print("\n4. Download from: https://pandoc.org/installing.html")
    
    return None

def find_main_tex(folder_path):
    """Find the main .tex file in the folder."""
    main_candidates = ['main.tex', 'document.tex', 'thesis.tex', 'paper.tex']
    
    # First, look for common main file names
    for candidate in main_candidates:
        main_path = os.path.join(folder_path, candidate)
        if os.path.exists(main_path):
            return candidate
    
    # If no common names found, look for .tex files with \documentclass
    tex_files = [f for f in os.listdir(folder_path) if f.endswith('.tex')]
    
    for tex_file in tex_files:
        tex_path = os.path.join(folder_path, tex_file)
        try:
            with open(tex_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
                if r'\documentclass' in content:
                    return tex_file
        except Exception as e:
            print(f"Warning: Could not read {tex_file}: {e}")
            continue
    
    return None

def create_combined_tex(folder_path, main_tex_file, output_name="combined_document"):
    """Create a combined LaTeX file from main.tex and all included files."""
    
    main_path = os.path.join(folder_path, main_tex_file)
    combined_path = os.path.join(folder_path, f"{output_name}.tex")
    
    print(f"Creating combined LaTeX file from {main_tex_file}...")
    
    def process_tex_recursively(tex_file_path, base_folder, processed_files=None):
        """Recursively process a LaTeX file and its includes."""
        if processed_files is None:
            processed_files = set()
        
        # Avoid infinite recursion
        abs_path = os.path.abspath(tex_file_path)
        if abs_path in processed_files:
            return ""
        processed_files.add(abs_path)
        
        try:
            with open(tex_file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
        except Exception as e:
            print(f"Warning: Could not read {tex_file_path}: {e}")
            return ""
        
        # Find all include/input commands in this file
        patterns = [
            (r'\\input\{([^}]+)\}', 'input'),
            (r'\\include\{([^}]+)\}', 'include'),
            (r'\\subfile\{([^}]+)\}', 'subfile'),
            (r'\\InputIfFileExists\{([^}]+)\}', 'inputifexists'),
        ]
        
        result_content = content
        
        for pattern, cmd_type in patterns:
            matches = list(re.finditer(pattern, content))
            
            # Process matches in reverse order to maintain positions
            for match in reversed(matches):
                filename = match.group(1).strip()
                full_command = match.group(0)
                
                # Handle relative paths properly
                if not filename.endswith('.tex'):
                    filename += '.tex'
                
                # Check if filename contains path separators (like Sections/1-Introduction.tex)
                if '/' in filename or '\\' in filename:
                    # Use the path as specified, relative to base_folder
                    included_path = os.path.join(base_folder, filename.replace('\\', os.sep))
                else:
                    # Look in the same directory as the current file first
                    current_dir = os.path.dirname(tex_file_path)
                    included_path = os.path.join(current_dir, filename)
                    
                    # If not found, look in base folder
                    if not os.path.exists(included_path):
                        included_path = os.path.join(base_folder, filename)
                
                # Normalize the path
                included_path = os.path.normpath(included_path)
                
                if os.path.exists(included_path):
                    print(f"Processing include: {filename} -> {included_path}")
                    
                    # Recursively process the included file
                    included_content = process_tex_recursively(included_path, base_folder, processed_files.copy())
                    
                    # Clean up the included content (remove document structure commands)
                    # Only clean if this is not the main file
                    if os.path.abspath(included_path) != os.path.abspath(main_path):
                        included_content = re.sub(r'\\documentclass.*?\n', '', included_content)
                        included_content = re.sub(r'\\usepackage.*?\n', '', included_content)
                        included_content = re.sub(r'\\begin\{document\}.*?\n', '', included_content)
                        included_content = re.sub(r'\\end\{document\}.*?\n', '', included_content)
                        included_content = re.sub(r'\\maketitle.*?\n', '', included_content)
                        included_content = re.sub(r'\\tableofcontents.*?\n', '', included_content)
                        
                        # Remove leading/trailing whitespace but preserve internal formatting
                        included_content = included_content.strip()
                        if included_content:
                            included_content = '\n% ========== Included from: ' + filename + ' ==========\n' + included_content + '\n% ========== End of: ' + filename + ' ==========\n'
                    
                    # Replace the command with the processed content
                    start_pos = match.start()
                    end_pos = match.end()
                    result_content = (result_content[:start_pos] + 
                                    included_content + 
                                    result_content[end_pos:])
                    
                else:
                    print(f"Warning: Included file not found: {filename}")
                    print(f"  Searched at: {included_path}")
                    
                    # List available files in the expected directory for debugging
                    search_dir = os.path.dirname(included_path)
                    if os.path.exists(search_dir):
                        available_files = [f for f in os.listdir(search_dir) if f.endswith('.tex')]
                        if available_files:
                            print(f"  Available .tex files in {search_dir}: {', '.join(available_files)}")
        
        return result_content
    
    try:
        # Process the main file recursively
        combined_content = process_tex_recursively(main_path, folder_path)
        
        # Write the combined file
        with open(combined_path, 'w', encoding='utf-8') as f:
            f.write(combined_content)
        
        print(f"✓ Combined LaTeX file created: {output_name}.tex")
        return f"{output_name}.tex"
    
    except Exception as e:
        print(f"Error creating combined file: {e}")
        return None

def convert_combined_to_word(folder_path, combined_tex_file, output_name="combined_document"):
    """Convert the combined LaTeX file to Word document."""
    
    pandoc_path = check_pandoc()
    if not pandoc_path:
        return False
    
    # Check for multiple possible bibliography file names
    bib_candidates = [BIBFILE, "References.bib", "references.bib", "bibliography.bib", "IEEEabrv.bib"]
    bib_path = None
    use_bibliography = False
    
    for bib_file in bib_candidates:
        test_path = os.path.join(folder_path, bib_file)
        if os.path.exists(test_path):
            bib_path = test_path
            use_bibliography = True
            print(f"Bibliography file found: {bib_file}")
            break
    
    if not use_bibliography:
        print(f"No bibliography file found. Checked: {', '.join(bib_candidates)}")
    
    tex_path = os.path.join(folder_path, combined_tex_file)
    docx_path = os.path.join(folder_path, f"{output_name}.docx")
    
    print(f"Converting {combined_tex_file} to {output_name}.docx...")
    
    # Build pandoc command
    cmd = [pandoc_path, tex_path, '-o', docx_path]
    
    # Add bibliography if available
    if use_bibliography:
        cmd.extend(['--bibliography', bib_path])
        cmd.extend(['--citeproc'])  # Process citations
    
    # Add additional pandoc options for better Word output
    cmd.extend([
        '--from', 'latex',
        '--to', 'docx',
        '--standalone',
        '--number-sections',  # Number sections automatically
        '--toc',  # Include table of contents
    ])
    
    # Check for reference document
    ref_doc_path = os.path.join(folder_path, 'reference.docx')
    if os.path.exists(ref_doc_path):
        cmd.extend(['--reference-doc', ref_doc_path])
        print("Using reference document for styling: reference.docx")
    
    try:
        # Run pandoc conversion
        result = subprocess.run(
            cmd,
            cwd=folder_path,
            capture_output=True,
            text=True,
            check=True
        )
        
        print(f"✓ Successfully converted: {output_name}.docx")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"✗ Error converting {combined_tex_file}:")
        print(f"  Command: {' '.join(cmd)}")
        print(f"  Error output: {e.stderr}")
        return False
        
    except Exception as e:
        print(f"✗ Unexpected error converting {combined_tex_file}: {str(e)}")
        return False

def convert_individual_files(folder_path):
    """Convert all .tex files individually (original functionality)."""
    
    pandoc_path = check_pandoc()
    if not pandoc_path:
        return
    
    # Check for bibliography files
    bib_candidates = [BIBFILE, "References.bib", "references.bib", "bibliography.bib", "IEEEabrv.bib"]
    bib_path = None
    use_bibliography = False
    
    for bib_file in bib_candidates:
        test_path = os.path.join(folder_path, bib_file)
        if os.path.exists(test_path):
            bib_path = test_path
            use_bibliography = True
            print(f"Bibliography file found: {bib_file}")
            break
    
    if not use_bibliography:
        print(f"No bibliography file found. Converting without bibliography.")
    
    converted_count = 0
    
    for filename in os.listdir(folder_path):
        if filename.endswith('.tex') and not filename.startswith('combined_'):
            tex_path = os.path.join(folder_path, filename)
            output_basename = os.path.splitext(filename)[0]
            docx_path = os.path.join(folder_path, output_basename + '.docx')
            
            print(f"Converting {filename} to {output_basename}.docx...")
            
            # Build pandoc command
            cmd = [pandoc_path, tex_path, '-o', docx_path]
            
            # Add bibliography if available
            if use_bibliography:
                cmd.extend(['--bibliography', bib_path])
                cmd.extend(['--citeproc'])  # Process citations
            
            # Add additional pandoc options for better Word output
            cmd.extend([
                '--from', 'latex',
                '--to', 'docx',
                '--standalone'
            ])
            
            try:
                # Run pandoc conversion
                result = subprocess.run(
                    cmd,
                    cwd=folder_path,
                    capture_output=True,
                    text=True,
                    check=True
                )
                
                print(f"✓ Successfully converted: {output_basename}.docx")
                converted_count += 1
                
            except subprocess.CalledProcessError as e:
                print(f"✗ Error converting {filename}:")
                print(f"  Command: {' '.join(cmd)}")
                print(f"  Error output: {e.stderr}")
                continue
            
            except Exception as e:
                print(f"✗ Unexpected error converting {filename}: {str(e)}")
                continue
    
    if converted_count > 0:
        print(f"\nConversion complete! {converted_count} file(s) converted successfully.")
    else:
        print("\nNo files were converted successfully.")

def main():
    print("Enhanced LaTeX to Word Document Converter")
    print("=" * 50)
    
    # Get folder path from user
    folder = input("Enter the path to the folder containing .tex files: ").strip()
    
    # Remove quotes if present
    folder = folder.strip('"\'')
    
    # Check if folder exists
    if not os.path.exists(folder):
        print(f"ERROR: Folder '{folder}' does not exist.")
        return
    
    if not os.path.isdir(folder):
        print(f"ERROR: '{folder}' is not a directory.")
        return
    
    # Check for .tex files (including in subdirectories)
    tex_files = []
    for root, dirs, files in os.walk(folder):
        for file in files:
            if file.endswith('.tex'):
                rel_path = os.path.relpath(os.path.join(root, file), folder)
                tex_files.append(rel_path)
    
    if not tex_files:
        print(f"ERROR: No .tex files found in '{folder}' or its subdirectories.")
        return
    
    print(f"Found {len(tex_files)} .tex file(s):")
    for tex_file in tex_files:
        print(f"  - {tex_file}")
    
    # Find main.tex or similar
    main_tex = find_main_tex(folder)
    
    if main_tex:
        print(f"\nMain LaTeX file detected: {main_tex}")
        
        # Ask user for conversion preference
        print("\nConversion options:")
        print("1. Create combined document from main.tex and all included files")
        print("2. Convert all .tex files individually")
        print("3. Both options")
        
        choice = input("Enter your choice (1-3): ").strip()
        
        if choice in ['1', '3']:
            print(f"\n--- Creating Combined Document ---")
            output_name = input("Enter output filename (without extension) [combined_document]: ").strip()
            if not output_name:
                output_name = "combined_document"
            
            combined_file = create_combined_tex(folder, main_tex, output_name)
            if combined_file:
                success = convert_combined_to_word(folder, combined_file, output_name)
                if success:
                    print(f"\n✓ Combined document created successfully: {output_name}.docx")
                
                # Clean up combined .tex file
                cleanup = input("Delete temporary combined .tex file? (y/n): ").strip().lower()
                if cleanup == 'y':
                    try:
                        os.remove(os.path.join(folder, combined_file))
                        print(f"✓ Cleaned up: {combined_file}")
                    except Exception as e:
                        print(f"Warning: Could not delete {combined_file}: {e}")
        
        if choice in ['2', '3']:
            print(f"\n--- Converting Individual Files ---")
            convert_individual_files(folder)
    
    else:
        print("\nNo main LaTeX file detected. Converting all files individually...")
        convert_individual_files(folder)

if __name__ == "__main__":
    main()