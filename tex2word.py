import os
import subprocess
import sys

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

def convert_tex_to_word(folder_path):
    """Convert all .tex files in the folder to Word documents."""
    
    pandoc_path = check_pandoc()
    if not pandoc_path:
        return
    
    if not check_pandoc():
        return
    
    # Check if bibliography file exists
    bib_path = os.path.join(folder_path, BIBFILE)
    use_bibliography = os.path.exists(bib_path)
    
    if use_bibliography:
        print(f"Bibliography file found: {BIBFILE}")
    else:
        print(f"No bibliography file found ({BIBFILE}). Converting without bibliography.")
    
    converted_count = 0
    
    for filename in os.listdir(folder_path):
        if filename.endswith('.tex'):
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
    print("LaTeX to Word Document Converter")
    print("=" * 40)
    
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
    
    # Check for .tex files
    tex_files = [f for f in os.listdir(folder) if f.endswith('.tex')]
    if not tex_files:
        print(f"ERROR: No .tex files found in '{folder}'.")
        return
    
    print(f"Found {len(tex_files)} .tex file(s): {', '.join(tex_files)}")
    
    # Perform conversion
    convert_tex_to_word(folder)

if __name__ == "__main__":
    main()