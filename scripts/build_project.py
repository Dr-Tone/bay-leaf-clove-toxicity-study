import os
import subprocess
from pathlib import Path

def build():
    src_dir = Path("src/project")
    dist_dir = Path("dist")
    dist_dir.mkdir(exist_ok=True)
    
    combined_md = dist_dir / "combined_project.md"
    output_docx = dist_dir / "final_project.docx"
    
    print("Gathering markdown files...")
    all_files = []
    
    # Traverse directories in alphanumeric order
    for root, dirs, files in os.walk(src_dir):
        dirs.sort()
        for file in sorted(files):
            if file.endswith(".md"):
                all_files.append(Path(root) / file)
    
    # Sort files by their full path to maintain alphanumeric order across folders
    all_files.sort()

    with open(combined_md, "w", encoding="utf-8") as outfile:
        for file_path in all_files:
            print(f"Adding {file_path}")
            with open(file_path, "r", encoding="utf-8") as infile:
                outfile.write(infile.read())
                outfile.write("\n\n")
    
    print(f"Concatenated markdown saved to {combined_md}")
    
    # Call the conversion script
    print("Converting to DOCX...")
    try:
        subprocess.run(["python", "scripts/md_to_docx_pro.py", str(combined_md), str(output_docx)], check=True)
        print(f"Build successful! Final document: {output_docx}")
    except subprocess.CalledProcessError as e:
        print(f"Error during conversion: {e}")

if __name__ == "__main__":
    build()
