import os
import re

base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
project_dir = os.path.join(base_dir, "src", "project")
output_dir = os.path.join(base_dir, "output")
output_file = os.path.join(output_dir, "IYARE-PROJECT-DRAFT.MD")

os.makedirs(output_dir, exist_ok=True)

# Clean static TOC file
toc_file = os.path.join(project_dir, "00_Front_Matter", "04_Table_of_Contents.md")
os.makedirs(os.path.dirname(toc_file), exist_ok=True)
with open(toc_file, 'w', encoding='utf-8') as f:
    f.write('<div style="text-align: center;">\n\n# TABLE OF CONTENTS\n\n</div>\n')

def compile_project():
    src_dir = 'src/project'
    dir_order = [
        '01_Chapter_1_INTRODUCTION_AND_LITERATURE_REVIEW', 
        '02_Chapter_2_MATERIALS_AND_METHODS', 
        '03_Chapter_3_RESULTS', 
        '04_Chapter_4_DISCUSSION',
        '05_Chapter_5_CONCLUSION'
    ]
    print("Compiling project...")
    
    with open(output_file, 'w', encoding='utf-8') as outfile:
        # Write Chapters
        for d in dir_order:
            d_path = os.path.join(src_dir, d)
            if not os.path.exists(d_path): continue
            for filename in sorted(os.listdir(d_path)):
                if filename.endswith(".md"):
                    filepath = os.path.join(d_path, filename)
                    with open(filepath, 'r', encoding='utf-8-sig') as infile:
                        outfile.write(infile.read() + "\n\n")
                        
        # Write References
        ref_path = os.path.join(src_dir, "99_References.md")
        if os.path.exists(ref_path):
            with open(ref_path, 'r', encoding='utf-8-sig') as infile:
                outfile.write(infile.read() + "\n\n")

        # Write Appendices
        app_dir = os.path.join(src_dir, "06_Appendices")
        if os.path.exists(app_dir):
            for filename in sorted(os.listdir(app_dir)):
                if filename.endswith(".md"):
                    filepath = os.path.join(app_dir, filename)
                    with open(filepath, 'r', encoding='utf-8') as infile:
                        outfile.write(infile.read() + "\n\n")

    print(f"Successfully compiled {output_file}.")

if __name__ == "__main__":
    compile_project()
