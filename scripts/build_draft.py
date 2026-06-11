import os

# Base directory setup
base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
src_dir = os.path.join(base_dir, "src", "project")
output_dir = os.path.join(base_dir, "output", "draft")
output_file = os.path.join(output_dir, "iyare-draft.md")

# Ensure output directory exists
os.makedirs(output_dir, exist_ok=True)

def compile_draft():
    # STRICTLY build only Chapter 1 and Chapter 2
    dir_order = [
        '01_Chapter_1_INTRODUCTION_AND_LITERATURE_REVIEW',
        '02_Chapter_2_MATERIALS_AND_METHODS'
    ]
    ref_file = os.path.join(src_dir, '99_References.md')
    
    print(f"--- Building Draft: Chapters 1, 2, and References ---")
    
    with open(output_file, 'w', encoding='utf-8') as outfile:
        # Process Folders (Chapters 1, 2)
        for d in dir_order:
            d_path = os.path.join(src_dir, d)
            if not os.path.exists(d_path):
                print(f"Warning: Directory {d} not found.")
                continue
            
            print(f"Processing: {d}")
            # Get sorted list of markdown files
            files = sorted([f for f in os.listdir(d_path) if f.endswith(".md")])
            for filename in files:
                filepath = os.path.join(d_path, filename)
                with open(filepath, 'r', encoding='utf-8-sig') as infile:
                    content = infile.read().strip()
                    if content:
                        outfile.write(content + "\n\n") # Natural spacing
        
        # Append References
        if os.path.exists(ref_file):
            print(f"Processing: References")
            with open(ref_file, 'r', encoding='utf-8-sig') as infile:
                ref_content = infile.read().strip()
                if ref_content:
                    outfile.write(ref_content + "\n\n")
        else:
            print(f"Warning: References file not found at {ref_file}")
                    
    print(f"Successfully compiled draft to: {output_file}")

if __name__ == "__main__":
    compile_draft()
