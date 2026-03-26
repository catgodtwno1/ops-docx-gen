#!/usr/bin/env python3
"""Pack an unpacked .docx directory back into a .docx file.

Usage: python pack.py <input_dir> <output.docx>
"""
import sys, os, zipfile

def pack(src_dir, dst_docx):
    with zipfile.ZipFile(dst_docx, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(src_dir):
            for f in files:
                full = os.path.join(root, f)
                arcname = os.path.relpath(full, src_dir)
                zf.write(full, arcname)
    print(f"Packed {dst_docx}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(f"Usage: {sys.argv[0]} <input_dir> <output.docx>")
        sys.exit(1)
    pack(sys.argv[1], sys.argv[2])
