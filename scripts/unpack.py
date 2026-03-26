#!/usr/bin/env python3
"""Unpack a .docx file into a directory for raw XML inspection.

Usage: python unpack.py <input.docx> <output_dir>
"""
import sys, os, zipfile

def unpack(src_docx, dst_dir):
    os.makedirs(dst_dir, exist_ok=True)
    with zipfile.ZipFile(src_docx, 'r') as zf:
        zf.extractall(dst_dir)
    print(f"Unpacked to {dst_dir}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(f"Usage: {sys.argv[0]} <input.docx> <output_dir>")
        sys.exit(1)
    unpack(sys.argv[1], sys.argv[2])
