# coding=utf-8
from docx import Document
from util import dump_source_code, traverse_dump, TreeDir, get_docx

import os

if __name__ == '__main__':
    dump_dir_path = os.getcwd()
    doc_obj = get_docx('output.docx')
    doc_obj.add_paragraph(TreeDir(dump_dir_path).str_trees)
    traverse_dump(dump_dir_path, doc_obj)
    doc_obj.save('output.docx')