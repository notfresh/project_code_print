# coding=utf-8
from docx import Document
from util import dump_source_code, traverse_dump, TreeDir, get_docx

import os

if __name__ == '__main__':
    dump_dir_path = os.getcwd()
    output_file = 'code_print.docx'
    doc_obj = get_docx(output_file)
    doc_obj.add_paragraph(TreeDir(dump_dir_path).str_trees)
    traverse_dump(dump_dir_path, doc_obj)
    doc_obj.save(output_file)