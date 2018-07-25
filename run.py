from docx import Document
from util import dump_source_code, traverse_dump, TreeDir, get_docx

if __name__ == '__main__':
    dump_dir_path = '/home/zz/PycharmProjects/api_gateway/app/modules/mobile/event'
    doc_obj = get_docx('target/output.docx')
    doc_obj.add_paragraph(TreeDir(dump_dir_path).str_trees)
    traverse_dump(dump_dir_path, doc_obj)
    doc_obj.save('target/output.docx')