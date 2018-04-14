import os

DOC_OUTPUT = 'target/output.docx'

def traverse_dir(file_path):
    files = sorted(os.listdir(file_path))
    for fi in files:
        fi_d = os.path.join(file_path, fi)
        if os.path.isdir(fi_d):
            traverse_dir(fi_d)
        else:
            print(fi_d)

def print_file(file_path, output_file=DOC_OUTPUT):
    doc_obj = get_docx(output_file)
    from docx.shared import Pt, Cm
    # print('file_path:: ', file_path)
    doc_obj.add_paragraph('file_path:: ' + file_path)

    with open(file_path, 'r') as file:
        line_num = 0
        for line in file:
            line_num += 1
            line_num_str = str(line_num)
            while len(line_num_str) < 3:
                line_num_str = '0' + line_num_str
            # print(line_num_str, line, end='')
            para = doc_obj.add_paragraph(line_num_str+' '+line.rstrip('\n'))
            para_format = para.paragraph_format
            para_format.space_before = Cm(0)
            para_format.space_after = Cm(0)
            para_format.line_spacing = 12
            # para.line_spacing = Pt(0)
    doc_obj.save(output_file)



def print_file2(file_path, output_file=DOC_OUTPUT):
    doc_obj = get_docx(output_file)
    from docx.shared import Pt, Cm
    # print('file_path:: ', file_path)
    from datetime import datetime
    timestamp = datetime.now().strftime('%Y-%m-%d::')
    doc_obj.add_paragraph(timestamp + 'file_path:: ' + file_path)

    with open(file_path, 'r') as file:
        line_num = 0
        file_text = ''
        for line in file:
            line_num += 1
            line_num_str = str(line_num)
            while len(line_num_str) < 3:
                line_num_str = '0' + line_num_str
            # print(line_num_str, line, end='')
            file_text += line_num_str+' '*4+line
            # para.line_spacing = Pt(0)
        para = doc_obj.add_paragraph(file_text)
        para_format = para.paragraph_format
        para_format.space_before = Cm(0)
        para_format.space_after = Cm(0)
        para_format.line_spacing = Pt(12)

    doc_obj.save(output_file)

def get_docx(file_path):
    from docx import Document
    doc_file = None
    if os.path.exists(file_path):
        doc_file = Document(file_path)
    else:
        doc_file = Document()
    return doc_file

TREE_CHILD = '----'
def tree_dir(file_path, depth=0):
    # 先打印根目录
    if depth == 0:
        print(file_path)
    depth += 1
    files = os.listdir(file_path)
    for fi in files:
        seperator = os.path.sep
        fi_d = os.path.join(file_path, fi)
        file_name = fi_d.split(seperator)[-1]
        if os.path.isdir(fi_d):
            print(TREE_CHILD * depth, file_name)
            tree_dir(fi_d, depth)
        else:
            print(TREE_CHILD*depth, file_name)

if __name__ == '__main__':
    # print_file2('source/extensions/cos_sign/cos_sign.py')
    # print(str(12))
    tree_dir('source/extensions/')

