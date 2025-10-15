import json
from docx import Document as Doc
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from copy import deepcopy
from docx.oxml.ns import qn
from lxml import etree
from xml.dom.minidom import parseString
from docx.shared import Pt
import openpyxl
import os
import sys
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import requests
from urllib.parse import urlparse

import process_num

def iter_block_items(parent):
    """
    Yield each paragraph and table child within *parent*, in document order.
    Each returned value is an instance of either Table or Paragraph. *parent*
    would most commonly be a reference to a main Document object, but
    also works for a _Cell object, which itself can contain paragraphs and tables.
    """
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

import shutil
import os
import re

# 基础数字映射
base_map = {'零': 0, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, 
            '六': 6, '七': 7, '八': 8, '九': 9, '十': 10}

cn_num_map = base_map.copy()
# 生成11-19
for i in range(1, 10):
    cn_num_map[f'十{list(base_map.keys())[i]}'] = 10 + i
# 生成20-99
for tens in range(2, 10):
    tens_char = list(base_map.keys())[tens]
    cn_num_map[f'{tens_char}十'] = tens * 10  # 二十、三十...
    for ones in range(1, 10):
        ones_char = list(base_map.keys())[ones]
        cn_num_map[f'{tens_char}十{ones_char}'] = tens * 10 + ones

def chinese_to_num(chinese):
    if chinese.isdigit():
        return int(chinese)
    
    return cn_num_map.get(chinese, 0)

def split_word(chapters,folder_path):
    original_file = chapters['fulltext']
    doc = Doc(original_file)
    chp_count = 1
    split_path = folder_path.split("/")
    desired_path = "/".join(split_path[:-1])
    doc_file = os.path.join(desired_path,"封面.docx")
    shutil.copyfile(original_file, doc_file)

    new_doc_0 = process_num.WithNumberDocxReader(doc_file, "")
    new_doc = new_doc_0.docx

    chapters["chapter"+str(chp_count)] = doc_file
    styles = doc.styles
    
    block_index = 0
    
    block_before_num = 0

    last_chapter_number = 0  # 记录最近一次拆分的章节编号

    for block in iter_block_items(doc):
        new_chapter_flag = False
        if isinstance(block, Paragraph):
            paragraph = block
            if paragraph and paragraph.text:


                # 拆分条件1：段落对齐格式居中（小标题）
                xml = paragraph._p.xml
                p_style_xml = etree.fromstring(xml)

                # 对齐格式情况 1
                jc_juzhong = None
                for elem in p_style_xml.iter():
                    if elem.tag.endswith("jc"):
                        jc_juzhong = elem.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
                
                style_name = paragraph.style.name


                # 对齐格式情况 2
                p_alignment = None
                for style in styles:
                    if style.name == style_name:
                        p_alignment = style.paragraph_format.alignment
                        # print(f"Alignment: {style.paragraph_format.alignment}",type(style.paragraph_format.alignment))
                
                target_alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  
                
                #居中满足情况之一
                if jc_juzhong == "center" or target_alignment == p_alignment:
                    
                    paragraph_with_number = new_doc_0.get_number_text(paragraph._element.pPr.numPr) + paragraph.text.strip()
                    paragraph_no_space = paragraph_with_number.replace(' ','')
                    print("paragraph_with_number:", paragraph_no_space)

                    pattern_chapter = r"第([一二三四五六七八九十零]{1,6}|[1-9]\d*)(部分|章)"
                    # pattern_chapter = r"第([一二三四五六七八九十]{1,2}|\d+)章"
                    matches_chapter = re.findall(pattern_chapter, paragraph_no_space)

                    # 如果按照章节能找到
                    if len(matches_chapter) > 0:
                        chap_num_str = matches_chapter[0]
                        chap_num = chinese_to_num(chap_num_str[0])

                        if chap_num > last_chapter_number:
                            if len(paragraph_no_space) < 30:
                                print("chapter_paragraph_with_number:", paragraph_no_space)
                                if  '\t' not in paragraph_no_space:
                                    new_chapter_flag = True
                                    last_chapter_number = chap_num

                    # 如果按照部分能找到
                    else:
                        pattern_part = r"第([一二三四五六七八九十]{1,2}|\d+)部分"
                        matches_part = re.findall(pattern_part, paragraph_no_space)
                        if len(matches_part) > 0:
                            chap_num_str = matches_part[0]
                            chap_num = chinese_to_num(chap_num_str)
                            if chap_num > last_chapter_number:
                                if len(paragraph_no_space) < 30 and '\t' not in paragraph_no_space:   
                                    # 标题不会太长避免误判
                                    # 排除包含制表符的段落（可能是目录等）
                                    new_chapter_flag = True
                                    last_chapter_number = chap_num

                if  new_chapter_flag == True:
                    
                    for b in list(iter_block_items(new_doc))[block_index-block_before_num:]:
                        b._element.getparent().remove(b._element)
                        
                    new_doc.save(doc_file)
                    chp_count += 1


                    file_name = paragraph.text + ".docx"
                    doc_file = os.path.join(desired_path, file_name)
                    shutil.copyfile(original_file, doc_file)

                    chapters["chapter"+str(chp_count)] = doc_file
                    new_doc_0 = process_num.WithNumberDocxReader(doc_file, "")
                    new_doc = new_doc_0.docx

                    for b in list(iter_block_items(new_doc))[:block_index]:
                        b._element.getparent().remove(b._element)
                    block_before_num = block_index

        block_index += 1
    
    new_doc.save(doc_file)
    return chapters

if __name__ == "__main__":
    

    fulltext_path = "../word/example.docx"
    
    chapters = split_word({"fulltext":fulltext_path}, fulltext_path)
    
    # 清理章前章后分页符
    for chapter in chapters:
        if chapter != 'fulltext':
            chapter_path = chapters[chapter]
            chapter_doc = Doc(chapter_path)

            chapter_doc.save(chapter_path)