import xml.etree.ElementTree as ET
from .utils import NAMESPACES
from .config import get_margin_twips

# A4纸张大小 (twips)
PAGE_WIDTH = '11906'   # 210mm
PAGE_HEIGHT = '16838'  # 297mm


def set_page_size(sectPr):
    """设置页面大小（A4）"""
    w_ns = NAMESPACES['w']
    pgSz = sectPr.find(f'{{{w_ns}}}pgSz')
    if pgSz is None:
        pgSz = ET.Element(f'{{{w_ns}}}pgSz')
        sectPr.append(pgSz)
    
    pgSz.set(f'{{{w_ns}}}w', PAGE_WIDTH)
    pgSz.set(f'{{{w_ns}}}h', PAGE_HEIGHT)


def set_page_margins(sectPr):
    """设置页边距"""
    w_ns = NAMESPACES['w']
    pgMar = sectPr.find(f'{{{w_ns}}}pgMar')
    if pgMar is None:
        pgMar = ET.Element(f'{{{w_ns}}}pgMar')
        sectPr.append(pgMar)
    
    pgMar.set(f'{{{w_ns}}}top', get_margin_twips('top'))
    pgMar.set(f'{{{w_ns}}}bottom', get_margin_twips('bottom'))
    pgMar.set(f'{{{w_ns}}}left', get_margin_twips('left'))
    pgMar.set(f'{{{w_ns}}}right', get_margin_twips('right'))
    pgMar.set(f'{{{w_ns}}}header', get_margin_twips('header'))
    pgMar.set(f'{{{w_ns}}}footer', get_margin_twips('footer'))
    pgMar.set(f'{{{w_ns}}}gutter', get_margin_twips('gutter'))

def apply_page_settings(root):
    """应用页面设置到整个文档"""
    print("正在设置页边距...")
    
    # 1. 处理body最后的sectPr
    body = root.find('w:body', NAMESPACES)
    if body is not None:
        sectPr = body.find('w:sectPr', NAMESPACES)
        if sectPr is not None:
            set_page_size(sectPr)
            set_page_margins(sectPr)
            
    # 2. 处理所有段落中的sectPr（分节符）
    for para in root.findall('.//w:p', NAMESPACES):
        pPr = para.find('w:pPr', NAMESPACES)
        if pPr is not None:
            sectPr = pPr.find('w:sectPr', NAMESPACES)
            if sectPr is not None:
                set_page_size(sectPr)
                set_page_margins(sectPr)
