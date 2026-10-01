# -*- coding: utf-8 -*-
"""
Word 原生 OMML 數學方程式建構器 (math_omml_builder.py)
零外部依賴純 Python 產生 Office Math Markup Language (<m:oMath>)
在 Word 中完全原生可編輯：
- 分式 (m_frac)
- 根式 (m_sqrt)
- 上下標 (m_sup, m_sub, m_subsup)
- 幾何線段頂標 (m_bar)
- 向量頂標 (m_acc)
- 聯立方程組 (m_cases)
- 自動縮放括號與絕對值 (m_delim)
- 矩陣與行列式 (m_matrix)
"""

from docx.oxml import parse_xml

MATH_NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

def m_run(text, italic=True):
    """
    建立 OMML 數學文字節點 <m:r>
    italic=True: 數學變數 (斜體)
    italic=False: 數字、運算子、中文 (正體)
    """
    sty = '' if italic else '<m:rPr><m:sty m:val="p"/></m:rPr>'
    esc = str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return f'<m:r>{sty}<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math" w:eastAsia="標楷體"/><w:sz w:val="22"/></w:rPr><m:t>{esc}</m:t></m:r>'

def m_frac(num_xml, den_xml):
    """建立原生分數 <m:f>：num / den"""
    return f'<m:f><m:num>{num_xml}</m:num><m:den>{den_xml}</m:den></m:f>'

def m_sqrt(base_xml, deg_xml=None):
    """建立原生根號 <m:rad>：平方根或 n 次方根"""
    if deg_xml is None:
        return f'<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>{base_xml}</m:e></m:rad>'
    return f'<m:rad><m:radPr><m:degHide m:val="0"/></m:radPr><m:deg>{deg_xml}</m:deg><m:e>{base_xml}</m:e></m:rad>'

def m_sup(base_xml, sup_xml):
    """建立原生上標/次方 <m:sSup>：base^{sup}"""
    return f'<m:sSup><m:e>{base_xml}</m:e><m:sup>{sup_xml}</m:sup></m:sSup>'

def m_sub(base_xml, sub_xml):
    """建立原生下標 <m:sSub>：base_{sub}"""
    return f'<m:sSub><m:e>{base_xml}</m:e><m:sub>{sub_xml}</m:sub></m:sSub>'

def m_subsup(base_xml, sub_xml, sup_xml):
    """建立原生上下標 <m:sSubSup>：如 C_k^n 或排列組合"""
    return f'<m:sSubSup><m:e>{base_xml}</m:e><m:sub>{sub_xml}</m:sub><m:sup>{sup_xml}</m:sup></m:sSubSup>'

def m_bar(base_xml, pos="top"):
    """建立幾何線段頂線 <m:bar>：如線段 AB"""
    return f'<m:bar><m:barPr><m:pos m:val="{pos}"/></m:barPr><m:e>{base_xml}</m:e></m:bar>'

def m_acc(base_xml, chr_val="&#x20D7;"):
    """建立頂部裝飾符號 <m:acc>：預設向量箭頭 &#x20D7;，圓弧可用 &#x2322;"""
    return f'<m:acc><m:accPr><m:chr m:val="{chr_val}"/></m:accPr><m:e>{base_xml}</m:e></m:acc>'

def m_cases(*eq_xml_list):
    """建立聯立方程組左大括號 <m:d> + <m:eqArr>"""
    rows = "".join([f'<m:e>{eq}</m:e>' for eq in eq_xml_list])
    return f'<m:d><m:dPr><m:begChr m:val="{{"/><m:endChr m:val=""/></m:dPr><m:e><m:eqArr>{rows}</m:eqArr></m:e></m:d>'

def m_delim(inner_xml, left="(", right=")"):
    """建立自動縮放括號或絕對值 <m:d>：如 ( ... ) 或 | ... |"""
    return f'<m:d><m:dPr><m:begChr m:val="{left}"/><m:endChr m:val="{right}"/></m:dPr><m:e>{inner_xml}</m:e></m:d>'

def m_matrix(grid_xml_rows):
    """
    建立矩陣 <m:m>
    grid_xml_rows: [[cell_xml1, cell_xml2], [cell_xml3, cell_xml4]]
    """
    rows_xml = ""
    for r in grid_xml_rows:
        cells_xml = "".join([f'<m:e>{c}</m:e>' for c in r])
        rows_xml += f'<m:mr>{cells_xml}</m:mr>'
    inner = f'<m:m>{rows_xml}</m:m>'
    return m_delim(inner, left="[", right="]")

def append_omath(paragraph, inner_omml_xml):
    """將組裝好的 OMML 字串轉為原生 <m:oMath> 物件並附加至 Word 段落中"""
    omath_str = f'<m:oMath {MATH_NS}>{inner_omml_xml}</m:oMath>'
    paragraph._p.append(parse_xml(omath_str))
