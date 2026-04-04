'''
Advanced OMML (Office Math Markup Language) Generator for Physics Lab Reports.

Provides a robust API to programmatically build complex mathematical formulas
as XML trees that are fully compatible and editable in Microsoft Word.
'''

from lxml import etree
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Namespaces
M_NS = 'http://schemas.microsoft.com/office/2006/math'
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NSMAP = {'m': M_NS, 'w': W_NS}

def _m_tag(tag):
    return f"{{{M_NS}}}{tag}"

def _w_tag(tag):
    return f"{{{W_NS}}}{tag}"

def _create_element(tag, ns_tag_func, attrs=None, text=None):
    elem = etree.Element(ns_tag_func(tag), nsmap=NSMAP)
    if attrs:
        for k, v in attrs.items():
            elem.set(ns_tag_func(k), str(v))
    if text is not None:
        elem.text = text
    return elem

def r(text, italic=False, bold=False, font_size=None):
    '''Creates a math run (m:r) with optional formatting.'''
    mr = _create_element('r', _m_tag)
    
    # Run properties
    rpr = _create_element('rPr', _m_tag)
    if italic:
        # Use 'i' for italic in math context
        rpr.append(_create_element('sty', _m_tag, attrs={'val': 'i'}))
    if bold:
        rpr.append(_create_element('sty', _m_tag, attrs={'val': 'b'}))
    if font_size:
        # sz is in half-points
        rpr.append(_create_element('sz', _m_tag, attrs={'val': str(int(font_size * 2))}))
    
    if len(rpr) > 0:
        mr.append(rpr)
        
    t_elem = _create_element('t', _m_tag, text=text)
    # Preserve space if needed
    if text and (text.startswith(' ') or text.endswith(' ')):
        t_elem.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    mr.append(t_elem)
    return mr

def e_wrap(content):
    '''Wraps content in an 'm:e' (element/argument) container.'''
    e = _create_element('e', _m_tag)
    if isinstance(content, list):
        for item in content:
            e.append(item)
    else:
        e.append(content)
    return e

def frac(num_content, den_content):
    '''Creates a fraction (m:f).'''
    f = _create_element('f', _m_tag)
    # Fraction properties (default bar)
    f.append(_create_element('fPr', _m_tag))
    
    num = _create_element('num', _m_tag)
    num.append(e_wrap(num_content))
    f.append(num)
    
    den = _create_element('den', _m_tag)
    den.append(e_wrap(den_content))
    f.append(den)
    return f

def rad(content, degree=None):
    '''Creates a radical/root (m:rad).'''
    rad_elem = _create_element('rad', _m_tag)
    
    # Radical properties
    rad_pr = _create_element('radPr', _m_tag)
    if degree is None:
        # Hide degree for square root
        rad_pr.append(_create_element('degHide', _m_tag, attrs={'val': '1'}))
    rad_elem.append(rad_pr)
    
    deg = _create_element('deg', _m_tag)
    if degree is not None:
        deg.append(e_wrap(degree))
    rad_elem.append(deg)
    
    e = _create_element('e', _m_tag)
    e.append(e_wrap(content))
    rad_elem.append(e)
    return rad_elem

def ssub(base, sub_content):
    '''Creates a subscript (m:sSub).'''
    ss = _create_element('sSub', _m_tag)
    ss.append(_create_element('sSubPr', _m_tag))
    
    e = _create_element('e', _m_tag)
    e.append(e_wrap(base))
    ss.append(e)
    
    sub = _create_element('sub', _m_tag)
    sub.append(e_wrap(sub_content))
    ss.append(sub)
    return ss

def ssup(base, sup_content):
    '''Creates a superscript (m:sSup).'''
    ss = _create_element('sSup', _m_tag)
    ss.append(_create_element('sSupPr', _m_tag))
    
    e = _create_element('e', _m_tag)
    e.append(e_wrap(base))
    ss.append(e)
    
    sup = _create_element('sup', _m_tag)
    sup.append(e_wrap(sup_content))
    ss.append(sup)
    return ss

def ssubsup(base, sub_content, sup_content):
    '''Creates a combined subscript and superscript (m:sSubSup).'''
    sss = _create_element('sSubSup', _m_tag)
    sss.append(_create_element('sSubSupPr', _m_tag))
    
    e = _create_element('e', _m_tag)
    e.append(e_wrap(base))
    sss.append(e)
    
    sub = _create_element('sub', _m_tag)
    sub.append(e_wrap(sub_content))
    sss.append(sub)
    
    sup = _create_element('sup', _m_tag)
    sup.append(e_wrap(sup_content))
    sss.append(sup)
    return sss

def d(content, beg_chr='(', end_chr=')', sep_chr=''):
    '''Creates a delimiter/grouping (m:d) like parentheses or brackets.'''
    d_elem = _create_element('d', _m_tag)
    d_pr = _create_element('dPr', _m_tag)
    d_pr.append(_create_element('begChr', _m_tag, attrs={'val': beg_chr}))
    d_pr.append(_create_element('endChr', _m_tag, attrs={'val': end_chr}))
    if sep_chr:
        d_pr.append(_create_element('sepChr', _m_tag, attrs={'val': sep_chr}))
    d_elem.append(d_pr)
    
    e = _create_element('e', _m_tag)
    e.append(e_wrap(content))
    d_elem.append(e)
    return d_elem

def nary(char, lower=None, upper=None, content=None):
    '''Creates an n-ary operator (m:nary) like Sum or Integral.'''
    nary_elem = _create_element('nary', _m_tag)
    nary_pr = _create_element('naryPr', _m_tag)
    nary_pr.append(_create_element('chr', _m_tag, attrs={'val': char}))
    # Position limits (undOvr for sum, subSup for integral usually)
    lim_loc = 'undOvr' if char in ['\u03a3', '\u2211'] else 'subSup'
    nary_pr.append(_create_element('limLoc', _m_tag, attrs={'val': lim_loc}))
    nary_elem.append(nary_pr)
    
    sub = _create_element('sub', _m_tag)
    if lower is not None:
        sub.append(e_wrap(lower))
    nary_elem.append(sub)
    
    sup = _create_element('sup', _m_tag)
    if upper is not None:
        sup.append(e_wrap(upper))
    nary_elem.append(sup)
    
    e = _create_element('e', _m_tag)
    if content is not None:
        e.append(e_wrap(content))
    nary_elem.append(e)
    return nary_elem

def omath(*elements):
    '''Root container for a math expression (m:oMath).'''
    om = _create_element('oMath', _m_tag)
    for elem in elements:
        if isinstance(elem, list):
            for sub_elem in elem:
                om.append(sub_elem)
        else:
            om.append(elem)
    return om

def add_formula_para(doc, omath_elem, label=None, align='center'):
    '''Inserts a centered math paragraph into a docx Document, with optional label.'''
    p = doc.add_paragraph()
    p.alignment = {
        'center': WD_ALIGN_PARAGRAPH.CENTER,
        'left': WD_ALIGN_PARAGRAPH.LEFT,
        'right': WD_ALIGN_PARAGRAPH.RIGHT
    }.get(align.lower(), WD_ALIGN_PARAGRAPH.CENTER)
    
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    
    # Append the OMath XML directly to the paragraph's underlying XML
    p._element.append(omath_elem)
    
    if label:
        # Simple tab-based label for now
        run = p.add_run(f'\t({label})')
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11)
    return p

if __name__ == "__main__":
    print("Advanced OMML Generator Loaded.")
