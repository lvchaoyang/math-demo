"""
OMML (Office Math Markup Language) 转 LaTeX 转换器
处理 Word 文档中的数学公式

修复重点：
1. 统一命名空间处理：在解析前递归移除所有 XML 命名空间，避免查找失败。
2. 修复上下标/极限解析：确保所有子元素查找逻辑一致。
3. 增强鲁棒性：完善异常处理和边界情况。
"""

import xml.etree.ElementTree as ET
import re
from typing import Optional, List, Dict, Any


class OMML2LaTeXConverter:
    """将 Word OMML 数学公式转换为 LaTeX"""
    
    # OMML 命名空间 URI (仅用于文档说明，实际解析中会移除)
    OMML_NS_URI = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
    
    def __init__(self):
        self.latex_buffer = []
        
    def _strip_namespaces(self, elem: ET.Element):
        """
        递归移除 XML 元素及其子元素的命名空间前缀。
        这是解决解析不稳定性的核心步骤。
        """
        # 处理标签名
        if elem.tag.startswith('{'):
            elem.tag = elem.tag.split('}', 1)[1]
        
        # 处理属性名
        for attr in list(elem.attrib.keys()):
            if attr.startswith('{'):
                new_attr = attr.split('}', 1)[1]
                elem.attrib[new_attr] = elem.attrib.pop(attr)
        
        # 递归处理子元素
        for child in elem:
            self._strip_namespaces(child)
            
        return elem

    def convert(self, omml_xml: str) -> str:
        """
        将 OMML XML 字符串转换为 LaTeX
        
        Args:
            omml_xml: OMML XML 字符串
            
        Returns:
            LaTeX 字符串
        """
        if not omml_xml or not isinstance(omml_xml, str):
            return ""
            
        try:
            # 解析 XML
            root = ET.fromstring(omml_xml)
            
            # 【关键修复】彻底移除命名空间，确保后续 find() 不需要前缀
            self._strip_namespaces(root)
            
            self.latex_buffer = []
            self._process_element(root)
            result = ''.join(self.latex_buffer)
            
            # 如果结果为空但原始内容有内容，可能是纯文本公式
            if not result.strip():
                return "".join(root.itertext())
                
            return result
            
        except ET.ParseError as e:
            return f"[XML Parse Error: {str(e)}]"
        except Exception as e:
            return f"[Formula Conversion Error: {str(e)}]"
    
    def _process_element(self, elem: ET.Element):
        """递归处理 OMML 元素"""
        if elem is None:
            return
            
        tag = elem.tag
        
        handlers = {
            'oMath': self._handle_math,
            'oMathPara': self._handle_math_para,
            'f': self._handle_fraction,
            'num': self._handle_numerator,
            'den': self._handle_denominator,
            'rad': self._handle_radical,
            'deg': self._handle_degree,
            'e': self._handle_base,
            'sup': self._handle_superscript,
            'sub': self._handle_subscript,
            'sSub': self._handle_subscript,
            'sSup': self._handle_superscript,
            'sSubSup': self._handle_sub_superscript,
            'r': self._handle_run,
            't': self._handle_text,
            'limLow': self._handle_limit,
            'limUpp': self._handle_limit_upper,
            'func': self._handle_function,
            'fName': self._handle_func_name,
            'nary': self._handle_nary,
            'eqArr': self._handle_eq_array,
            'd': self._handle_delimiter,
            'acc': self._handle_accent,
            'bar': self._handle_bar,
            'box': self._handle_box,
            'matrix': self._handle_matrix,
            'm': self._handle_matrix_row,
            'eqArrRow': self._handle_eq_arr_row, # 显式处理行
        }
        
        handler = handlers.get(tag)
        if handler:
            handler(elem)
        else:
            # 默认处理：遍历子元素
            for child in elem:
                self._process_element(child)
    
    def _handle_math(self, elem: ET.Element):
        """处理数学公式容器"""
        for child in elem:
            self._process_element(child)
    
    def _handle_math_para(self, elem: ET.Element):
        """处理数学段落"""
        for child in elem:
            self._process_element(child)
    
    def _handle_fraction(self, elem: ET.Element):
        """处理分数 \frac{num}{den}"""
        self.latex_buffer.append('\\frac{')
        
        num = elem.find('num')
        if num is not None:
            for child in num:
                self._process_element(child)
                
        self.latex_buffer.append('}{')
        
        den = elem.find('den')
        if den is not None:
            for child in den:
                self._process_element(child)
                
        self.latex_buffer.append('}')
    
    def _handle_numerator(self, elem: ET.Element):
        """处理分子"""
        for child in elem:
            self._process_element(child)
    
    def _handle_denominator(self, elem: ET.Element):
        """处理分母"""
        for child in elem:
            self._process_element(child)
    
    def _handle_radical(self, elem: ET.Element):
        """处理根式 \\sqrt[n]{expr}"""
        deg = elem.find('deg')
        has_deg = False
        
        if deg is not None:
            # 检查度数是否有实际内容（排除默认的2）
            deg_text = ''.join(deg.itertext()).strip()
            if deg_text and deg_text != '2':
                has_deg = True

        if has_deg:
            self.latex_buffer.append('\\sqrt[')
            for child in deg:
                self._process_element(child)
            self.latex_buffer.append(']{')
        else:
            self.latex_buffer.append('\\sqrt{')

        e = elem.find('e')
        if e is not None:
            for child in e:
                self._process_element(child)
        self.latex_buffer.append('}')
    
    def _handle_degree(self, elem: ET.Element):
        """处理度数"""
        for child in elem:
            self._process_element(child)
    
    def _handle_base(self, elem: ET.Element):
        """处理底数/基数"""
        for child in elem:
            self._process_element(child)
    
    def _handle_superscript(self, elem: ET.Element):
        """处理上标 base^{sup}"""
        e = elem.find('e')
        if e is not None:
            for child in e:
                self._process_element(child)

        self.latex_buffer.append('^{')
        sup = elem.find('sup')
        if sup is not None:
            for child in sup:
                self._process_element(child)
        self.latex_buffer.append('}')
    
    def _handle_subscript(self, elem: ET.Element):
        """处理下标 base_{sub}"""
        # 【修复】不再使用 ns 前缀，因为已经在 convert 中清洗过
        e = elem.find('e')
        if e is not None:
            for child in e:
                self._process_element(child)

        self.latex_buffer.append('_{')
        sub = elem.find('sub')
        if sub is not None:
            for child in sub:
                self._process_element(child)
        self.latex_buffer.append('}')
    
    def _handle_sub_superscript(self, elem: ET.Element):
        """处理上下标组合 base_{sub}^{sup}"""
        e = elem.find('e')
        if e is not None:
            for child in e:
                self._process_element(child)

        # 先处理下标，再处理上标 (顺序通常不影响渲染，但保持习惯)
        sub = elem.find('sub')
        if sub is not None:
            self.latex_buffer.append('_{')
            for child in sub:
                self._process_element(child)
            self.latex_buffer.append('}')

        sup = elem.find('sup')
        if sup is not None:
            self.latex_buffer.append('^{')
            for child in sup:
                self._process_element(child)
            self.latex_buffer.append('}')
    
    def _handle_run(self, elem: ET.Element):
        """处理文本运行"""
        for child in elem:
            self._process_element(child)
    
    def _handle_text(self, elem: ET.Element):
        """处理文本内容"""
        text = elem.text or ''
        if not text:
            return
            
        # 转换数学符号（希腊字母等）
        text = self._convert_math_symbols(text)
        # 转义 LaTeX 特殊字符
        text = self._escape_latex(text)
        self.latex_buffer.append(text)
    
    def _handle_limit(self, elem: ET.Element):
        """处理极限（下标形式）\\lim_{x\\to a}"""
        e = elem.find('e')
        if e is not None:
            for child in e:
                self._process_element(child)

        lim = elem.find('lim')
        if lim is not None:
            self.latex_buffer.append('_{')
            for child in lim:
                self._process_element(child)
            self.latex_buffer.append('}')
    
    def _handle_limit_upper(self, elem: ET.Element):
        """处理上极限"""
        e = elem.find('e')
        if e is not None:
            for child in e:
                self._process_element(child)

        lim = elem.find('lim')
        if lim is not None:
            self.latex_buffer.append('^{')
            for child in lim:
                self._process_element(child)
            self.latex_buffer.append('}')
    
    def _handle_function(self, elem: ET.Element):
        """处理函数 func(arg)"""
        f_name = elem.find('fName')
        if f_name is not None:
            for child in f_name:
                self._process_element(child)
        else:
            self.latex_buffer.append('\\text{func}')

        e = elem.find('e')
        if e is not None:
            self.latex_buffer.append('(')
            for child in e:
                self._process_element(child)
            self.latex_buffer.append(')')
    
    def _handle_func_name(self, elem: ET.Element):
        """处理函数名"""
        text_elem = elem.find('t')
        if text_elem is not None and text_elem.text:
            func_name = text_elem.text.strip()
            known_funcs = [
                'sin', 'cos', 'tan', 'cot', 'sec', 'csc',
                'log', 'ln', 'exp', 'lg', 'lb',
                'lim', 'limsup', 'liminf',
                'max', 'min', 'sup', 'inf',
                'dim', 'ker', 'deg', 'gcd', 'lcm',
                'det', 'tr', 'hom', 'arg',
                'arcsin', 'arccos', 'arctan', 'arccot', 'arcsec', 'arccsc',
                'sinh', 'cosh', 'tanh', 'coth', 'sech', 'csch',
                'arsinh', 'arcosh', 'artanh'
            ]
            if func_name.lower() in known_funcs:
                self.latex_buffer.append(f'\\{func_name.lower()}')
            else:
                self.latex_buffer.append(f'\\text{{{func_name}}}')
        else:
            for child in elem:
                self._process_element(child)
    
    def _handle_nary(self, elem: ET.Element):
        """处理 n 元运算符（求和、积分等）"""
        char_elem = elem.find('chr')
        sub_elem = elem.find('sub')
        sup_elem = elem.find('sup')
        e_elem = elem.find('e')

        op = '\\sum'  # 默认
        if char_elem is not None:
            char = char_elem.get('val', '∑')
            op_map = {
                '∑': '\\sum', '∏': '\\prod', '∐': '\\coprod',
                '∫': '\\int', '∬': '\\iint', '∭': '\\iiint',
                '∮': '\\oint', '∯': '\\oiint', '∰': '\\oiiint',
                '⋃': '\\bigcup', '⋂': '\\bigcap',
                '⋁': '\\bigvee', '⋀': '\\bigwedge',
                '⊕': '\\bigoplus', '⊗': '\\bigotimes',
                '⨁': '\\bigoplus', '⨂': '\\bigotimes',
                '⨄': '\\biguplus', '⨅': '\\bigsqcap', '⨆': '\\bigsqcup',
            }
            op = op_map.get(char, char)

        self.latex_buffer.append(op)

        if sub_elem is not None:
            self.latex_buffer.append('_{')
            for child in sub_elem:
                self._process_element(child)
            self.latex_buffer.append('}')

        if sup_elem is not None:
            self.latex_buffer.append('^{')
            for child in sup_elem:
                self._process_element(child)
            self.latex_buffer.append('}')

        if e_elem is not None:
            for child in e_elem:
                self._process_element(child)
    
    def _handle_eq_array(self, elem: ET.Element):
        """处理方程组/对齐公式"""
        self.latex_buffer.append('\\begin{aligned}')
        
        for child in elem:
            tag = child.tag
            if tag == 'eqArrRow':
                self._handle_eq_arr_row(child)
                
        self.latex_buffer.append('\\end{aligned}')

    def _handle_eq_arr_row(self, elem: ET.Element):
        """处理方程数组的行"""
        first_cell = True
        for cell in elem:
            tag = cell.tag
            if tag == 'e':
                if not first_cell:
                    self.latex_buffer.append(' & ')
                first_cell = False
                for e_child in cell:
                    self._process_element(e_child)
        self.latex_buffer.append(' \\\\\n')
    
    def _handle_delimiter(self, elem: ET.Element):
        """处理括号/分隔符 \\left( ... \\right)"""
        beg_char = elem.find('begChr')
        end_char = elem.find('endChr')

        beg = beg_char.get('val', '(') if beg_char is not None else '('
        end = end_char.get('val', ')') if end_char is not None else ')'

        delim_map = {
            '(': '(', ')': ')',
            '[': '[', ']': ']',
            '{': '\\{', '}': '\\}',
            '|': '|', '‖': '\\|',
            '⌈': '\\lceil', '⌉': '\\rceil',
            '⌊': '\\lfloor', '⌋': '\\rfloor',
            '⟨': '\\langle', '⟩': '\\rangle',
            '': '.', # 空字符用于单侧括号
        }

        beg_latex = delim_map.get(beg, beg)
        end_latex = delim_map.get(end, end)

        grow = elem.find('grow')
        if grow is not None:
            self.latex_buffer.append(f'\\left{beg_latex}')
        else:
            self.latex_buffer.append(beg_latex)

        e = elem.find('e')
        if e is not None:
            for child in e:
                self._process_element(child)

        if grow is not None:
            self.latex_buffer.append(f'\\right{end_latex}')
        else:
            self.latex_buffer.append(end_latex)
    
    def _handle_accent(self, elem: ET.Element):
        """处理重音符号 \\hat{x}"""
        char_elem = elem.find('chr')
        accent = char_elem.get('val', '^') if char_elem is not None else '^'

        accent_map = {
            '^': '\\hat', '~': '\\tilde', '-': '\\bar', '.': '\\dot',
            '..': '\\ddot', '→': '\\vec', '`': '\\grave', "'": '\\acute',
            'ˇ': '\\check', '°': '\\mathring', '̂': '\\hat', '̃': '\\tilde',
            '̄': '\\bar', '̇': '\\dot', '̈': '\\ddot', '⃗': '\\vec',
        }

        accent_cmd = accent_map.get(accent, '\\hat')
        self.latex_buffer.append(f'{accent_cmd}{{')

        e = elem.find('e')
        if e is not None:
            for child in e:
                self._process_element(child)

        self.latex_buffer.append('}')
    
    def _handle_bar(self, elem: ET.Element):
        """处理上划线/下划线"""
        pos = elem.get('pos', 'top')
        cmd = '\\overline' if pos == 'top' else '\\underline'

        self.latex_buffer.append(f'{cmd}{{')
        e = elem.find('e')
        if e is not None:
            for child in e:
                self._process_element(child)
        self.latex_buffer.append('}')
    
    def _handle_box(self, elem: ET.Element):
        """处理方框 \\boxed{}"""
        self.latex_buffer.append('\\boxed{')
        e = elem.find('e')
        if e is not None:
            for child in e:
                self._process_element(child)
        self.latex_buffer.append('}')
    
    def _handle_matrix(self, elem: ET.Element):
        """处理矩阵"""
        matrix_type = 'matrix'  # 默认无括号

        beg_char = elem.find('begChr')
        if beg_char is not None:
            char = beg_char.get('val', '')
            matrix_type_map = {
                '(': 'pmatrix',
                '[': 'bmatrix',
                '{': 'Bmatrix',
                '|': 'vmatrix',
                '‖': 'Vmatrix',
                '': 'matrix',
            }
            matrix_type = matrix_type_map.get(char, 'matrix')

        self.latex_buffer.append(f'\\begin{{{matrix_type}}}')

        for child in elem:
            tag = child.tag
            if tag == 'm':
                self._handle_matrix_row(child)

        self.latex_buffer.append(f'\\end{{{matrix_type}}}')
    
    def _handle_matrix_row(self, elem: ET.Element):
        """处理矩阵行"""
        first = True
        for child in elem:
            tag = child.tag
            if tag == 'e':
                if not first:
                    self.latex_buffer.append(' & ')
                first = False
                for sub_child in child:
                    self._process_element(sub_child)
        self.latex_buffer.append(' \\\\\n')
    
    def _convert_math_symbols(self, text: str) -> str:
        """转换数学符号为 LaTeX 命令"""
        # 合并所有映射
        all_symbols = {
            # 大写希腊字母
            'Α': '\\Alpha', 'Β': '\\Beta', 'Γ': '\\Gamma', 'Δ': '\\Delta',
            'Ε': '\\Epsilon', 'Ζ': '\\Zeta', 'Η': '\\Eta', 'Θ': '\\Theta',
            'Ι': '\\Iota', 'Κ': '\\Kappa', 'Λ': '\\Lambda', 'Μ': '\\Mu',
            'Ν': '\\Nu', 'Ξ': '\\Xi', 'Ο': '\\Omicron', 'Π': '\\Pi',
            'Ρ': '\\Rho', 'Σ': '\\Sigma', 'Τ': '\\Tau', 'Υ': '\\Upsilon',
            'Φ': '\\Phi', 'Χ': '\\Chi', 'Ψ': '\\Psi', 'Ω': '\\Omega',
            # 小写希腊字母
            'α': '\\alpha', 'β': '\\beta', 'γ': '\\gamma', 'δ': '\\delta',
            'ε': '\\epsilon', 'ζ': '\\zeta', 'η': '\\eta', 'θ': '\\theta',
            'ι': '\\iota', 'κ': '\\kappa', 'λ': '\\lambda', 'μ': '\\mu',
            'ν': '\\nu', 'ξ': '\\xi', 'ο': '\\omicron', 'π': '\\pi',
            'ρ': '\\rho', 'σ': '\\sigma', 'τ': '\\tau', 'υ': '\\upsilon',
            'φ': '\\phi', 'χ': '\\chi', 'ψ': '\\psi', 'ω': '\\omega',
            # 运算符
            '±': '\\pm', '∓': '\\mp', '×': '\\times', '÷': '\\div',
            '∗': '\\ast', '⋆': '\\star', '∘': '\\circ', '•': '\\bullet',
            '∩': '\\cap', '∪': '\\cup', '⊔': '\\sqcup', '⊓': '\\sqcap',
            '∧': '\\wedge', '∨': '\\vee', '⊕': '\\oplus', '⊗': '\\otimes',
            '⊖': '\\ominus', '⊙': '\\odot', '†': '\\dagger', '‡': '\\ddagger',
            '∞': '\\infty', '∂': '\\partial', '∇': '\\nabla',
            '√': '\\surd', '∠': '\\angle', '⊥': '\\perp',
            # 关系符
            '≤': '\\leq', '≥': '\\geq', '≡': '\\equiv', '≈': '\\approx',
            '≠': '\\neq', '∼': '\\sim', '≃': '\\simeq', '≅': '\\cong',
            '≪': '\\ll', '≫': '\\gg', '∝': '\\propto', '⊂': '\\subset',
            '⊃': '\\supset', '⊆': '\\subseteq', '⊇': '\\supseteq', '∈': '\\in',
            '∋': '\\ni', '∉': '\\notin',
            # 箭头
            '←': '\\leftarrow', '→': '\\rightarrow', '↑': '\\uparrow',
            '↓': '\\downarrow', '↔': '\\leftrightarrow', '⇒': '\\Rightarrow',
            '⇐': '\\Leftarrow', '⇔': '\\Leftrightarrow', '↦': '\\mapsto',
            # 其他
            '…': '\\dots', '⋯': '\\cdots', '⋮': '\\vdots', '⋱': '\\ddots',
            'ℵ': '\\aleph', 'ℑ': '\\Im', 'ℜ': '\\Re', '℘': '\\wp',
            '∀': '\\forall', '∃': '\\exists', '∅': '\\emptyset',
        }

        result = text
        # 按长度降序排序，避免短符号覆盖长符号（虽然这里都是单字符，但保持习惯）
        for char in sorted(all_symbols.keys(), key=len, reverse=True):
            result = result.replace(char, all_symbols[char])
        return result

    def _escape_latex(self, text: str) -> str:
        """转义 LaTeX 特殊字符"""
        special_chars = {
            '\\': '\\textbackslash{}',
            '{': '\\{',
            '}': '\\}',
            '$': '\\$',
            '&': '\\&',
            '#': '\\#',
            '^': '\\^{}',
            '_': '\\_',
            '~': '\\textasciitilde{}',
            '%': '\\%',
        }
        
        result = text
        for char, replacement in special_chars.items():
            result = result.replace(char, replacement)
        return result


def convert_omml_to_latex(omml_xml: str) -> str:
    """
    便捷函数：将 OMML XML 转换为 LaTeX
    
    Args:
        omml_xml: OMML XML 字符串
        
    Returns:
        LaTeX 字符串
    """
    converter = OMML2LaTeXConverter()
    return converter.convert(omml_xml)