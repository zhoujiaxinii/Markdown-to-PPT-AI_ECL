# -*- coding: utf-8 -*-
"""
PPT生成器 - 智能模板匹配版本 V10
创建新PPT，按顺序复制幻灯片内容
"""

from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
import re
from pypinyin import pinyin, Style
from lxml import etree

class PPTGenerator:
    def __init__(self, template_path, md_content):
        self.template_path = template_path
        self.md_content = md_content
        self.template_prs = Presentation(template_path)
        self._analyze_templates()
    
    def _analyze_templates(self):
        """分析模板"""
        self.templates = {
            'cover': None,
            'toc': None,
            'section': None,
            'end': None,
            'content_1': [],
            'content_2': [],
            'content_3': [],
            'content_4': [],
        }
        
        for idx, slide in enumerate(self.template_prs.slides):
            placeholders = self._find_all_placeholders(slide)
            
            # 结束页
            is_end = False
            for shape in slide.shapes:
                if shape.has_text_frame:
                    text = shape.text_frame.text
                    if '谢' in text or 'xiè' in text.lower():
                        self.templates['end'] = idx
                        is_end = True
                        break
            if is_end:
                continue
            
            # 封面页
            if 'h0_0' in placeholders:
                self.templates['cover'] = idx
                continue
            
            # 章节页
            if 'h1_0' in placeholders and 'h2_0' not in placeholders:
                self.templates['section'] = idx
                continue
            
            # 正文页
            if 'h2_0' in placeholders:
                text_count = len([p for p in placeholders if p.startswith('h3_')])
                key = f'content_{min(text_count, 4)}'
                self.templates[key].append(idx)
                continue
            
            # 目录页
            for shape in slide.shapes:
                if shape.has_text_frame:
                    text = shape.text_frame.text.lower()
                    if '目录' in text or 'mù lù' in text:
                        self.templates['toc'] = idx
                        break
        
        print(f"\n=== 模板分析 ===")
        print(f"封面页: 第{self.templates['cover']+1}页" if self.templates['cover'] is not None else "无")
        print(f"目录页: 第{self.templates['toc']+1}页" if self.templates['toc'] is not None else "无")
        print(f"章节页: 第{self.templates['section']+1}页" if self.templates['section'] is not None else "无")
        cnt = lambda k: len(self.templates[f'content_{k}'])
        print(f"正文模板: 1框{cnt(1)}个, 2框{cnt(2)}个, 3框{cnt(3)}个, 4框{cnt(4)}个")
        print(f"结束页: 第{self.templates['end']+1}页" if self.templates['end'] is not None else "无")
    
    def _find_all_placeholders(self, slide):
        placeholders = {}
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text
                matches = re.findall(r'\{\{(\w+)\}\}', text)
                for match in matches:
                    if match not in placeholders:
                        placeholders[match] = shape
        return placeholders
    
    def _parse_pinyin_chars(self, text):
        pinyin_list = []
        char_list = []
        
        for char in text:
            if '\u4e00' <= char <= '\u9fff':
                py = pinyin(char, style=Style.TONE)[0][0]
                py_clean = re.sub(r'(\d)$', '', py)
                pinyin_list.append(py_clean)
                char_list.append(char)
            elif char.strip():
                pinyin_list.append('')
                char_list.append(char)
        
        return pinyin_list, char_list
    
    def _create_pinyin_table(self, slide, left, top, width, height, text, font_size=24):
        pinyin_list, char_list = self._parse_pinyin_chars(text)
        if not char_list:
            return None
        
        pt_to_emu = lambda pt: int(pt * 914400 / 72)
        
        def calc_total_width(fs):
            return sum(pt_to_emu(fs * 0.6) * (len(py) if py else 1) + pt_to_emu(fs * 0.5) for py in pinyin_list)
        
        total_width = calc_total_width(font_size)
        actual_font_size = font_size
        
        if total_width > width:
            actual_font_size = max(int(font_size * (width / total_width)), 10)
            total_width = calc_total_width(actual_font_size)
        
        letter_width_emu = pt_to_emu(actual_font_size * 0.5)
        col_widths = [pt_to_emu(actual_font_size * 0.6) * (len(py) if py else 1) + letter_width_emu for py in pinyin_list]
        
        table_shape = slide.shapes.add_table(2, len(char_list), left, top, int(total_width), height)
        table = table_shape.table
        
        for col_idx, col_width in enumerate(col_widths):
            table.columns[col_idx].width = int(col_width)
        
        table.rows[0].height = int(height * 0.45)
        table.rows[1].height = int(height * 0.55)
        
        for col_idx, (py, char) in enumerate(zip(pinyin_list, char_list)):
            for row_idx, content in enumerate([py, char]):
                cell = table.cell(row_idx, col_idx)
                cell.text = content
                cell.text_frame.word_wrap = False
                para = cell.text_frame.paragraphs[0]
                para.font.size = Pt(actual_font_size)
                para.font.name = 'Arial' if row_idx == 0 else 'SimSun'
                para.font.color.rgb = RGBColor(0, 0, 0)
                para.alignment = PP_ALIGN.CENTER
                cell.vertical_anchor = MSO_ANCHOR.BOTTOM if row_idx == 0 else MSO_ANCHOR.TOP
        
        self._hide_table_style(table)
        return table_shape
    
    def _hide_table_style(self, table):
        tbl = table._tbl
        tblPr = tbl.tblPr
        if tblPr is None:
            tblPr = etree.SubElement(tbl, '{http://schemas.openxmlformats.org/drawingml/2006/main}tblPr')
        
        for row in table.rows:
            for cell in row.cells:
                tcPr = cell._tc.get_or_add_tcPr()
                for child in list(tcPr):
                    if 'ln' in child.tag or 'solidFill' in child.tag:
                        tcPr.remove(child)
                etree.SubElement(tcPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}noFill')
    
    def _copy_slide(self, source_slide, target_prs):
        """复制幻灯片到目标演示文稿"""
        # 使用空白布局
        blank_layout = target_prs.slide_layouts[6]
        new_slide = target_prs.slides.add_slide(blank_layout)
        
        # 复制所有shapes
        for shape in source_slide.shapes:
            if shape.has_text_frame:
                new_shape = new_slide.shapes.add_textbox(
                    shape.left, shape.top, shape.width, shape.height
                )
                new_shape.text_frame.word_wrap = shape.text_frame.word_wrap
                
                for para_idx, para in enumerate(shape.text_frame.paragraphs):
                    if para_idx == 0:
                        new_para = new_shape.text_frame.paragraphs[0]
                    else:
                        new_para = new_shape.text_frame.add_paragraph()
                    
                    new_para.text = para.text
                    new_para.alignment = para.alignment
                    try:
                        if para.font.size:
                            new_para.font.size = para.font.size
                    except:
                        pass
                    try:
                        if para.font.name:
                            new_para.font.name = para.font.name
                    except:
                        pass
                    try:
                        if para.font.color and para.font.color.rgb:
                            new_para.font.color.rgb = para.font.color.rgb
                    except:
                        pass
        
        return new_slide
    
    def generate(self):
        print("\n=== 开始生成PPT ===")
        
        # 创建新的PPT
        new_prs = Presentation()
        
        # 删除默认的空白幻灯片
        while len(new_prs.slides) > 0:
            rId = new_prs.slides._sldIdLst[0].rId
            new_prs.part.drop_rel(rId)
            del new_prs.slides._sldIdLst[0]
        
        # 构建页面计划
        pages = []
        
        # 正文模板使用指针
        content_pointers = {1: 0, 2: 0, 3: 0, 4: 0}
        
        def get_content_template(text_count):
            key = min(text_count, 4)
            templates = self.templates[f'content_{key}']
            if not templates:
                return None
            idx = templates[content_pointers[key] % len(templates)]
            content_pointers[key] += 1
            return idx
        
        # 封面页
        if self.md_content.get('h0') and self.templates.get('cover') is not None:
            pages.append(('cover', self.templates['cover'], {'title': self.md_content['h0'][0]}))
        
        # 目录页
        if self.md_content.get('h1') and self.templates.get('toc') is not None:
            pages.append(('toc', self.templates['toc'], {'sections': self.md_content['h1']}))
        
        # 章节和正文页
        current_section = None
        
        for h2 in self.md_content['h2']:
            section = h2.get('section', '')
            
            # 章节页
            if section and section != current_section and self.templates.get('section') is not None:
                current_section = section
                pages.append(('section', self.templates['section'], {'title': section}))
            
            # 正文页
            text_count = len(h2.get('content', []))
            idx = get_content_template(text_count)
            if idx is not None:
                pages.append(('content', idx, h2))
        
        # 结束页
        if self.templates.get('end') is not None:
            pages.append(('end', self.templates['end'], {}))
        
        print(f"计划生成 {len(pages)} 页")
        
        # 复制幻灯片并填充内容
        for page_type, template_idx, data in pages:
            # 复制模板幻灯片
            source_slide = self.template_prs.slides[template_idx]
            new_slide = self._copy_slide(source_slide, new_prs)
            
            # 填充内容
            if page_type == 'cover':
                self._fill_placeholder_in_slide(new_slide, 'h0_0', data['title'], 36)
                self._clear_unused_in_slide(new_slide, ['h0_0'])
                print(f"第{len(new_prs.slides)}页: 封面 - {data['title']}")
            
            elif page_type == 'toc':
                for j, section in enumerate(data['sections']):
                    self._fill_placeholder_in_slide(new_slide, f'h1_{j}', section, 20)
                print(f"第{len(new_prs.slides)}页: 目录")
            
            elif page_type == 'section':
                self._fill_placeholder_in_slide(new_slide, 'h1_0', data['title'], 32)
                self._clear_unused_in_slide(new_slide, ['h1_0'])
                print(f"第{len(new_prs.slides)}页: 章节 - {data['title']}")
            
            elif page_type == 'content':
                used = ['h2_0']
                self._fill_placeholder_in_slide(new_slide, 'h2_0', data['title'], 28)
                for j, text in enumerate(data.get('content', [])):
                    self._fill_placeholder_in_slide(new_slide, f'h3_{j}', text, 20)
                    used.append(f'h3_{j}')
                self._clear_unused_in_slide(new_slide, used)
                print(f"第{len(new_prs.slides)}页: 正文 - {data['title']}")
            
            elif page_type == 'end':
                print(f"第{len(new_prs.slides)}页: 结束页")
        
        # 保存
        output_path = 'output.pptx'
        new_prs.save(output_path)
        
        print(f"\n=== PPT生成完成 ===")
        print(f"输出: {output_path}, 共{len(new_prs.slides)}页")
        
        return output_path
    
    def _fill_placeholder_in_slide(self, slide, placeholder_name, content, font_size=24):
        """在幻灯片中填充占位符"""
        placeholders = self._find_all_placeholders(slide)
        if placeholder_name in placeholders:
            shape = placeholders[placeholder_name]
            self._create_pinyin_table(slide, shape.left, shape.top, shape.width, shape.height, content, font_size)
            shape.text_frame.clear()
            shape.left = Emu(0)
            return True
        return False
    
    def _clear_unused_in_slide(self, slide, used):
        """清除幻灯片中未使用的占位符"""
        placeholders = self._find_all_placeholders(slide)
        for name in placeholders:
            if name not in used:
                shape = placeholders[name]
                shape.text_frame.clear()
                shape.left = Emu(0)


if __name__ == '__main__':
    import sys
    from md_parser import MDParser
    
    if len(sys.argv) > 2:
        parser = MDParser(sys.argv[1])
        md_content = parser.parse()
        generator = PPTGenerator(sys.argv[2], md_content)
        generator.generate()
    else:
        print("用法: python ppt_generator.py <md文件> <模板文件>")