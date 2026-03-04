# -*- coding: utf-8 -*-
"""
PPT生成器 - 表格对齐版本
使用表格实现拼音-汉字精确对齐
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml.ns import nsdecls
from pptx.oxml import parse_xml
import re
from pypinyin import pinyin, Style

class PPTGenerator:
    def __init__(self, template_path, md_content):
        self.template_path = template_path
        self.md_content = md_content
        self.presentation = Presentation(template_path)
    
    def _parse_pinyin_chars(self, text):
        """解析文本，返回拼音和字符列表"""
        pinyin_list = []
        char_list = []
        
        for char in text:
            if '\u4e00' <= char <= '\u9fff':  # 汉字
                py = pinyin(char, style=Style.TONE)[0][0]
                # 去掉声调数字
                py_clean = re.sub(r'(\d)$', '', py)
                pinyin_list.append(py_clean)
                char_list.append(char)
            elif char.strip():  # 非空白字符
                pinyin_list.append('')
                char_list.append(char)
        
        return pinyin_list, char_list
    
    def _create_pinyin_table(self, slide, left, top, width, height, text, font_size=24):
        """创建拼音-汉字对齐表格
        
        表格结构：
        ┌─────┬─────┬─────┬─────┐
        │ zhè │ shì │ bái │ sè  │  ← 拼音行
        ├─────┼─────┼─────┼─────┤
        │ 这  │ 是  │ 白  │ 色  │  ← 汉字行
        └─────┴─────┴─────┴─────┘
        
        每个单元格包含一个拼音-汉字对，保证精确对齐
        """
        pinyin_list, char_list = self._parse_pinyin_chars(text)
        
        if not char_list:
            return None
        
        num_chars = len(char_list)
        
        # 创建表格：2行 x N列
        rows = 2
        cols = num_chars
        
        # 使用幻灯片的尺寸单位（EMU）
        table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
        table = table_shape.table
        
        # 设置列宽（均分）
        col_width = width // cols
        for col_idx in range(cols):
            table.columns[col_idx].width = col_width
        
        # 设置行高
        table.rows[0].height = height // 2  # 拼音行
        table.rows[1].height = height // 2  # 汉字行
        
        # 填充表格
        for col_idx, (py, char) in enumerate(zip(pinyin_list, char_list)):
            # 拼音单元格（第0行）
            pinyin_cell = table.cell(0, col_idx)
            pinyin_cell.text = py
            pinyin_para = pinyin_cell.text_frame.paragraphs[0]
            pinyin_para.font.size = Pt(font_size)
            pinyin_para.font.name = 'Arial'
            pinyin_para.font.color.rgb = RGBColor(0, 0, 0)
            pinyin_para.alignment = PP_ALIGN.CENTER
            pinyin_cell.vertical_anchor = MSO_ANCHOR.BOTTOM
            
            # 汉字单元格（第1行）
            char_cell = table.cell(1, col_idx)
            char_cell.text = char
            char_para = char_cell.text_frame.paragraphs[0]
            char_para.font.size = Pt(font_size)
            char_para.font.name = 'SimSun'
            char_para.font.color.rgb = RGBColor(0, 0, 0)
            char_para.alignment = PP_ALIGN.CENTER
            char_cell.vertical_anchor = MSO_ANCHOR.TOP
        
        # 隐藏表格边框
        self._hide_table_borders(table)
        
        return table_shape
    
    def _hide_table_borders(self, table):
        """隐藏表格边框"""
        tbl = table._tbl
        
        # 设置表格样式为无框线
        tblPr = tbl.tblPr
        if tblPr is None:
            tblPr = parse_xml(r'<a:tblPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>')
            tbl.insert(0, tblPr)
        
        # 遍历所有单元格，移除边框
        for row in table.rows:
            for cell in row.cells:
                tcPr = cell._tc.get_or_add_tcPr()
                # 设置边框为无
                lnL = parse_xml(r'<a:lnL w="0" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:noFill/></a:lnL>')
                lnR = parse_xml(r'<a:lnR w="0" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:noFill/></a:lnR>')
                lnT = parse_xml(r'<a:lnT w="0" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:noFill/></a:lnT>')
                lnB = parse_xml(r'<a:lnB w="0" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:noFill/></a:lnB>')
                
                # 尝试添加边框设置
                try:
                    tcPr.append(lnL)
                    tcPr.append(lnR)
                    tcPr.append(lnT)
                    tcPr.append(lnB)
                except:
                    pass
    
    def _replace_placeholder_with_pinyin_table(self, slide, placeholder_type, content, font_size=24):
        """替换占位符为拼音-汉字表格"""
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text
                if placeholder_type in text:
                    # 获取shape的位置和大小
                    left = shape.left
                    top = shape.top
                    width = shape.width
                    height = shape.height
                    
                    # 创建拼音表格
                    table_shape = self._create_pinyin_table(slide, left, top, width, height, content, font_size)
                    
                    if table_shape:
                        # 隐藏原始shape（移除文本）
                        shape.text_frame.clear()
                        # 将原始shape移到不可见位置
                        shape.left = Emu(0)
                        shape.top = Emu(0)
                        shape.width = Emu(0)
                        shape.height = Emu(0)
                    
                    return True
        return False
    
    def _create_pinyin_text(self, text_frame, text, font_size=24):
        """创建双行对齐的拼音+汉字文本（备用方案，用于无表格情况）"""
        text_frame.clear()
        
        pinyin_list, char_list = self._parse_pinyin_chars(text)
        
        # 拼音行
        pinyin_text = ' '.join(pinyin_list)
        char_text = ''.join(char_list)
        
        pinyin_para = text_frame.paragraphs[0] if text_frame.paragraphs else text_frame.add_paragraph()
        pinyin_para.text = pinyin_text
        pinyin_para.font.size = Pt(font_size)
        pinyin_para.font.name = 'Arial'
        pinyin_para.alignment = PP_ALIGN.CENTER
        
        char_para = text_frame.add_paragraph()
        char_para.text = char_text
        char_para.font.size = Pt(font_size)
        char_para.font.name = 'SimSun'
        char_para.alignment = PP_ALIGN.CENTER
        char_para.space_before = Pt(4)
    
    def _replace_placeholder_with_pinyin(self, slide, placeholder_type, content, font_size=24):
        """替换占位符并添加拼音-汉字对齐（优先使用表格）"""
        return self._replace_placeholder_with_pinyin_table(slide, placeholder_type, content, font_size)
    
    def _replace_placeholder_simple(self, slide, placeholder_type, content, font_size=24):
        """简单替换占位符"""
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if placeholder_type in run.text:
                            run.text = content
                            return True
        return False
    
    def _find_placeholder_shape(self, slide, placeholder_type):
        """查找包含指定占位符的shape"""
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if placeholder_type in run.text:
                            return shape
        return None
    
    def _find_all_placeholders(self, slide):
        """查找页面所有占位符"""
        placeholders = {}
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text
                matches = re.findall(r'\{\{(\w+)\}\}', text)
                for match in matches:
                    if match not in placeholders:
                        placeholders[match] = shape
        return placeholders
    
    def _match_slides(self):
        """匹配模板页"""
        matched_slides = []
        
        # 1. 封面页
        cover_slide = None
        for slide in self.presentation.slides:
            if self._find_placeholder_shape(slide, 'h0_0'):
                cover_slide = slide
                break
        if cover_slide:
            matched_slides.append(('cover', cover_slide))
        
        # 2. 分类模板页
        section_slides = []
        content_slides = []
        
        for slide in self.presentation.slides:
            if slide == cover_slide:
                continue
            placeholders = self._find_all_placeholders(slide)
            if 'h1_0' in placeholders and 'h2_0' not in placeholders:
                section_slides.append(slide)
            elif 'h2_0' in placeholders:
                content_slides.append(slide)
        
        # 3. 目录页
        if content_slides:
            toc_slide = self.presentation.slides.add_slide(content_slides[0].slide_layout)
            matched_slides.append(('toc', toc_slide))
        
        # 4. 章节页
        for i, h1 in enumerate(self.md_content['h1']):
            if i < len(section_slides):
                matched_slides.append(('section', section_slides[i]))
            else:
                if section_slides:
                    new_slide = self.presentation.slides.add_slide(section_slides[0].slide_layout)
                    matched_slides.append(('section', new_slide))
                elif content_slides:
                    new_slide = self.presentation.slides.add_slide(content_slides[0].slide_layout)
                    matched_slides.append(('section', new_slide))
        
        # 5. 正文页
        for h2 in self.md_content['h2']:
            if h2.get('video'):
                if content_slides:
                    new_slide = self.presentation.slides.add_slide(content_slides[0].slide_layout)
                    matched_slides.append(('video', new_slide))
            elif h2.get('audio'):
                if content_slides:
                    new_slide = self.presentation.slides.add_slide(content_slides[0].slide_layout)
                    matched_slides.append(('audio', new_slide))
            else:
                if content_slides:
                    matched_slides.append(('content', content_slides[0]))
                else:
                    if cover_slide:
                        new_slide = self.presentation.slides.add_slide(cover_slide.slide_layout)
                        matched_slides.append(('content', new_slide))
        
        # 6. 结束页
        end_slide = None
        for slide in self.presentation.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    if 'xiè xie' in shape.text_frame.text or '谢谢' in shape.text_frame.text:
                        end_slide = slide
                        break
            if end_slide:
                break
        if end_slide:
            matched_slides.append(('end', end_slide))
        
        return matched_slides
    
    def _generate_content(self):
        """生成PPT内容"""
        # 1. 封面页
        if self.md_content['h0'] and len(self.presentation.slides) > 0:
            cover_slide = self.presentation.slides[0]
            title = self.md_content['h0'][0]
            self._replace_placeholder_with_pinyin(cover_slide, 'h0_0', title, font_size=36)
        
        # 2. 目录页
        if len(self.presentation.slides) > 1 and self.md_content['h1']:
            toc_slide = self.presentation.slides[1]
            for i, h1 in enumerate(self.md_content['h1']):
                if i == 0:
                    self._replace_placeholder_with_pinyin(toc_slide, 'h2_0', h1, font_size=24)
                else:
                    placeholder = f'h3_{i-1}' if i <= 4 else None
                    if placeholder:
                        self._replace_placeholder_with_pinyin(toc_slide, placeholder, h1, font_size=18)
        
        # 3. 章节页和正文页
        slide_idx = 2
        section_count = len(self.md_content['h1'])
        
        # 章节页
        for i in range(section_count):
            if slide_idx < len(self.presentation.slides):
                section_slide = self.presentation.slides[slide_idx]
                self._replace_placeholder_with_pinyin(section_slide, 'h1_0', self.md_content['h1'][i], font_size=32)
                slide_idx += 1
        
        # 正文页
        h2_index = 0
        while h2_index < len(self.md_content['h2']) and slide_idx < len(self.presentation.slides):
            h2 = self.md_content['h2'][h2_index]
            content_slide = self.presentation.slides[slide_idx]
            
            # 小节名
            self._replace_placeholder_with_pinyin(content_slide, 'h2_0', h2['title'], font_size=28)
            
            # 文本框
            for j, text in enumerate(h2.get('content', [])):
                if j <= 3:
                    placeholder = f'h3_{j}'
                    self._replace_placeholder_with_pinyin(content_slide, placeholder, text, font_size=20)
            
            # 图片
            for j in range(len(h2.get('images', []))):
                if j <= 4:
                    placeholder = f'img_{j}'
                    self._replace_placeholder_simple(content_slide, placeholder, f'[图片{j+1}]')
            
            # 音频
            if h2.get('audio'):
                self._replace_placeholder_simple(content_slide, 'audio_0', '[音频播放]')
            
            # 视频
            if h2.get('video'):
                self._replace_placeholder_simple(content_slide, 'video_0', '[视频播放]')
            
            slide_idx += 1
            h2_index += 1
    
    def generate(self):
        """生成PPT"""
        self._match_slides()
        self._generate_content()
        output_path = 'output.pptx'
        self.presentation.save(output_path)
        return output_path


if __name__ == '__main__':
    import sys
    from md_parser import MDParser
    
    if len(sys.argv) > 2:
        md_file = sys.argv[1]
        template_file = sys.argv[2]
        
        print(f"解析MD: {md_file}")
        parser = MDParser(md_file)
        md_content = parser.parse()
        
        print(f"生成PPT: {template_file} -> output.pptx")
        generator = PPTGenerator(template_file, md_content)
        output = generator.generate()
        print(f"完成: {output}")
    else:
        print("用法: python ppt_generator.py <md文件> <模板文件>")