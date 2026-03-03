# -*- coding: utf-8 -*-
"""
PPT生成器
支持：1-4文本框、0-5图片、音视频模板匹配、拼音标注
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import re
from pypinyin import pinyin, Style

class PPTGenerator:
    def __init__(self, template_path, md_content):
        self.template_path = template_path
        self.md_content = md_content
        self.presentation = Presentation(template_path)
    
    def _add_pinyin(self, text, font_size=24):
        """为汉字添加拼音，拼音显示在汉字上方，拼音字号=汉字字号"""
        result = []
        for char in text:
            if '\u4e00' <= char <= '\u9fff':  # 判断是否为汉字
                py = pinyin(char, style=Style.TONE)[0][0]
                # 拼音在汉字上方，用斜杠分隔
                result.append(f'{py}/{char}')
            else:
                result.append(char)
        return ''.join(result)
    
    def _add_pinyin_rich(self, text_frame, font_size=24):
        """为文本框中的文字添加拼音（富文本方式）"""
        # 清除现有内容
        text_frame.clear()
        
        # 创建拼音和汉字组合
        for char in text_frame.add_paragraph():
            if '\u4e00' <= char <= '\u9fff':
                # 是汉字，添加拼音
                py = pinyin(char, style=Style.TONE)[0][0]
                # 拼音在上的格式：拼音(小字)在汉字上面
                # 这里简化处理：用特殊格式标记，后续可优化
                run = text_frame.add_run()
                run.text = f'{py} '  # 拼音
                run.font.size = Pt(font_size // 2)  # 拼音字号小一半
                
                run2 = text_frame.add_run()
                run2.text = char  # 汉字
                run2.font.size = Pt(font_size)
            else:
                run = text_frame.add_run()
                run.text = char
                run.font.size = Pt(font_size)
    
    def _find_placeholder(self, slide, placeholder_type):
        """查找指定类型的占位符"""
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if placeholder_type in run.text:
                            return shape
        return None
    
    def _find_all_placeholders(self, slide):
        """查找页面所有占位符"""
        placeholders = {}
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text
                # 查找所有 {{xxx}} 格式的占位符
                matches = re.findall(r'\{\{(\w+)\}\}', text)
                for match in matches:
                    if match not in placeholders:
                        placeholders[match] = shape
        return placeholders
    
    def _replace_placeholder(self, slide, placeholder_type, content):
        """替换占位符内容"""
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if placeholder_type in run.text:
                            # 替换文本，添加拼音
                            pinyin_text = self._add_pinyin(content, font_size=24)
                            run.text = pinyin_text
                            return True
        return False
    
    def _get_text_box_count(self, h2_item):
        """获取该小节需要几个文本框"""
        return len(h2_item.get('content', []))
    
    def _get_image_count(self, h2_item):
        """获取该小节需要几张图片"""
        return len(h2_item.get('images', []))
    
    def _match_slides(self):
        """匹配模板页，根据内容自动匹配对应的模板"""
        matched_slides = []
        
        # 1. 封面页
        cover_slide = None
        for slide in self.presentation.slides:
            if self._find_placeholder(slide, 'h0_0'):
                cover_slide = slide
                break
        if cover_slide:
            matched_slides.append(('cover', cover_slide))
        
        # 2. 目录页
        toc_slide = None
        for slide in self.presentation.slides:
            if self._find_placeholder(slide, 'h1_0') and slide != cover_slide:
                toc_slide = slide
                break
        if toc_slide:
            matched_slides.append(('toc', toc_slide))
        
        # 3. 分类模板页
        section_slides = []  # 章节页
        content_slides = []  # 正文页（普通）
        audio_slides = []    # 音频页（带喇叭）
        video_slides = []    # 视频页
        
        for slide in self.presentation.slides:
            if slide == cover_slide or slide == toc_slide:
                continue
            placeholders = self._find_all_placeholders(slide)
            # 判断模板类型
            if 'audio_0' in placeholders:
                audio_slides.append(slide)
            elif 'video_0' in placeholders:
                video_slides.append(slide)
            elif 'h1_0' in placeholders and 'h2_0' not in placeholders:
                section_slides.append(slide)
            elif 'h2_0' in placeholders:
                content_slides.append(slide)
        
        # 4. 章节页匹配
        for i, h1 in enumerate(self.md_content['h1']):
            if i < len(section_slides):
                matched_slides.append(('section', section_slides[i]))
            else:
                if section_slides:
                    new_slide = self.presentation.slides.add_slide(section_slides[0].slide_layout)
                    matched_slides.append(('section', new_slide))
        
        # 5. 正文页匹配（包含音视频）
        for h2 in self.md_content['h2']:
            text_count = self._get_text_box_count(h2)
            image_count = self._get_image_count(h2)
            
            # 判断页面类型
            if h2.get('video'):
                # 视频单独一页
                if video_slides:
                    matched_slides.append(('video', video_slides[0]))
                else:
                    # 没有视频模板，复制普通模板
                    if content_slides:
                        new_slide = self.presentation.slides.add_slide(content_slides[0].slide_layout)
                        matched_slides.append(('video', new_slide))
            elif h2.get('audio'):
                # 音频+文本
                if audio_slides:
                    matched_slides.append(('audio', audio_slides[0]))
                else:
                    if content_slides:
                        new_slide = self.presentation.slides.add_slide(content_slides[0].slide_layout)
                        matched_slides.append(('audio', new_slide))
            else:
                # 普通正文页
                if content_slides:
                    matched_slides.append(('content', content_slides[0]))
                else:
                    # 复制封面页作为后备
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
            self._replace_placeholder(cover_slide, 'h0_0', title)
        
        # 2. 目录页
        if len(self.presentation.slides) > 1:
            toc_slide = self.presentation.slides[1]
            for i, h1 in enumerate(self.md_content['h1']):
                placeholder = f'h1_{i}'
                self._replace_placeholder(toc_slide, placeholder, h1)
        
        # 3. 正文页内容
        slide_idx = 2
        h2_index = 0
        
        # 处理章节页
        for i, h1 in enumerate(self.md_content['h1']):
            if slide_idx < len(self.presentation.slides):
                section_slide = self.presentation.slides[slide_idx]
                self._replace_placeholder(section_slide, 'h1_0', h1)
                slide_idx += 1
        
        # 处理正文页（包括音视频）
        while h2_index < len(self.md_content['h2']) and slide_idx < len(self.presentation.slides):
            h2 = self.md_content['h2'][h2_index]
            content_slide = self.presentation.slides[slide_idx]
            
            # 替换小节名
            self._replace_placeholder(content_slide, 'h2_0', h2['title'])
            
            # 替换文本框内容 (h3_0 ~ h3_3)
            for j, text in enumerate(h2.get('content', [])):
                if j <= 3:  # 最多4个文本框
                    placeholder = f'h3_{j}'
                    self._replace_placeholder(content_slide, placeholder, text)
            
            # 替换图片 (img_0 ~ img_4)
            for j, img_url in enumerate(h2.get('images', [])):
                if j <= 4:  # 最多5张图片
                    placeholder = f'img_{j}'
                    # TODO: 下载图片并插入PPT
                    # 目前先跳过，后续完善
            
            # 替换音频
            if h2.get('audio'):
                self._replace_placeholder(content_slide, 'audio_0', h2['audio'])
            
            # 替换视频
            if h2.get('video'):
                self._replace_placeholder(content_slide, 'video_0', h2['video'])
            
            slide_idx += 1
            h2_index += 1
    
    def generate(self):
        """生成PPT"""
        self._match_slides()
        self._generate_content()
        output_path = 'output.pptx'
        self.presentation.save(output_path)
        return output_path


# 测试用
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