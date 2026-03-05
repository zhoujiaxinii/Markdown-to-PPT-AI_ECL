# -*- coding: utf-8 -*-
"""
PPT生成器 - V12
在模板上直接操作，XML深拷贝幻灯片，用rId追踪排序
"""

from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.parts.slide import SlidePart
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI
from lxml import etree
import re, copy
from pypinyin import pinyin, Style

NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


class PPTGenerator:
    def __init__(self, template_path, md_content):
        self.template_path = template_path
        self.md_content = md_content
        self.prs = Presentation(template_path)
        self._analyze_templates()

    # ── 模板分析 ──────────────────────────────────────────

    def _analyze_templates(self):
        self.templates = {
            'cover': None, 'toc': None, 'section': None, 'end': None,
            'content_1': [], 'content_2': [], 'content_3': [], 'content_4': [],
        }
        for idx, slide in enumerate(self.prs.slides):
            ph = self._find_all_placeholders(slide)
            if any('谢' in s.text_frame.text or 'xiè' in s.text_frame.text.lower()
                   for s in slide.shapes if s.has_text_frame):
                self.templates['end'] = idx; continue
            if 'h0_0' in ph:
                self.templates['cover'] = idx; continue
            if 'h1_0' in ph and 'h2_0' not in ph:
                self.templates['section'] = idx; continue
            if 'h2_0' in ph:
                n = len([p for p in ph if p.startswith('h3_')])
                self.templates[f'content_{min(n,4)}'].append(idx); continue
            for s in slide.shapes:
                if s.has_text_frame and ('目录' in s.text_frame.text or 'mù lù' in s.text_frame.text.lower()):
                    self.templates['toc'] = idx; break

        _p = lambda k: f"第{self.templates[k]+1}页" if self.templates[k] is not None else "无"
        cnt = lambda k: len(self.templates[f'content_{k}'])
        print(f"\n=== 模板分析 ===")
        print(f"封面:{_p('cover')} 目录:{_p('toc')} 章节:{_p('section')} 结束:{_p('end')}")
        print(f"正文: 1框×{cnt(1)} 2框×{cnt(2)} 3框×{cnt(3)} 4框×{cnt(4)}")

    # ── 幻灯片XML深拷贝 ──────────────────────────────────

    def _clone_slide(self, source_idx):
        """XML深拷贝，返回新幻灯片的rId"""
        src = self.prs.slides[source_idx]
        layout_part = src.slide_layout.part

        existing_nums = []
        for p in self.prs.part.package.iter_parts():
            pn = str(p.partname)
            if '/slides/slide' in pn and 'Layout' not in pn:
                try: existing_nums.append(int(pn.split('/')[-1].replace('slide','').replace('.xml','')))
                except: pass
        new_partname = PackURI(f'/ppt/slides/slide{max(existing_nums)+1}.xml')

        new_part = SlidePart.new(new_partname, self.prs.part.package, layout_part)
        new_part._element = copy.deepcopy(src._element)

        for rel in src.part.rels.values():
            if rel.reltype == RT.SLIDE_LAYOUT: continue
            if rel.is_external:
                new_part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
            else:
                new_part.relate_to(rel.target_part, rel.reltype, rel.rId)

        rId = self.prs.part.relate_to(new_part, RT.SLIDE)
        new_id = max(int(s.get('id')) for s in self.prs.slides._sldIdLst) + 1
        sldId = etree.SubElement(self.prs.slides._sldIdLst, f'{{{NS_P}}}sldId')
        sldId.set('id', str(new_id))
        sldId.set(f'{{{NS_R}}}id', rId)

        return rId

    def _get_slide_by_rId(self, rId):
        """通过rId获取幻灯片对象"""
        return self.prs.part.related_slide(rId)

    def _get_rId_by_idx(self, idx):
        """获取原始模板页的rId"""
        return self.prs.slides._sldIdLst[idx].rId

    # ── 拼音表格 ─────────────────────────────────────────

    def _parse_pinyin(self, text):
        py_list, ch_list = [], []
        for c in text:
            if '\u4e00' <= c <= '\u9fff':
                py = pinyin(c, style=Style.TONE)[0][0]
                py_list.append(re.sub(r'\d$', '', py)); ch_list.append(c)
            elif c.strip():
                py_list.append(''); ch_list.append(c)
        return py_list, ch_list

    def _create_pinyin_table(self, slide, left, top, width, height, text, fs=24):
        py_list, ch_list = self._parse_pinyin(text)
        if not ch_list: return None
        emu = lambda pt: int(pt * 914400 / 72)
        def total_w(f):
            return sum(emu(f*0.6)*(len(p) if p else 1) + emu(f*0.5) for p in py_list)
        tw = total_w(fs)
        if tw > width:
            fs = max(int(fs * width / tw), 10); tw = total_w(fs)
        cw = [emu(fs*0.6)*(len(p) if p else 1) + emu(fs*0.5) for p in py_list]
        tbl_shape = slide.shapes.add_table(2, len(ch_list), left, top, int(tw), height)
        tbl = tbl_shape.table
        for i, w in enumerate(cw): tbl.columns[i].width = int(w)
        tbl.rows[0].height = int(height * 0.45)
        tbl.rows[1].height = int(height * 0.55)
        for ci, (p, c) in enumerate(zip(py_list, ch_list)):
            for ri, txt in enumerate([p, c]):
                cell = tbl.cell(ri, ci)
                cell.text = txt; cell.text_frame.word_wrap = False
                pa = cell.text_frame.paragraphs[0]
                pa.font.size = Pt(fs)
                pa.font.name = 'Arial' if ri == 0 else 'SimSun'
                pa.font.color.rgb = RGBColor(0, 0, 0)
                pa.alignment = PP_ALIGN.CENTER
                cell.vertical_anchor = MSO_ANCHOR.BOTTOM if ri == 0 else MSO_ANCHOR.TOP
        self._hide_borders(tbl)
        return tbl_shape

    def _hide_borders(self, tbl):
        """彻底隐藏表格边框和背景，实现真正透明"""
        ns_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'

        # 1. 移除内置表格样式（这是白色底色的根源）
        tblPr = tbl._tbl.tblPr
        if tblPr is not None:
            for style_id in tblPr.findall(f'{{{ns_a}}}tableStyleId'):
                tblPr.remove(style_id)
            tblPr.set('firstRow', '0')
            tblPr.set('bandRow', '0')
            tblPr.set('firstCol', '0')
            tblPr.set('lastRow', '0')
            tblPr.set('lastCol', '0')

        # 2. 每个单元格：四边框 noFill + 单元格背景 noFill
        for row in tbl.rows:
            for cell in row.cells:
                tcPr = cell._tc.get_or_add_tcPr()
                # 清除已有的边框和填充
                for ch in list(tcPr):
                    tag = ch.tag.split('}')[-1] if '}' in ch.tag else ch.tag
                    if tag in ('lnL','lnR','lnT','lnB','solidFill','noFill'):
                        tcPr.remove(ch)
                # 显式设置四边框为无
                for border in ('lnL', 'lnR', 'lnT', 'lnB'):
                    ln = etree.SubElement(tcPr, f'{{{ns_a}}}{border}')
                    ln.set('w', '0')
                    ln.set('cap', 'flat')
                    etree.SubElement(ln, f'{{{ns_a}}}noFill')
                # 单元格背景透明
                etree.SubElement(tcPr, f'{{{ns_a}}}noFill')

    # ── 占位符操作 ────────────────────────────────────────

    def _find_all_placeholders(self, slide):
        ph = {}
        for s in slide.shapes:
            if s.has_text_frame:
                for m in re.findall(r'\{\{(\w+)\}\}', s.text_frame.text):
                    ph.setdefault(m, s)
        return ph

    def _fill(self, slide, name, content, fs=24):
        ph = self._find_all_placeholders(slide)
        if name not in ph: return False
        s = ph[name]
        self._create_pinyin_table(slide, s.left, s.top, s.width, s.height, content, fs)
        s.text_frame.clear(); s.left = Emu(0)
        return True

    def _clear_unused(self, slide, used):
        for name, s in self._find_all_placeholders(slide).items():
            if name not in used:
                s.text_frame.clear(); s.left = Emu(0)

    # ── 填充单页 ─────────────────────────────────────────

    def _fill_slide(self, slide, typ, data, num):
        if typ == 'cover':
            self._fill(slide, 'h0_0', data, 36)
            self._clear_unused(slide, ['h0_0'])
            print(f"  {num}. 封面: {data}")
        elif typ == 'toc':
            for j, s in enumerate(data):
                self._fill(slide, f'h1_{j}', s, 20)
            print(f"  {num}. 目录: {len(data)}章")
        elif typ == 'section':
            self._fill(slide, 'h1_0', data, 32)
            self._clear_unused(slide, ['h1_0'])
            print(f"  {num}. 章节: {data}")
        elif typ == 'content':
            used = ['h2_0']
            self._fill(slide, 'h2_0', data['title'], 28)
            for j, t in enumerate(data.get('content', [])):
                self._fill(slide, f'h3_{j}', t, 20); used.append(f'h3_{j}')
            self._clear_unused(slide, used)
            print(f"  {num}. 正文: {data['title']}")
        elif typ == 'end':
            print(f"  {num}. 结束页")

    # ── 生成 ─────────────────────────────────────────────

    def generate(self):
        print("\n=== 开始生成PPT ===")

        # 构建页面计划
        pages = []
        ptr = {1: 0, 2: 0, 3: 0, 4: 0}
        def next_content(n):
            k = min(n, 4)
            ts = self.templates[f'content_{k}']
            if not ts: return None
            idx = ts[ptr[k] % len(ts)]; ptr[k] += 1; return idx

        if self.md_content.get('h0') and self.templates['cover'] is not None:
            pages.append(('cover', self.templates['cover'], self.md_content['h0'][0]))
        if self.md_content.get('h1') and self.templates['toc'] is not None:
            pages.append(('toc', self.templates['toc'], self.md_content['h1']))

        cur_sec = None
        for h2 in self.md_content['h2']:
            sec = h2.get('section', '')
            if sec and sec != cur_sec and self.templates['section'] is not None:
                cur_sec = sec
                pages.append(('section', self.templates['section'], sec))
            idx = next_content(len(h2.get('content', [])))
            if idx is not None:
                pages.append(('content', idx, h2))

        if self.templates['end'] is not None:
            pages.append(('end', self.templates['end'], None))

        print(f"计划 {len(pages)} 页")

        # 保存原始模板页的 rId 映射
        orig_rIds = {}
        for i in range(len(self.prs.slides)):
            orig_rIds[i] = self.prs.slides._sldIdLst[i].rId

        # 第一步：分配幻灯片（先克隆，不填充）
        used_once = {}         # 模板idx -> True (已使用)
        ordered_rIds = []      # 最终顺序的 rId 列表

        for i, (typ, tidx, data) in enumerate(pages):
            if tidx not in used_once:
                used_once[tidx] = True
                rId = orig_rIds[tidx]
            else:
                rId = self._clone_slide(tidx)
            ordered_rIds.append(rId)

        # 第二步：填充内容（克隆完成后再填充，避免克隆已填充的内容）
        for i, (typ, tidx, data) in enumerate(pages):
            rId = ordered_rIds[i]
            slide = self._get_slide_by_rId(rId)
            self._fill_slide(slide, typ, data, i + 1)

        # 删除未使用的模板页
        keep_rIds = set(ordered_rIds)
        sldIdLst = self.prs.slides._sldIdLst
        to_del = [s for s in list(sldIdLst) if s.rId not in keep_rIds]
        for sldId in to_del:
            self.prs.part.drop_rel(sldId.rId)
            sldIdLst.remove(sldId)
        print(f"删除 {len(to_del)} 页未使用模板")

        # 按最终顺序重排 sldIdLst
        rId_to_sldId = {s.rId: s for s in list(sldIdLst)}
        for s in list(sldIdLst):
            sldIdLst.remove(s)
        for rId in ordered_rIds:
            sldIdLst.append(rId_to_sldId[rId])

        out = 'output.pptx'
        self.prs.save(out)
        print(f"\n=== 完成：{out}，共 {len(self.prs.slides)} 页 ===")
        return out


if __name__ == '__main__':
    import sys
    from md_parser import MDParser
    if len(sys.argv) > 2:
        parser = MDParser(sys.argv[1])
        PPTGenerator(sys.argv[2], parser.parse()).generate()
    else:
        print("用法: python ppt_generator.py <md文件> <模板文件>")