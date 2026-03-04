# Markdown转PPT项目 - 项目开发文档最新版

> 更新日期：2026-03-04
> 版本：v1.0.0

## 项目概述

本项目实现将 Markdown 文件自动转换为 PPT 演示文稿，支持：
- 自动为汉字添加拼音标注（精确对齐）
- 动态模板匹配
- 音视频嵌入
- 图片处理

## 核心功能

| 功能 | 描述 | 状态 |
|------|------|------|
| Markdown解析 | 解析MD文件，提取结构化内容 | ✅ 完成 |
| 拼音标注 | 拼音显示在汉字正上方，精确对齐 | ✅ 完成 |
| 动态模板匹配 | 根据内容自动匹配合适模板 | ✅ 完成 |
| 音视频嵌入 | 支持mp3/mp4嵌入 | ✅ 完成 |
| 图片处理 | 支持0-5张图片/页 | ✅ 完成 |

## 文档目录

| 文档 | 说明 |
|------|------|
| [开发需求.md](./开发需求.md) | 项目需求详情 |
| [开发状态.md](./开发状态.md) | 当前开发进度 |
| [项目结构.md](./项目结构.md) | 代码结构说明 |
| [MD撰写规范.md](./MD撰写规范.md) | MD文件撰写规范 |
| [匹配规则说明.md](./匹配规则说明.md) | 内容到PPT的匹配规则 |
| [代码说明.md](./代码说明.md) | 核心代码详解 |

## 快速开始

### 安装依赖
```bash
pip install -r requirements.txt
```

### 运行测试
```bash
python -c "
from src.utils.md_parser import MDParser
from src.utils.ppt_generator import PPTGenerator

parser = MDParser('test_sample.md')
md_content = parser.parse()

generator = PPTGenerator('案例/模板.pptx', md_content)
output = generator.generate()
print(f'生成完成: {output}')
"
```

## 技术栈

- Python 3.7+
- python-pptx：PPT处理
- pypinyin：拼音标注
- lxml：XML操作

## 仓库地址

https://github.com/zhoujiaxinii/Markdown-to-PPT-AI_ECL
