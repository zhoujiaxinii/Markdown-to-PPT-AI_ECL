# Markdown转PPT项目

## 项目功能
- 将Markdown文件转换为PPT文件
- 支持从模板库中匹配合适的模板页
- 自动为汉字添加拼音标注
- 支持处理图片、音频、视频和游戏链接
- 保持模板的原有样式

## 项目结构
```
ppt生成/
├── src/
│   └── utils/
│       ├── md_parser.py      # Markdown解析器
│       └── ppt_generator.py  # PPT生成器
├── test/                    # 测试文件
├── app.py                   # 主入口文件
├── requirements.txt         # 依赖文件
└── README.md                # 项目说明
```

## 安装依赖
```bash
pip install -r requirements.txt
```

## 使用方法

### 1. 启动API服务
```bash
python app.py
```

### 2. 调用API接口

**API地址**: `http://localhost:5000/api/convert`

**请求方法**: POST

**请求参数**:
- `md_file`: Markdown文件（.md格式）
- `ppt_template`: PPT模板文件（.pptx格式）

**返回结果**:
```json
{
  "success": true,
  "output_path": "output.pptx"
}
```

### 3. 直接运行测试
```bash
python test_ppt_generator.py
```

## 模板要求

模板文件必须包含以下占位符：
- 封面页：`{{h0_0}}`
- 目录页：`{{h1_0}}`, `{{h1_1}}`, ...
- 章节页：`{{h1_0}}`
- 正文页：`{{h2_0}}`, `{{h3_0}}`, `{{h4_0}}`, ...

## 注意事项
1. 仅支持.md/.docx格式的Markdown文件
2. 仅支持.pptx格式的PPT模板
3. 模板中必须包含指定的占位符
4. 生成的PPT会保持模板的原有样式
5. 所有汉字会自动添加拼音标注

## 示例

### 输入
- Markdown文件：包含标题、文本和多媒体链接
- PPT模板：包含占位符的模板文件

### 输出
- 生成的PPT文件：包含匹配的模板页和填充的内容

## 技术栈
- Python 3.7+
- Flask：Web API框架
- python-pptx：PPT处理库
- markdown：Markdown解析库
- pypinyin：拼音标注库