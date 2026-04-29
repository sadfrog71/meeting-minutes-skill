# Meeting Minutes Skill - 会议纪要整理

将原始会议记录整理为结构化、可执行的会议纪要，并支持导出标准格式 Word 文档。

## 功能特性

- ✅ 会议记录解析与结构化整理
- ✅ 标准格式输出（Markdown）
- ✅ Word 文档导出（企业风格）
- ✅ 责任分工表格化
- ✅ 后续工作可执行化

## 文件结构

```
meeting-minutes-skill/
├── SKILL.md                      # 核心规范
├── README.md                     # 说明文档
├── scripts/
│   └── generate_docx.py          # Word 文档生成脚本
├── templates/
│   └── output-template.md         # 输出模板
└── examples/
    ├── sample-input.md            # 示例输入
    └── sample-output.md           # 示例输出
```

## 快速开始

### 1. 整理会议纪要

将原始会议记录发给 AI，自动整理为结构化格式。

### 2. 导出 Word

```bash
# 安装依赖
pip install python-docx

# 生成 Word 文档
python scripts/generate_docx.py -o output.docx \
  --company "公司名称" \
  --date "2026年3月17日" \
  --title "公司官网改版讨论会" \
  --time "2026年3月16日 14:00-15:00" \
  --location "2号会议室" \
  --theme "公司官网改版讨论会" \
  --attendees "汪宠、刘朝阳、徐泓卿..." \
  --host "汪宠" \
  --absent "马跃" \
  --recorder "刘朝阳"
```

## Word 格式规范

| 元素 | 样式 |
|------|------|
| 公司名称+日期 | 居中、黑体、14pt、深蓝 |
| 会议标题 | 居中、黑体、16pt、深蓝 |
| 信息表格 | 标签列深青底白字 |
| 章节标题 | 黑体、12pt、深青 |
| 正文 | 宋体、10.5pt |
| 表格边框 | 浅灰线 |

## 触发关键词

- 会议纪要、整理会议、meeting minutes
- 会议记录、会议总结
- 生成会议纪要、导出会议纪要
