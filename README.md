# zongce-proof-compiler（综测活动证明汇总）

一个面向 AI Agent（WorkBuddy / Claude Code 等支持 Agent Skills 的平台）的技能包：根据用户提供的**综测评定细则**、**本人姓名**和一批**活动证明文件**（图片 / PDF / Word），自动定位包含该姓名的证明页、按细则分类定分，并生成带加分汇总表的证明汇总 Word。

## 功能

- **姓名定位**：PDF 文字层搜索优先；扫描件 / 大型合集 PDF（几十上百页）使用本地中文 OCR（RapidOCR）三级匹配（精确 / 宽松 / 部分）预筛候选页
- **多格式输入**：图片、PDF、Word（.docx/.doc 经本机 MS Word 转 PDF）
- **跨页证明处理**：命中页相邻页一并渲染，一份证明两页全部保留
- **人工兜底**：OCR 仅用于定位，所有候选页须经 Agent 视觉确认；形近异名（如"刘俊阳"）自动标出待核
- **汇总 Word 输出**：封面 + 加分汇总表（含合计行）+ 按细则分类分章 + 一页一图带标题
- **隐私安全**：全程本地处理，无网络请求，不需要任何 API Key

## 目录结构

```
zongce-proof-compiler/
├── SKILL.md                  # 技能说明与工作流程（Agent 读取）
├── requirements.txt          # Python 依赖
├── scripts/
│   ├── extract_pages.py      # 定位并渲染含姓名的页面（图片/PDF/Word）
│   ├── ocr_scan.py           # 扫描版 PDF 的 OCR 预筛（中文）
│   └── build_docx.py         # 按 plan JSON 生成汇总 Word
├── .env.example              # 环境变量示例（本技能无需密钥）
├── .gitignore
└── LICENSE                   # MIT
```

## 安装

**作为 Agent Skill（WorkBuddy）**：将本目录放入 `~/.workbuddy/skills/`（用户级）或项目 `.workbuddy/skills/`（项目级），在对话中说"整理综测证明"即可触发。

**作为普通 Python 工具**：

```bash
pip install -r requirements.txt
```

> 注意：处理 .docx/.doc 输入需要 Windows 本机安装 Microsoft Word；无 Word 时 PDF 与图片输入不受影响。

## 使用

向 Agent 提供三样东西：

1. 《综测评定细则》（任意格式）
2. 你的姓名（精确到字）
3. 活动证明文件（可多份）

Agent 将按 SKILL.md 工作流执行：解析细则 → 定位姓名页 → 视觉核对（存疑必问）→ 分类定分 → 生成并交付汇总 Word。

脚本也可独立使用：

```bash
# 定位并渲染含"张三"的页面（含相邻页）
python scripts/extract_pages.py --outdir out --name 张三 --neighbors proof1.pdf proof2.docx

# 扫描版 PDF 的 OCR 预筛
python scripts/ocr_scan.py --name 张三 --dpi 200 --out hits.json big_scanned.pdf

# 由 plan JSON 生成汇总 Word
python scripts/build_docx.py --plan plan.json --out 综测证明汇总.docx
```

## 安全说明

- 无硬编码密钥 / Token / 密码；无网络请求；不读取任何用户配置文件
- 所有中间产物（页面图片、manifest）仅写入用户指定的工作目录
- 证明文件涉及个人隐私，请勿将提取的页面图片上传至公共仓库

## License

MIT
