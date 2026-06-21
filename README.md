# report-to-ppt

把内容报告（Markdown）转成专业可编辑的 PowerPoint 演示文稿。

## 安装

### 方式一：从 GitHub Release 下载

1. 在 [Releases](https://github.com/Ciphercici/report-to-ppt/releases) 页面下载 `report-to-ppt.zip`
2. 解压后为report-to-ppt.skill
3. 解压到 Claude Code 的 skills 目录：

**Windows (PowerShell):**
```powershell
Expand-Archive report-to-ppt.skill -DestinationPath "$env:USERPROFILE\.claude\skills\report-to-ppt"
```

**macOS / Linux:**
```bash
unzip report-to-ppt.skill -d ~/.claude/skills/report-to-ppt/
```

### 方式二：Git Clone

```bash
git clone https://github.com/Ciphercici/report-to-ppt.git ~/.claude/skills/report-to-ppt/
```

### 方式三：手动复制

把整个 `report-to-ppt/` 目录放到 `~/.claude/skills/` 下即可。

### 验证安装

```bash
ls ~/.claude/skills/report-to-ppt/SKILL.md && echo "安装成功"
```

> 依赖：`pip install python-pptx`（Python 3.8+，无其他依赖）

## 为什么不用 Pandoc / Marp

| 工具 | 问题 |
|------|------|
| **Pandoc** | 按 `--slide-level` 切页，长章节挤成一张、短章节空白 |
| **Marp** | PPTX 输出是图片背景，文字不可编辑 |
| **python-pptx 直转** | 只切段落不改内容，文字堆砌，没有信息层级 |

报告体 md 不能直接"格式转换"为 PPT。必须先做**内容提炼和叙事重组**，再逐页排版。

## 工作流

```
原始报告 .md
     │
     ▼  Claude 读懂内容，设计叙事结构
.slide.md 大纲（Marp 格式，12-18 页）
     │
     ▼  build_ppt.py 自动排版
输出 .pptx（文字可编辑，设计专业）
```

### 第一步：写大纲

Claude 读报告 → 提炼关键信息 → 按叙事逻辑写成 `.slide.md`：

```markdown
---
marp: true
size: 16:9
---

<!-- _class: lead -->
# 报告标题
## 副标题

---

<!-- _class: lead -->
## 摘要
核心发现段落...

- 要点一
- 要点二

---

## 章节标题
- 论据一
- 论据二

![bg right:40%](./assets/证据截图.png)
```

格式规则：
- `---` = 新页
- `<!-- _class: lead -->` = 深色背景（封面/章节过渡/总结）
- `![bg right:40%](path)` = 右侧图片
- `**粗体数字**` = 自动识别为 KPI 卡片

### 第二步：生成 PPTX

```bash
python scripts/build_ppt.py outline.slide.md -o output.pptx
```

脚本自动：
- 解析 Marp 格式大纲
- 套用设计系统（标题栏、KPI 卡片、页码、配色）
- 图片等比缩放 + 边框
- `###` 子标题渲染为小节标题
- 表格 markdown 转键值对

### 第三步：验证

```bash
python -c "
from pptx import Presentation
prs = Presentation('output.pptx')
for i, s in enumerate(prs.slides):
    texts = [sh.text_frame.text[:60] for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
    imgs = sum(1 for sh in s.shapes if sh.shape_type == 13)
    print(f'{i+1:2d}: {len(s.shapes)} shapes, {imgs} img | {texts[0] if texts else \"(empty)\"}')
"
```

正常输出：15±3 页，每页 5-25 个形状，文件 > 1 MB。

## 三种模式

| 模式 | 命令 | 质量 | 适用 |
|------|------|------|------|
| **Outline**（推荐） | `build_ppt.py report.slide.md` | 高 | 最终交付 |
| **Auto**（快速草稿） | `build_ppt.py report.md` | 中 | 快速看结构 |
| **Manual**（精细控制） | 编辑脚本中 slide 函数 | 最高 | 微调个别页面 |

### Outline 模式

Claude 设计 `.slide.md` 大纲 → 脚本排版。内容提炼由 Claude 完成，排版由脚本完成，各司其职。

### Auto 模式

脚本直接解析原始报告 md，自动拆章节、提取图片、检测 KPI、分页。速度快但要点质量取决于原文写法。

### Manual 模式

复制 `build_ppt.py` 到项目目录，编辑 slide 函数，完全控制每页内容和布局。用于 outline 模式无法满足的个别页面。

## 设计系统

| 元素 | 规格 |
|------|------|
| 尺寸 | 16:9 (13.333" × 7.5") |
| 主色 | `#3B82C4` (蓝) |
| 强调色 | `#E84D4D` (红) |
| 标题 | 30pt, `#1A1A1A`, Bold |
| 正文 | 17pt, `#444444` |
| KPI 卡片 | `#F0F3F7` 背景，36pt 数值 |
| 封面/总结 | `#1E2A38` 深色背景 |

详见 `references/design-system.md`。

## 文件结构

```
report-to-ppt/
├── README.md
├── SKILL.md                      # Skill 定义（Claude 的指令）
├── scripts/
│   └── build_ppt.py              # 核心脚本（outline / auto / manual 三模式）
└── references/
    ├── design-system.md          # 完整视觉规范
    └── pitfalls.md              # 常见坑与解决方案
```
