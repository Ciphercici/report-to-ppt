# report-to-ppt

把内容报告（Markdown）转成专业可编辑的 PowerPoint 演示文稿。

## 安装

### 方式一：从 GitHub Release 下载

1. 在 [Releases](https://github.com/<user>/report-to-ppt/releases) 页面下载 `report-to-ppt.skill`
2. 解压到 Claude Code 的 skills 目录：

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
git clone https://github.com/<user>/report-to-ppt.git ~/.claude/skills/report-to-ppt/
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

## 工具链

| 工具 | 角色 |
|------|------|
| `build_ppt.py` | 内容排版 → 生成初稿 PPTX |
| `officecli` | 后期微调：改文字、调颜色、校验质量 |

### 安装 officecli

```bash
# macOS / Linux
curl -fsSL https://d.officecli.ai/install.sh | bash

# Windows (PowerShell)
irm https://d.officecli.ai/install.ps1 | iex
```

## 工作流

```
原始报告 .md
     │
     ▼  Claude 读懂内容，设计叙事结构
.slide.md 大纲（Marp 格式，12-18 页）
     │
     ▼  build_ppt.py 自动排版
初稿 .pptx
     │
     ▼  officecli 微调 + 校验
终稿 .pptx（文字可编辑，设计专业）
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

### 第二步：生成初稿

```bash
python scripts/build_ppt.py outline.slide.md -o output.pptx
```

### 第三步：officecli 微调

生成后的 PPTX 可以直接用自然语言指令修改，**不需要写 Python 代码**：

```bash
# 查看结构
officecli view output.pptx outline

# 改文字
officecli set output.pptx '/slide[2]/shape[2]' --prop text="修正后的标题"

# 调颜色
officecli set output.pptx '/slide[1]' --prop fill=1E2A38

# 全局替换
officecli set output.pptx / --find "旧文字" --replace "新文字"

# 改字号
officecli set output.pptx '/slide[3]/shape[2]' --prop font.size=18pt

# 删元素
officecli remove output.pptx '/slide[4]/shape[3]'

# 添加新内容
officecli add output.pptx '/slide[5]' --type shape --prop text="补充要点" --prop x=2cm --prop y=6cm
```

本质是用自然语言 → officecli 命令 → 直接操作 PPTX DOM，省去手工写 Python 再重新生成的循环。

### 第四步：验证

```bash
officecli view output.pptx stats          # 页数、形状统计
officecli view output.pptx issues         # 自动检测排版问题
```

正常输出：12-18 页，文件 > 1 MB，零 issue。

## 三种生成模式

| 模式 | 命令 | 质量 | 适用 |
|------|------|------|------|
| **Outline**（推荐） | `build_ppt.py report.slide.md` | 高 | 最终交付 |
| **Auto**（快速草稿） | `build_ppt.py report.md` | 中 | 快速看结构 |
| **Outline + officecli**（终极） | Outline 生成后用 officecli 逐页精调 | 最高 | 要最完美效果 |

### Outline 模式

Claude 设计 `.slide.md` 大纲 → 脚本排版。内容提炼由 Claude 完成，排版由脚本完成。

### Auto 模式

脚本直接解析原始报告 md，自动拆章节、提取图片、检测 KPI、分页。速度快但要点质量取决于原文写法。生成后可用 officecli 清理。

### Outline + officecli 精调

Outline 模式生成初稿后，用 officecli 逐页精细调整。这是"终极模式"——Claude 做内容、脚本做排版、officecli 做微调，各司其职。

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
├── SKILL.md
├── report-to-ppt.skill
├── assets/
├── scripts/
│   └── build_ppt.py              # PPT 生成脚本（outline / auto 双模式）
└── references/
    ├── design-system.md          # 完整视觉规范
    └── pitfalls.md               # 常见坑与解决方案
```

## 依赖

```bash
pip install python-pptx
```

以及 [officecli](https://officecli.ai)（用于第三步微调和第四步验证）：

```bash
# macOS / Linux
curl -fsSL https://d.officecli.ai/install.sh | bash

# Windows (PowerShell)
irm https://d.officecli.ai/install.ps1 | iex
```

