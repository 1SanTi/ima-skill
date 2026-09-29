---
name: survey-design-collector
description: 问卷设计与在线收集系统生成器。把一份问卷需求（录音口述、文字大纲、散乱题目）一键变成「Word 问卷 + 手机扫码答题网站 + 后台数据自动汇总看板」三件套，并产出可一键运行的完整系统。当用户说“设计一份问卷”“根据录音设计问卷”“做个调查问卷”“生成问卷”“在线问卷”“扫码答题的网站”“问卷数据汇总”“问卷后台”“问卷统计”“把问卷做成网页/网站”“老师们扫码填问卷”时使用。也适用于把已有题目整理成规范问卷或在线收集系统。不适用于非问卷类文档生成、单纯的问卷统计数据分析（无新问卷设计）、考试组卷（用 question-generator）。
version: 1.0.0
tags:
  - 问卷
  - 调查问卷
  - 在线问卷
  - 扫码答题
  - 数据汇总
  - 问卷后台
---

# 问卷设计与在线收集系统生成器

把「一份问卷需求」变成可直接使用的**三件套**：

1. **Word 问卷**（规范排版，可打印、可分发）
2. **在线答题网站**（手机扫码即答，零依赖，双击即用）
3. **后台数据汇总看板**（各题自动统计图表 + 明细 + 一键导出 Excel + 答题二维码）

全流程零第三方依赖（仅 Python 标准库），产物可拷到任意电脑直接运行。

---

## 何时使用

- 用户给出一段**录音/口述/文字**，要求“设计一份问卷”“根据录音设计问卷”。
- 用户已有题目，想做成**在线问卷 / 扫码答题网站**并**自动汇总数据**。
- 用户要一个“调查问卷 + 后台统计”的完整方案。

---

## 执行流程（SOP）

### 第 1 步　获取并理解需求
- **录音**：先用语音转文字把录音转成文字（可用 `speech-to-text` 技能或直接 `fetch` 音频文件），
  提炼出：问卷主题、面向对象、有几个板块（维度）、每个维度下的具体题目、题型偏好（单选/多选/量表/填空）。
- **文字**：直接读用户给的大纲。
- 不确定的地方（如是否加“基本信息”板块、面向哪个单位）**要么按惯例补齐并在交付时说明，要么向用户确认**。

### 第 2 步　编写问卷规格 `questions.json`
严格按 `references/spec-format.md` 的格式，把问卷写成 JSON：
- 顶层：`title`（标题）、`subtitle`、`intro`（卷首语）、`footer`、`scale_options`（量表五个档位）、`sections`。
- 每个 section 下 `questions[]`，每题含 `id`、`type`、`title`、`required`，选项题含 `options`。
- 题型：`single` 单选｜`multi` 多选｜`scale` 量表（选项取 `scale_options`）｜`text` 填空。
- `allow_other: true` 自动加“其他（请注明）”；`exclusive: ["都没用过"]` 表示该选项与其他互斥。
- **建议**：正式问卷一般先加一个精简的“基本信息”板块（学段 / 学科 / 教龄），便于后续交叉分析。

> 设计要点：题干用陈述式、避免诱导；选项穷尽且互斥；量表统一档位；题目数量控制在 3—5 分钟可完成。

### 第 3 步　一键生成三件套
```bash
python3 scripts/build_survey.py --spec questions.json --out <输出目录>
# 若已知最终答题网址，可加 --url http://xxx/ 直接生成二维码
```
产物：
```
<输出目录>/
├── <问卷标题>.docx          规范排版的 Word 问卷
└── 在线问卷系统/
    ├── server.py            答题页 + 后台 + 接口（零依赖）
    ├── questions.json       本问卷题目
    ├── 启动_Windows.bat / 启动_Mac_Linux.sh   一键启动
    ├── make_qr.py           单独生成二维码
    ├── vendor/qrcode/       内置二维码生成
    └── 使用说明.md
```

### 第 4 步　自检（必做）
- 打开生成的 `题库/在线问卷系统/questions.json`，确认题目、选项、必答标记无误。
- 启动一次系统做冒烟测试：
  ```bash
  cd 在线问卷系统 && PORT=8123 python3 server.py &
  curl -s http://127.0.0.1:8123/ | head            # 答题页可打开
  curl -s -X POST http://127.0.0.1:8123/submit --data-urlencode "q1=..." # 提交入库
  curl -s http://127.0.0.1:8123/api/stats | head   # 统计正常
  ```
- 确认 Word 里每道题**题干与选项未被分页切开**（脚本已设置 keep-with-next）。

### 第 5 步　交付与部署说明
向用户说明：
- 三件套各自怎么用；二维码在**后台页面**自动显示（`http://电脑IP:8000/admin`）。
- **老师怎么访问**：同一 WiFi 下扫码即答（校内最常用）；若要公网访问，需放到有公网 IP 的服务器上。
- 详见 `references/deploy-guide.md`，可直接复用其中文案。
- 可用 `provide_file` 把 Word 问卷、以及打包后的「在线问卷系统」压缩包交付给用户下载。

### 第 6 步（可选）入库归档
如需归档，用 `upload-to-kb` 技能把 Word 问卷与系统压缩包存入用户指定知识库。

---

## 关键要点与经验

- **零依赖是核心竞争力**：只用 Python 标准库 + 内置 `vendor/qrcode`，用户无需 pip 安装任何东西；
  Word 生成才需要 `python-docx`（生成端环境需有）。
- **答题页用纯 HTML 表单 + 少量 JS**：提交走 `POST /submit`，服务端校验必答项（缺项会带着已填内容回显并标红）。
- **二维码在服务端用 SVG 实时生成**（`/api/qr.svg?url=...`），后台页面按当前访问网址自动出码，
  老师换网络/改网址后重新出码即可，无需重打包。
- **数据落盘在 `data/responses.db`**（SQLite），备份=复制该文件，清空=删该文件。
- 局域网访问打不开，90% 是**电脑防火墙**未放行 Python，提示用户允许即可。
- 端口被占用可用 `PORT=8080 python3 server.py` 换端口。

---

## 目录结构

```
survey-design-collector/
├── SKILL.md
├── scripts/
│   ├── build_survey.py    一键生成（读 spec → Word + 网站 + 二维码）
│   ├── spec_to_word.py    由 questions.json 生成 Word 问卷
│   └── make_qr.py         由网址生成二维码 PNG
├── assets/app/            在线问卷系统模板（server.py / 启动脚本 / vendor/qrcode）
└── references/
    ├── spec-format.md     questions.json 规格说明 + 录音转问卷方法
    └── deploy-guide.md    部署与对外说明模板
```
