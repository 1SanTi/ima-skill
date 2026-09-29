# IMA 技能合集 · ima-skill

> 刘东安（[@1SanTi](https://github.com/1SanTi)）自建的 IMA（ima.copilot）技能合集，开源共享。

共 **14** 个技能，每个技能一个独立文件夹，内含 `SKILL.md`（技能定义）及配套 `scripts/`、`references/`、`assets/`。

## 技能一览

### 🎓 教学与命题

| 技能 | 功能简介 |
| --- | --- |
| [`语文试卷命制`](./语文试卷命制) | 语文试卷命制助手 |
| [`生物试卷命制`](./生物试卷命制) | 初中生物学试卷命制助手 |
| [`teaching-package-generator`](./teaching-package-generator) | 教学授课资料一体化生成（AI课程包生成器） |

### 📄 文档与文稿

| 技能 | 功能简介 |
| --- | --- |
| [`docx-format-styler`](./docx-format-styler) | 把一个 Word 文档的字体字号格式套用到另一个文档（"照着 A 改 B 的格式"） |
| [`note-to-word-qrcode`](./note-to-word-qrcode) | 把 IMA 笔记（尤其是政策汇编/资料汇编类，含多个文件条目）按文件分节整理成一个大 Word 文档，并生成可用于扫码下载该 Word 的二… |
| [`研修总结撰写`](./研修总结撰写) | 依据会议/活动/培训的笔记（含录音转写）撰写「研修总结」「培训总结」「学习心得体会」并生成规范排版的 Word 文档 |
| [`campus-news-writer`](./campus-news-writer) | 校园新闻稿撰写与知识库存档 |

### 🗂️ 教育申报与材料

| 技能 | 功能简介 |
| --- | --- |
| [`职称申报表格按意见修改`](./职称申报表格按意见修改) | 教师职称申报表格（公示表、申报表等）的按意见就地修订与归档 |
| [`edu-award-material-revision`](./edu-award-material-revision) | 教育系统先进人物（优秀教师/优秀班主任/优秀乡村教师/优秀教育工作者等）申报材料的按需修改与归档 |
| [`edu-advanced-figure-application-form`](./edu-advanced-figure-application-form) | 教育系统先进人物（优秀教师/优秀班主任/优秀教育工作者/优秀教研团队）申报表自动填写 |
| [`gongfei-entry-form-filling`](./gongfei-entry-form-filling) | 公费师范生/新教师“入编（入职）表格”代填与打包 |

### 🏫 学校事务

| 技能 | 功能简介 |
| --- | --- |
| [`club-activity-scheme`](./club-activity-scheme) | 依据体卫艺开学工作检查表，批量生成学校社团活动方案、全员文体活动计划及实施方案、各社团活动计划与学生名单（公文格式Word文档）并分类存入知… |

### 🛠️ 通用工具

| 技能 | 功能简介 |
| --- | --- |
| [`link-to-qrcode`](./link-to-qrcode) | 链接/文本转二维码生成器 |
| [`survey-design-collector`](./survey-design-collector) | 问卷设计与在线收集系统生成器 |

## 使用方式

把需要的技能文件夹放入 IMA 技能目录（本环境为 `/root/.skills/`）后即可被识别调用；`SKILL.md` 头部含 `name` / `description` 等元信息。

## 许可

本项目采用 [MIT License](./LICENSE) 开源。
