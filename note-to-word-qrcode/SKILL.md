---
name: note-to-word-qrcode
description: 把 IMA 笔记（尤其是政策汇编/资料汇编类，含多个文件条目）按文件分节整理成一个大 Word 文档，并生成可用于扫码下载该 Word 的二维码。当用户说“把这篇笔记按文件整理成 Word”“笔记转 Word 并生成二维码”“扫码下载这个文档”“把笔记整理成 Word 再给个二维码”时触发。不适用于：单文件笔记的普通导出（用 note-to-docx）、仅生成二维码（用 link-to-qrcode）、知识库文件处理。
---

# 笔记 → 分文件 Word → 扫码下载二维码

把一篇 IMA 笔记按“文件”分节整理成一个大 Word（封面 + 文件清单 + 每文件独立一章 + 参考文献/链接验证等尾部），再把 Word 的下载链接生成二维码交付给用户。

## 触发场景

- 用户给出一篇笔记（或 note_id / 笔记标题），要求“按文件分别整理好，输出一个比较大的 Word”
- 用户要求“给 Word 文档配一个二维码，扫码下载”
- 典型对象：政策汇编、文件汇编、资料汇编类笔记（正文由多个 `序号. 《文件名》：链接` 条目组成）

## 工作流

### 1. 取笔记原文（Markdown）
优先用 `ima-note` skill 的 `export_note` 接口导出为 Markdown（Fetch 返回的是生成内容，非原文时用它兜底）：

```bash
curl -s -X POST "https://ima.qq.com/openapi/note/v1/export_note" \
  -H "ima-openapi-clientid: $IMA_OPENAPI_CLIENTID" -H "ima-openapi-apikey: $IMA_OPENAPI_APIKEY" \
  -H "Content-Type: application/json" \
  -d '{"note_id":"<NOTE_ID>","target_content_format":1}'
# 取返回的 data.content_url，再 curl 下载成 note_raw.md
```

若用户只给了标题，先用 `search_note` 按标题定位拿 note_id。

### 2. Markdown → 分文件 Word

```bash
python3 scripts/note_md_to_docx.py <input.md> <output.docx>
```

脚本做什么：
- 文档标题取第一条 `# 标题`；
- 用启发式识别“文件条目行”（行首 `数字.` + 含 `《…》` / 含 Markdown 链接 / 以“通知/意见/方案/规划/纲要/办法/条例/指南/决定/计划/报告/措施”结尾）作为分章边界；
- 每个文件一章（Heading 1），其下 `来源：<超链接>`；正文按 `**加粗**`、`一、`、`（一）`、`# ` 等还原为多级标题与段落；
- 被压成一整段的正文按 `一、/（一）/1.` 自动切分，恢复可读性；
- 结尾的“参考文献”“链接验证结果”等 `##` 章节单独成章；
- 封面 + “收录文件清单”，每章之间分页。

可用参数/要点：
- 若误判章边界或漏判，检查 `is_file_heading()` 的关键词表与阈值；
- 生成后建议校验：`python3 -c "from docx import Document; d=Document('out.docx'); print(sum(1 for p in d.paragraphs if p.style.name=='Heading 1'))"`，H1 数应约等于“文件数 + 1(清单) + 尾部章节数”。

### 3. 取 Word 下载链接
用 `provide_file` 工具把生成的 .docx 交给用户，得到 `download_url`（约 90 天有效）。

### 4. 生成二维码

```bash
python3 scripts/make_qr.py --url "<download_url>" \
  --out qr.png --title "文档标题" --subtitle "扫码下载Word文档"
```

- 默认 H 级纠错 + 中文标题/副标题 + 回读校验（校验对纯二维码区域解码，避免标题文字干扰）；
- 扫码不佳时：加大 `--box-size`（如 14~16）或把纠错降为 `--ec Q`（降低码密度）；
- 批量：`--file urls.txt --outdir qrdir`。

### 5. 交付
用 `provide_file` 把二维码 PNG 交给用户（沙箱文件用户看不到，必须 provide_file 才会显示）。

## 注意
- 下载链接是临时链接（约 90 天），二维码内容即该链接，过期需重新生成。
- 笔记内嵌的 PDF/图片附件不会自动并入 Word，脚本会在原位留“【本文件为 PDF 附件】”提示；如需并入，另用 media 导出流程。
- 依赖：python-docx、qrcode[pil]、Pillow、可选 opencv-python（回读校验）。缺 qrcode 时脚本会自动 pip 安装。
