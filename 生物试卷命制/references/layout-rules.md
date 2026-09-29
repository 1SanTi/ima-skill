# 试卷排版规范（打印友好）

目标：**图文清晰、版面紧凑、便于打印**（全卷通常控制在 8 页以内）。

## 一、页面与字号

- 纸张 A4（21×29.7 cm），页边距上下约 1.4 cm、左右约 1.6 cm。
- 正文（题干、选项）约 10 pt；题注（图/表名）约 7.5~8 pt；小字标注（素养/难度）8 pt、浅灰。
- 行距 1.0~1.15，段前后间距 0.5~2 pt，避免大空隙。

## 二、选择题：左文右图

- 每题用一个**无边框 1×2 表格**：左列放题干＋选项（约 12.9 cm），右列放配图＋图注（约 4.9 cm）。
- 配图宽度控制在 **≤4.6 cm**（高度同期限 ≤4.2 cm），既不挤压文字也不越界。
- 无配图的题目直接用普通段落，保持紧凑。
- 选项：短选项两列（制表位 ~6 cm），长选项每行一个。

## 三、非选择题

- 单图居中，宽 ≤6.2 cm；多图用 1×n 无边框表格并排（每图宽 ≤列宽−0.3 cm）。
- 数据表用真实 Word 表格重建（比截图清晰），表下加小字题注。
- 作答区：填空用下划线字符；需写句子处留足行距（≥1.35）。

## 四、避免"断题/孤行"

- 题号与标题段落设置 `keep_with_next = True`，防止标题孤悬页尾。

## 五、页数与越界自检（必做）

1. 转 PDF 数页数：
   `soffice --headless --convert-to pdf --outdir <dir> <docx>`
2. 数页数并检查图片是否越界（A4 页宽 595.3 pt，左右边距 1.6 cm 时内容右界 ≈ 549.9 pt）：

```python
import pymupdf
d = pymupdf.open(pdf)
print("页数:", d.page_count)
ML, MR = 45.35, 549.95
for i in range(d.page_count):
    for im in d[i].get_image_info():
        b = im["bbox"]
        if b[0] < ML-1 or b[2] > MR+1:
            print("越界 page", i+1, b)
```

## 六、图片按框缩放的正确算法（易错）

给图片定尺寸时应"等比缩放进 (max_w, max_h) 的框"：

```python
pw, ph = Image.open(png).size
r = pw / ph
if r >= max_w / max_h:      # 宽高比大 → 宽度受限
    w, h = max_w, max_w / r
else:                        # 宽高比小 → 高度受限
    w, h = max_h * r, max_h
p.add_run().add_picture(png, width=Cm(w), height=Cm(h))
```

⚠️ 该条件**极易写反**。写反会把图片放大到超出页边距（本项目曾出现 21 张图有 12 张越界）。务必用上面的 PDF 越界自检复核。
