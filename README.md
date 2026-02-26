# Excel 单词拆分工具

用于把 Excel 第 2 列（B 列）中“一个单元格包含多个单词条目”的内容拆成 **每个单词一行**。

输出为 3 列：
1. 第1列（复制原 A 列）
2. 单词（从每条记录行首提取）
3. 释义（词性、释义、用法等剩余部分）

## 用法

```bash
python split_vocab_xlsx.py dict.xlsx -o dict_result.csv
```

也可以在当前目录只有一个 `.xlsx` 文件时省略输入文件名：

```bash
python split_vocab_xlsx.py -o dict_result.csv
```

如果你想先验证解析逻辑：

```bash
python split_vocab_xlsx.py --self-test
```

## 说明

- 脚本按 `.xlsx` 的 XML 结构读取数据，不依赖第三方库（适合离线环境）。
- 默认读取第一个工作表。
- B 列每个单元格内支持多行文本；当某行看起来是“新单词开头”时会启动新记录；否则视为上一条的续行。
- 输出为 UTF-8 BOM 的 CSV，直接用 Excel 打开一般不会乱码。


## 常见问题（GitHub 文件读不到）
