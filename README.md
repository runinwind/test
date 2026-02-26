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

## 解析规则（当前版本）

- 支持以下常见词条头格式：
  - `word n./v./adj. ...`
  - `word 中文释义...`（如 `participate 参与`）
  - `word: 释义...`（中英文冒号都支持）
  - 多词短语词头（如 `immune system 免疫系统`）
- 缩进行默认按“续行”处理（用于例句/补充说明）。
- 但缩进行如果明显是新词条（如 `  proper 适当的`、`  symbol n. ...`）也会拆分成新词，避免漏词。
- 一行里多个紧凑词头会拆开（如 `cater, crater ...`）。
- 释义和例句中的换行会保留在同一词条内，不会丢失。

## 说明

- 脚本按 `.xlsx` 的 XML 结构读取数据，不依赖第三方库（适合离线环境）。
- 默认读取第一个工作表。
- 输出为 UTF-8 BOM 的 CSV，直接用 Excel 打开一般不会乱码。
