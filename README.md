# 中国企业温室气体排放因子库网页平台

静态网页平台，展示 Excel 中的全部因子数据、基本说明、导航目录、Q&A 和参考文献。

## 更新数据

```bash
python3 scripts/extract_workbook.py
```

如果源 Excel 路径变化：

```bash
SOURCE_XLSX=/path/to/file.xlsx python3 scripts/extract_workbook.py
```

## 本地访问

```bash
python3 -m http.server 5174
```

打开 `http://localhost:5174/`。
