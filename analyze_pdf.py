#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""PDFテーブル構造を詳細に解析するスクリプト"""

import pdfplumber

pdf = pdfplumber.open('審査基準表.pdf')
tables = pdf.pages[0].extract_tables()

print(f"Tables found: {len(tables)}")
print(f"Table 0: {len(tables[0])} rows")

# 全列数を確認
if tables[0]:
    print(f"Columns in first row: {len(tables[0][0])}")

print("\n=== 全行データ ===")
for i, row in enumerate(tables[0]):
    print(f"Row {i:3d}: {row}")
