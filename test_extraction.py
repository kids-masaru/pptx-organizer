#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
階層構造抽出のテストスクリプト - ファイル出力版
"""

from main import extract_categories_from_pdf

def test_extraction():
    with open('extraction_test_result.txt', 'w', encoding='utf-8') as f:
        f.write("=== PDF階層構造抽出テスト ===\n\n")
        categories = extract_categories_from_pdf('審査基準表.pdf')
        
        f.write(f"抽出件数: {len(categories)} カテゴリ\n\n")
        
        for cat in categories:
            f.write(f"No.{cat['No']}: {cat.get('MainCategory', '')}\n")
            for sub in cat.get('SubItems', []):
                f.write(f"   - {sub}\n")
            f.write("\n")
        
        f.write("=== テスト完了 ===\n")
    
    print("結果をextraction_test_result.txtに保存しました")

if __name__ == '__main__':
    test_extraction()
