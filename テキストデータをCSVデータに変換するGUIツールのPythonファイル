#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
テキストデータをCSVデータに変換するGUIツール
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import csv
import re
import os
from pathlib import Path
from typing import List, Dict, Any, Optional

class TextToCSVConverterGUI:
    """テキストデータをCSVに変換するGUIアプリケーション"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("テキスト→CSV変換ツール")
        self.root.geometry("900x700")
        
        # 変数の初期化
        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.parse_mode = tk.StringVar(value="auto")
        self.delimiter = tk.StringVar()
        self.pattern = tk.StringVar()
        self.widths = tk.StringVar()
        self.headers = tk.StringVar()
        self.encoding = tk.StringVar(value="utf-8")
        
        self.preview_data = []
        self.setup_ui()
    
    def setup_ui(self):
        """UIセットアップ"""
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # ファイル選択エリア
        self.create_file_selection_area(main_frame)
        
        # 設定エリア
        self.create_settings_area(main_frame)
        
        # プレビューエリア
        self.create_preview_area(main_frame)
        
        # ボタンエリア
        self.create_button_area(main_frame)
        
        # ステータスバー
        self.create_status_area(main_frame)
        
        # グリッドの重み設定
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
    
    def create_file_selection_area(self, parent):
        """ファイル選択エリア作成"""
        file_frame = ttk.LabelFrame(parent, text="ファイル選択", padding="5")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        # 入力ファイル
        ttk.Label(file_frame, text="入力ファイル:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        ttk.Entry(file_frame, textvariable=self.input_file_path, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(file_frame, text="参照", command=self.browse_input_file).grid(row=0, column=2)
        
        # 出力ファイル
        ttk.Label(file_frame, text="出力ファイル:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        ttk.Entry(file_frame, textvariable=self.output_file_path, width=50).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 5), pady=(5, 0))
        ttk.Button(file_frame, text="参照", command=self.browse_output_file).grid(row=1, column=2, pady=(5, 0))
    
    def create_settings_area(self, parent):
        """設定エリア作成"""
        settings_frame = ttk.LabelFrame(parent, text="変換設定", padding="5")
        settings_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        settings_frame.columnconfigure(1, weight=1)
        
        # 解析モード
        ttk.Label(settings_frame, text="解析モード:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        mode_combo = ttk.Combobox(settings_frame, textvariable=self.parse_mode, width=15)
        mode_combo['values'] = ('auto', 'structured', 'key_value', 'pattern', 'fixed_width')
        mode_combo.grid(row=0, column=1, sticky=tk.W, padx=(0, 10))
        mode_combo.bind('<<ComboboxSelected>>', self.on_mode_changed)
        
        # エンコーディング
        ttk.Label(settings_frame, text="エンコーディング:").grid(row=0, column=2, sticky=tk.W, padx=(10, 5))
        encoding_combo = ttk.Combobox(settings_frame, textvariable=self.encoding, width=10)
        encoding_combo['values'] = ('utf-8', 'shift_jis', 'cp932', 'euc-jp')
        encoding_combo.grid(row=0, column=3, sticky=tk.W)
        
        # 区切り文字（structuredモード用）
        self.delimiter_label = ttk.Label(settings_frame, text="区切り文字:")
        self.delimiter_entry = ttk.Entry(settings_frame, textvariable=self.delimiter, width=10)
        
        # 正規表現パターン（patternモード用）
        self.pattern_label = ttk.Label(settings_frame, text="正規表現:")
        self.pattern_entry = ttk.Entry(settings_frame, textvariable=self.pattern, width=30)
        
        # 列幅（fixed_widthモード用）
        self.widths_label = ttk.Label(settings_frame, text="列幅(カンマ区切り):")
        self.widths_entry = ttk.Entry(settings_frame, textvariable=self.widths, width=20)
        
        # ヘッダー
        ttk.Label(settings_frame, text="ヘッダー(カンマ区切り):").grid(row=2, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        ttk.Entry(settings_frame, textvariable=self.headers, width=40).grid(row=2, column=1, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # 初期状態では追加設定を非表示
        self.update_mode_specific_controls()
    
    def create_preview_area(self, parent):
        """プレビューエリア作成"""
        preview_frame = ttk.LabelFrame(parent, text="プレビュー", padding="5")
        preview_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(1, weight=1)
        
        # プレビューボタン
        ttk.Button(preview_frame, text="プレビュー更新", command=self.update_preview).grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        # プレビューテキスト
        self.preview_text = scrolledtext.ScrolledText(preview_frame, height=15, width=80)
        self.preview_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    
    def create_button_area(self, parent):
        """ボタンエリア作成"""
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=3, column=0, columnspan=2, pady=(0, 10))
        
        ttk.Button(button_frame, text="変換実行", command=self.convert_file, style="Accent.TButton").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="クリア", command=self.clear_all).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="サンプル読込", command=self.load_sample).pack(side=tk.LEFT)
    
    def create_status_area(self, parent):
        """ステータスエリア作成"""
        self.status_var = tk.StringVar(value="準備完了")
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E))
        status_frame.columnconfigure(0, weight=1)
        
        ttk.Label(status_frame, textvariable=self.status_var, relief=tk.SUNKEN).grid(row=0, column=0, sticky=(tk.W, tk.E))
    
    def browse_input_file(self):
        """入力ファイル選択"""
        filename = filedialog.askopenfilename(
            title="入力テキストファイルを選択",
            filetypes=[("テキストファイル", "*.txt"), ("すべてのファイル", "*.*")]
        )
        if filename:
            self.input_file_path.set(filename)
            # 出力ファイル名を自動生成
            if not self.output_file_path.get():
                base_name = Path(filename).stem
                output_path = Path(filename).parent / f"{base_name}_converted.csv"
                self.output_file_path.set(str(output_path))
    
    def browse_output_file(self):
        """出力ファイル選択"""
        filename = filedialog.asksaveasfilename(
            title="出力CSVファイルを選択",
            defaultextension=".csv",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
        )
        if filename:
            self.output_file_path.set(filename)
    
    def on_mode_changed(self, event=None):
        """解析モード変更時の処理"""
        self.update_mode_specific_controls()
    
    def update_mode_specific_controls(self):
        """モード固有のコントロール表示/非表示"""
        # 既存のコントロールを削除
        for widget in [self.delimiter_label, self.delimiter_entry, 
                      self.pattern_label, self.pattern_entry,
                      self.widths_label, self.widths_entry]:
            widget.grid_remove()
        
        mode = self.parse_mode.get()
        settings_frame = self.delimiter_label.master
        
        if mode == 'structured':
            self.delimiter_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
            self.delimiter_entry.grid(row=1, column=1, sticky=tk.W, pady=(5, 0))
        elif mode == 'pattern':
            self.pattern_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
            self.pattern_entry.grid(row=1, column=1, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))
        elif mode == 'fixed_width':
            self.widths_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
            self.widths_entry.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))
    
    def parse_structured_text(self, text: str, delimiter: str = None) -> List[List[str]]:
        """構造化されたテキストを解析"""
        lines = text.strip().split('\n')
        data = []
        
        for line in lines:
            if not line.strip():
                continue
                
            if delimiter:
                row = [cell.strip() for cell in line.split(delimiter)]
            else:
                if '\t' in line:
                    row = [cell.strip() for cell in line.split('\t')]
                elif ',' in line and line.count(',') >= 2:
                    row = [cell.strip() for cell in line.split(',')]
                else:
                    row = re.split(r'\s{2,}', line.strip())
                    if len(row) == 1:
                        row = line.split()
            
            data.append(row)
        
        return data
    
    def parse_key_value_text(self, text: str) -> List[List[str]]:
        """キー：値形式のテキストを解析"""
        lines = text.strip().split('\n')
        data = []
        
        for line in lines:
            if not line.strip():
                continue
            
            separators = ['：', ':', '=', '\t']
            for sep in separators:
                if sep in line:
                    parts = line.split(sep, 1)
                    if len(parts) == 2:
                        key = parts[0].strip()
                        value = parts[1].strip()
                        data.append([key, value])
                    break
        
        return data
    
    def parse_pattern_text(self, text: str, pattern: str) -> List[List[str]]:
        """正規表現パターンを使用してテキストを解析"""
        lines = text.strip().split('\n')
        data = []
        
        try:
            regex = re.compile(pattern)
            for line in lines:
                match = regex.search(line)
                if match:
                    groups = match.groups()
                    data.append(list(groups))
        except re.error as e:
            raise ValueError(f"正規表現エラー: {e}")
        
        return data
    
    def parse_fixed_width_text(self, text: str, widths: List[int]) -> List[List[str]]:
        """固定幅テキストを解析"""
        lines = text.strip().split('\n')
        data = []
        
        for line in lines:
            if not line.strip():
                continue
            
            row = []
            start = 0
            for width in widths:
                end = start + width
                cell = line[start:end].strip() if end <= len(line) else line[start:].strip()
                row.append(cell)
                start = end
            
            data.append(row)
        
        return data
    
    def update_preview(self):
        """プレビュー更新"""
        if not self.input_file_path.get():
            messagebox.showwarning("警告", "入力ファイルを選択してください。")
            return
        
        try:
            self.status_var.set("プレビュー生成中...")
            self.root.update()
            
            # ファイル読み込み
            with open(self.input_file_path.get(), 'r', encoding=self.encoding.get()) as f:
                text = f.read()
            
            # テキスト解析
            mode = self.parse_mode.get()
            
            if mode == 'structured' or mode == 'auto':
                data = self.parse_structured_text(text, self.delimiter.get() or None)
            elif mode == 'key_value':
                data = self.parse_key_value_text(text)
            elif mode == 'pattern':
                if not self.pattern.get():
                    messagebox.showerror("エラー", "正規表現パターンを入力してください。")
                    return
                data = self.parse_pattern_text(text, self.pattern.get())
            elif mode == 'fixed_width':
                if not self.widths.get():
                    messagebox.showerror("エラー", "列幅を入力してください。")
                    return
                try:
                    widths = [int(w.strip()) for w in self.widths.get().split(',')]
                    data = self.parse_fixed_width_text(text, widths)
                except ValueError:
                    messagebox.showerror("エラー", "列幅は整数をカンマ区切りで入力してください。")
                    return
            
            # プレビュー表示
            self.preview_data = data
            self.display_preview(data)
            self.status_var.set(f"プレビュー完了: {len(data)}行のデータを解析しました。")
            
        except FileNotFoundError:
            messagebox.showerror("エラー", f"ファイル '{self.input_file_path.get()}' が見つかりません。")
        except UnicodeDecodeError:
            messagebox.showerror("エラー", "ファイルのエンコーディングが正しくありません。")
        except Exception as e:
            messagebox.showerror("エラー", f"プレビュー生成中にエラーが発生しました: {e}")
            self.status_var.set("エラーが発生しました。")
    
    def display_preview(self, data: List[List[str]]):
        """プレビューデータを表示"""
        self.preview_text.delete(1.0, tk.END)
        
        if not data:
            self.preview_text.insert(tk.END, "解析可能なデータが見つかりませんでした。")
            return
        
        # ヘッダー行
        headers = None
        if self.headers.get():
            headers = [h.strip() for h in self.headers.get().split(',')]
            header_line = ','.join(headers) + '\n'
            self.preview_text.insert(tk.END, f"[ヘッダー] {header_line}")
        
        # データ行（最初の10行まで）
        max_preview_rows = 10
        for i, row in enumerate(data[:max_preview_rows]):
            csv_line = ','.join([f'"{cell}"' if ',' in str(cell) else str(cell) for cell in row])
            self.preview_text.insert(tk.END, f"[{i+1:3d}行目] {csv_line}\n")
        
        if len(data) > max_preview_rows:
            self.preview_text.insert(tk.END, f"\n... 他 {len(data) - max_preview_rows} 行")
    
    def convert_file(self):
        """ファイル変換実行"""
        if not self.input_file_path.get():
            messagebox.showwarning("警告", "入力ファイルを選択してください。")
            return
        
        if not self.output_file_path.get():
            messagebox.showwarning("警告", "出力ファイルを指定してください。")
            return
        
        try:
            self.status_var.set("変換中...")
            self.root.update()
            
            # プレビューデータがある場合はそれを使用、なければ解析実行
            if not self.preview_data:
                self.update_preview()
            
            data = self.preview_data
            if not data:
                return
            
            # CSV出力
            with open(self.output_file_path.get(), 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                
                # ヘッダー行を書き込み
                if self.headers.get():
                    headers = [h.strip() for h in self.headers.get().split(',')]
                    writer.writerow(headers)
                
                # データ行を書き込み
                for row in data:
                    writer.writerow(row)
            
            self.status_var.set(f"変換完了: {len(data)}行のデータを出力しました。")
            messagebox.showinfo("完了", f"変換が完了しました。\n出力ファイル: {self.output_file_path.get()}")
            
        except Exception as e:
            messagebox.showerror("エラー", f"変換中にエラーが発生しました: {e}")
            self.status_var.set("エラーが発生しました。")
    
    def clear_all(self):
        """全てクリア"""
        self.input_file_path.set("")
        self.output_file_path.set("")
        self.delimiter.set("")
        self.pattern.set("")
        self.widths.set("")
        self.headers.set("")
        self.preview_text.delete(1.0, tk.END)
        self.preview_data = []
        self.status_var.set("クリアしました。")
    
    def load_sample(self):
        """サンプルデータ読込"""
        # サンプルファイルを作成
        sample_text = """名前	年齢	職業	部署
田中太郎	30	エンジニア	開発部
佐藤花子	25	デザイナー	企画部
山田次郎	35	営業	営業部
鈴木美咲	28	マーケター	マーケティング部"""
        
        sample_file = "sample_data.txt"
        with open(sample_file, 'w', encoding='utf-8') as f:
            f.write(sample_text)
        
        self.input_file_path.set(sample_file)
        self.output_file_path.set("sample_output.csv")
        self.parse_mode.set("structured")
        self.headers.set("名前,年齢,職業,部署")
        
        self.update_preview()
        self.status_var.set("サンプルデータを読み込みました。")

def main():
    """メイン関数"""
    root = tk.Tk()
    app = TextToCSVConverterGUI(root)
    
    # スタイル設定
    style = ttk.Style()
    if "Accent.TButton" not in style.theme_names():
        style.configure("Accent.TButton", foreground="white", background="blue")
    
    root.mainloop()

if __name__ == '__main__':
    main()
