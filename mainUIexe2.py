import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
import re
import os
import sys


def resource_path(relative_path):
    """ Get the absolute path to the resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# フォント設定
FONT_CONFIG = {
    'ja': {'name': 'IPAexGothic', 'file': resource_path('ipaexg.ttf')},
    'ko': [

        {'name': 'NanumGothic', 'file': resource_path('NanumGothic.ttf')}
    ]
}


def register_fonts():
    """フォント登録関数"""
    # 日本語フォント登録（リソースパスを利用）
    try:
        font_path = resource_path(FONT_CONFIG['ja']['file'])
        pdfmetrics.registerFont(TTFont(FONT_CONFIG['ja']['name'], font_path))
    except Exception as e:
        messagebox.showerror("エラー", f"日本語フォントの登録に失敗: {str(e)}")
        return None

    # 韓国語フォント登録
    for ko_font in FONT_CONFIG['ko']:
        try:
            ko_font_path = resource_path(ko_font['file'])
            pdfmetrics.registerFont(TTFont(ko_font['name'], ko_font_path))
            return ko_font['name']
        except Exception:
            continue
    messagebox.showerror("エラー", "韓国語フォントが見つかりません")
    return None

def parse_srt(file_path):
    """SRTファイル解析関数"""
    entries = []
    timecode_pattern = re.compile(
        r'(\d+:\d+:\d+)[,.](\d+)\s*-->\s*(\d+:\d+:\d+)[,.](\d+)'
    )
    
    try:
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            content = f.read().strip()
            
        blocks = content.split('\n\n')
        for block in blocks:
            lines = block.split('\n')
            if len(lines) >= 3:
                time_match = timecode_pattern.match(lines[1])
                if time_match:
                    start = f"{time_match.group(1)},{time_match.group(2).ljust(3,'0')[:3]}"
                    text = ' '.join(lines[2:]).replace('\n', ' ')
                    entries.append((start, text))
    except Exception as e:
        messagebox.showerror("エラー", f"SRT解析エラー: {str(e)}")
        return None
        
    return entries

def time_to_ms(time_str):
    """時間をミリ秒に変換"""
    h, m, s = time_str.split(':')
    s, ms = s.split(',')
    return int(h)*3600000 + int(m)*60000 + int(s)*1000 + int(ms)

def align_subtitles(ko_entries, ja_entries, threshold=500):
    """字幕同期関数"""
    aligned = []
    ko_ptr = 0
    ja_ptr = 0

    while ko_ptr < len(ko_entries) and ja_ptr < len(ja_entries):
        ko_time = time_to_ms(ko_entries[ko_ptr][0])
        ja_time = time_to_ms(ja_entries[ja_ptr][0])

        if abs(ko_time - ja_time) <= threshold:
            aligned.append((
                ko_entries[ko_ptr][0],
                ko_entries[ko_ptr][1],
                ja_entries[ja_ptr][1]
            ))
            ko_ptr += 1
            ja_ptr += 1
        elif ko_time < ja_time:
            aligned.append((ko_entries[ko_ptr][0], ko_entries[ko_ptr][1], ''))
            ko_ptr += 1
        else:
            aligned.append((ja_entries[ja_ptr][0], '', ja_entries[ja_ptr][1]))
            ja_ptr += 1

    while ko_ptr < len(ko_entries):
        aligned.append((ko_entries[ko_ptr][0], ko_entries[ko_ptr][1], ''))
        ko_ptr += 1

    while ja_ptr < len(ja_entries):
        aligned.append((ja_entries[ja_ptr][0], '', ja_entries[ja_ptr][1]))
        ja_ptr += 1

    return aligned

def create_pdf(output_path, aligned_entries, ko_font_name):
    """PDF生成関数"""
    styles = getSampleStyleSheet()
    
    # スタイル設定（韓国語）
    ko_style = styles['Normal'].clone('KoreanStyle')
    ko_style.fontName = ko_font_name
    ko_style.fontSize = 9
    ko_style.leading = 11
    ko_style.splitLongWords = False
    ko_style.alignment = 4

    # スタイル設定（日本語）
    ja_style = styles['Normal'].clone('JapaneseStyle')
    ja_style.fontName = FONT_CONFIG['ja']['name']
    ja_style.fontSize = 9
    ja_style.leading = 11
    ja_style.splitLongWords = False
    ja_style.alignment = 4

    doc = SimpleDocTemplate(output_path, pagesize=A4,
                              rightMargin=10*mm, leftMargin=10*mm,
                              topMargin=10*mm, bottomMargin=10*mm)
    
    data = []
    for timecode, ko_text, ja_text in aligned_entries:
        # HTMLタグ除去
        ko_text = re.sub(r'<[^>]+>', '', ko_text)
        ja_text = re.sub(r'<[^>]+>', '', ja_text)
        
        # 段落作成
        ko_para = Paragraph(
            f"<font color='#0066CC' size='8'>{timecode}</font><br/>{ko_text}",
            ko_style
        )
        ja_para = Paragraph(
            f"<font color='#CC0033' size='8'>{timecode}</font><br/>{ja_text}",
            ja_style
        )
        
        data.append([ko_para, ja_para])

    table = Table(data, colWidths=[doc.width/2.0]*2)
    table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('PADDING', (0,0), (-1,-1), 3),
        ('BACKGROUND', (0,0), (0,-1), '#F8F8FF'),
        ('BACKGROUND', (1,0), (1,-1), '#FFF8F0'),
        ('BOX', (0,0), (-1,-1), 0.5, colors.grey),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.grey),
    ]))

    doc.build([table])

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("字幕PDF生成ツール")
        self.root.geometry("500x300")

        # 韓国語字幕ファイル選択
        self.ko_label = tk.Label(root, text="韓国語字幕ファイル:")
        self.ko_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.ko_entry = tk.Entry(root, width=40)
        self.ko_entry.grid(row=0, column=1, padx=10, pady=10)
        self.ko_button = tk.Button(root, text="選択", command=self.select_ko_file)
        self.ko_button.grid(row=0, column=2, padx=10, pady=10)

        # 日本語字幕ファイル選択
        self.ja_label = tk.Label(root, text="日本語字幕ファイル:")
        self.ja_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.ja_entry = tk.Entry(root, width=40)
        self.ja_entry.grid(row=1, column=1, padx=10, pady=10)
        self.ja_button = tk.Button(root, text="選択", command=self.select_ja_file)
        self.ja_button.grid(row=1, column=2, padx=10, pady=10)

        # 出力PDFファイル選択
        self.output_label = tk.Label(root, text="出力PDFファイル:")
        self.output_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.output_entry = tk.Entry(root, width=40)
        self.output_entry.grid(row=2, column=1, padx=10, pady=10)
        self.output_button = tk.Button(root, text="選択", command=self.select_output_file)
        self.output_button.grid(row=2, column=2, padx=10, pady=10)

        # 進捗バー
        self.progress = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
        self.progress.grid(row=3, column=0, columnspan=3, padx=10, pady=20)

        # 実行ボタン
        self.run_button = tk.Button(root, text="実行", command=self.run)
        self.run_button.grid(row=4, column=1, padx=10, pady=20)

    def select_ko_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("SRTファイル", "*.srt")])
        if file_path:
            self.ko_entry.delete(0, tk.END)
            self.ko_entry.insert(0, file_path)

    def select_ja_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("SRTファイル", "*.srt")])
        if file_path:
            self.ja_entry.delete(0, tk.END)
            self.ja_entry.insert(0, file_path)

    def select_output_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDFファイル", "*.pdf")])
        if file_path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, file_path)

    def run(self):
        ko_file = self.ko_entry.get()
        ja_file = self.ja_entry.get()
        output_file = self.output_entry.get()

        if not ko_file or not ja_file or not output_file:
            messagebox.showerror("エラー", "すべてのファイルを選択してください")
            return

        # フォント登録
        ko_font_name = register_fonts()
        if not ko_font_name:
            return

        # SRTファイル解析
        self.progress["value"] = 20
        self.root.update_idletasks()
        ko_entries = parse_srt(ko_file)
        ja_entries = parse_srt(ja_file)

        if not ko_entries or not ja_entries:
            return

        # 字幕同期
        self.progress["value"] = 50
        self.root.update_idletasks()
        aligned = align_subtitles(ko_entries, ja_entries)

        # PDF生成
        self.progress["value"] = 80
        self.root.update_idletasks()
        create_pdf(output_file, aligned, ko_font_name)

        # 完了
        self.progress["value"] = 100
        self.root.update_idletasks()
        messagebox.showinfo("完了", "PDFが正常に作成されました")

if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
