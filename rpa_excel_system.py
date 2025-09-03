import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill
import os
from datetime import datetime

class RPAExcelSystem:
    def __init__(self):
        # 機種データの初期化
        self.machine_data = {
            '200': {'chain_data': 31.75, 'inches': [10, 11, 12, 14, 15, 16, 18, 19, 20], 'rings': [8, 9, 10, 11, 12, 13, 14, 15, 16], 'data1': 20},
            '201': {'chain_data': 31.75, 'inches': [10, 11, 12, 14, 15, 16, 18, 19, 20], 'rings': [8, 10, 11, 12, 13, 14, 15, 16], 'data1': 20},
            '350': {'chain_data': 50.8, 'inches': [10, 11, 12, 14, 15, 16, 18, 19, 20], 'rings': [5, 6, 7, 8, 9, 10, 11, 12], 'data1': 14},
            '351': {'chain_data': 50.8, 'inches': [10, 11, 12, 14, 15, 16, 18, 19, 20], 'rings': [5, 6, 7, 8, 9, 10, 11, 12], 'data1': 14}
        }
        
        # 高さとB値の対応表
        self.height_b_mapping = {
            '200': {2750: 225, 2500: 225, 2250: 0},
            '201': {3000: 250, 2750: 200, 2500: 175, 2250: 0}
        }
        
        # H値の計算用データ（350, 351のみ使用）
        self.h_values = {
            '350': {'10': 388, '12': 388, '14': 288, '16': 288},
            '351': {'10': 388, '12': 388, '14': 288, '16': 288}
        }
        
        # C値の計算用データ
        self.c_values = {
            '200': {4: 630, 5: 1000, 6: 1000, 7: 1500, 8: 1500, 9: 1500, 10: 1500, 11: 1500, 12: 1500, 13: 1500, 14: 1500, 15: 1500, 16: 1500, 17: 1500},
            '201': {4: 730, 5: 980, 6: 1000, 7: 1480, 8: 1500, 9: 1500, 10: 1500, 11: 1500, 12: 1500, 13: 1500, 14: 1500, 15: 1500, 16: 1500, 17: 1500},
            '350': {3: 1000, 4: 1000, 5: 1000, 6: 1500, 7: 2000, 8: 2000, 9: 2000, 10: 2000, 11: 2000, 12: 2000, 13: 2000, 14: 2000, 15: 2000, 16: 2000, 17: 2000},
            '351': {3: 850, 4: 900, 5: 900, 6: 1500, 7: 1900, 8: 1900, 9: 1900, 10: 1900, 11: 1900, 12: 1900, 13: 1900, 14: 1900, 15: 1900, 16: 1900, 17: 1900}
        }
        
        # D値の計算用データ
        self.d_values = {
            '200': {11: 1000, 12: 1000, 13: 1500, 14: 1500, 15: 1500, 16: 1500, 17: 1500},
            '201': {11: 980, 12: 1000, 13: 1480, 14: 1500, 15: 1500, 16: 1500, 17: 1500},
            '350': {10: 500, 11: 500, 12: 1000, 13: 1000, 14: 1500, 15: 1500, 16: 2000, 17: 2000},
            '351': {10: 500, 11: 500, 12: 1000, 13: 1000, 14: 1500, 15: 1500, 16: 2000, 17: 2000}
        }
        
        self.setup_gui()
    
    def setup_gui(self):
        """GUIの設定"""
        self.root = tk.Tk()
        self.root.title("会社RPAシステム - Excel連携版")
        self.root.geometry("800x700")
        
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # タイトル
        title_label = ttk.Label(main_frame, text="RPA計算システム - Excel連携版", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 入力フィールド
        ttk.Label(main_frame, text="機種:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.machine_var = tk.StringVar()
        machine_combo = ttk.Combobox(main_frame, textvariable=self.machine_var, 
                                    values=['200', '201', '350', '351'], state="readonly")
        machine_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        machine_combo.bind('<<ComboboxSelected>>', self.on_machine_change)
        
        ttk.Label(main_frame, text="棚数:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.shelf_var = tk.StringVar()
        shelf_entry = ttk.Entry(main_frame, textvariable=self.shelf_var)
        shelf_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(main_frame, text="インチ数:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.inch_var = tk.StringVar()
        self.inch_combo = ttk.Combobox(main_frame, textvariable=self.inch_var, state="readonly")
        self.inch_combo.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(main_frame, text="高さ:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.height_var = tk.StringVar()
        height_entry = ttk.Entry(main_frame, textvariable=self.height_var)
        height_entry.grid(row=4, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # ボタンフレーム
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=20)
        
        calc_button = ttk.Button(button_frame, text="計算実行", command=self.calculate)
        calc_button.pack(side=tk.LEFT, padx=5)
        
        excel_button = ttk.Button(button_frame, text="Excel出力", command=self.export_to_excel)
        excel_button.pack(side=tk.LEFT, padx=5)
        
        batch_button = ttk.Button(button_frame, text="一括処理", command=self.batch_process)
        batch_button.pack(side=tk.LEFT, padx=5)
        
        # 結果表示
        result_frame = ttk.LabelFrame(main_frame, text="計算結果", padding="10")
        result_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        self.result_text = tk.Text(result_frame, height=15, width=70)
        self.result_text.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # スクロールバー
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.result_text.configure(yscrollcommand=scrollbar.set)
        
        # 初期化
        self.on_machine_change()
        self.calculation_history = []
    
    def on_machine_change(self, event=None):
        """機種が変更された時の処理"""
        machine = self.machine_var.get()
        if machine in self.machine_data:
            inches = self.machine_data[machine]['inches']
            self.inch_combo['values'] = inches
            if inches:
                self.inch_combo.set(inches[0])
    
    def calculate(self):
        """計算実行"""
        try:
            # 入力値の取得
            machine = self.machine_var.get()
            shelf_count = int(self.shelf_var.get())
            inch = int(self.inch_var.get())
            height = int(self.height_var.get())
            
            if not machine:
                messagebox.showerror("エラー", "機種を選択してください")
                return
            
            # 機種データの取得
            machine_info = self.machine_data.get(machine)
            if not machine_info:
                messagebox.showerror("エラー", "無効な機種です")
                return
            
            # インチ数に対応するリング数の取得
            inch_index = machine_info['inches'].index(inch)
            ring_count = machine_info['rings'][inch_index]
            
            # A値の計算
            chain_data = machine_info['chain_data']
            data1 = machine_info['data1']
            A = ((shelf_count * ring_count - data1) / 2 * chain_data) + 60
            
            # B値の計算
            B = self.calculate_b_value(machine, height)
            
            # C値の計算
            C = self.calculate_c_value(machine, height)
            
            # D値の計算
            D = self.calculate_d_value(machine, height)
            
            # H値の計算
            H = self.calculate_h_value(machine, inch)
            
            # 結果を履歴に保存
            result_data = {
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'machine': machine,
                'shelf_count': shelf_count,
                'inch': inch,
                'height': height,
                'ring_count': ring_count,
                'chain_data': chain_data,
                'data1': data1,
                'A': A,
                'B': B,
                'C': C,
                'D': D,
                'H': H
            }
            self.calculation_history.append(result_data)
            
            # 結果の表示
            result = f"""計算結果:
機種: {machine}
棚数: {shelf_count}
インチ数: {inch}
高さ: {height}
リング数: {ring_count}
チェーンデータ: {chain_data}
データ1: {data1}

出力値:
A = {A:.2f}
B = {B}
C = {C}
D = {D}
H = {H}

計算詳細:
- H値: 機種{machine}、インチ数{inch}の場合
- C値: 機種{machine}、高さ{height}の場合
- D値: 機種{machine}、高さ{height}の場合

計算履歴: {len(self.calculation_history)}件
"""
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(1.0, result)
            
        except ValueError as e:
            messagebox.showerror("エラー", "正しい数値を入力してください")
        except Exception as e:
            messagebox.showerror("エラー", f"計算中にエラーが発生しました: {str(e)}")
    
    def calculate_b_value(self, machine, height):
        """B値の計算"""
        if machine in ['200', '201']:
            if machine == '200':
                if height > 2750:
                    return 250
                elif height >= 2500:
                    return 225
                elif height >= 2250:
                    return 0
                else:
                    return 0
            else:  # 201
                if height > 3000:
                    return 250
                elif height >= 2750:
                    return 200
                elif height >= 2500:
                    return 175
                elif height >= 2250:
                    return 0
                else:
                    return 0
        else:
            return 0
    
    def calculate_h_value(self, machine, inch):
        """H値の計算（350, 351のみ使用）"""
        if machine in ['350', '351']:
            if inch in [10, 12]:
                return 388
            elif inch in [14, 16]:
                return 288
            else:
                return 0  # その他のインチ数の場合
        else:
            return 0  # 200, 201の場合は使用しない
    
    def calculate_c_value(self, machine, height):
        """C値の計算"""
        # 高さからH値を計算（分解能250）
        h_index = self.height_to_h_index(height)
        
        if machine in self.c_values and h_index in self.c_values[machine]:
            # 機種351の場合はインチ数も考慮
            if machine == '351':
                inch = int(self.inch_var.get())
                if h_index == 3:
                    if inch in [14, 16]:
                        return 850
                    else:  # 10, 12インチ
                        return 950
                elif h_index == 4:
                    if inch in [14, 16]:
                        return 900
                    else:  # 10, 12インチ
                        return 1000
                elif h_index == 5:
                    if inch in [14, 16]:
                        return 900
                    else:  # 10, 12インチ
                        return 1000
                else:
                    return self.c_values[machine].get(h_index, 0)
            else:
                return self.c_values[machine].get(h_index, 0)
        return 0
    
    def calculate_d_value(self, machine, height):
        """D値の計算"""
        # 高さからH値を計算（分解能250）
        h_index = self.height_to_h_index(height)
        
        if machine in self.d_values and h_index in self.d_values[machine]:
            return self.d_values[machine][h_index]
        return 0
    
    def height_to_h_index(self, height):
        """高さからHインデックスを計算（分解能250）"""
        # 高さ3000から始まり、分解能250
        if height < 3000:
            return 3  # 最小値
        elif height > 6750:  # 3000 + 250 * 15
            return 17  # 最大値
        else:
            # 3000から始まる高さを250で割ってインデックスを計算
            return 3 + ((height - 3000) // 250)
    
    def export_to_excel(self):
        """Excelファイルに出力"""
        if not self.calculation_history:
            messagebox.showwarning("警告", "計算履歴がありません")
            return
        
        try:
            # ファイル保存ダイアログ
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            
            if not filename:
                return
            
            # DataFrameの作成
            df = pd.DataFrame(self.calculation_history)
            
            # Excelファイルに出力
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='計算結果', index=False)
                
                # ワークブックとワークシートの取得
                workbook = writer.book
                worksheet = writer.sheets['計算結果']
                
                # ヘッダーのスタイル設定
                header_font = Font(bold=True, color="FFFFFF")
                header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                
                for cell in worksheet[1]:
                    cell.font = header_font
                    cell.fill = header_fill
                
                # 列幅の自動調整
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            messagebox.showinfo("成功", f"Excelファイルに出力しました:\n{filename}")
            
        except Exception as e:
            messagebox.showerror("エラー", f"Excel出力中にエラーが発生しました: {str(e)}")
    
    def batch_process(self):
        """一括処理（サンプルデータ）"""
        sample_data = [
            {'machine': '200', 'shelf_count': 16, 'inch': 12, 'height': 2500},
            {'machine': '201', 'shelf_count': 20, 'inch': 15, 'height': 3000},
            {'machine': '350', 'shelf_count': 12, 'inch': 10, 'height': 2000},
            {'machine': '351', 'shelf_count': 18, 'inch': 18, 'height': 2500}
        ]
        
        self.calculation_history.clear()
        
        for data in sample_data:
            self.machine_var.set(data['machine'])
            self.shelf_var.set(str(data['shelf_count']))
            self.inch_var.set(str(data['inch']))
            self.height_var.set(str(data['height']))
            self.calculate()
        
        messagebox.showinfo("完了", f"一括処理が完了しました。{len(sample_data)}件のデータを処理しました。")
    
    def run(self):
        """GUIの実行"""
        self.root.mainloop()

if __name__ == "__main__":
    rpa = RPAExcelSystem()
    rpa.run()
