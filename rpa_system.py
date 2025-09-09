"""
検査表自動計算システム

このプログラムは、チェーンガイドの検査表を自動計算するためのGUIアプリケーションです。
機種、棚数、インチ数、高さを入力することで、A値、B値、C値、D値、H値、スタビ値を計算します。
"""

import tkinter as tk
from tkinter import ttk, messagebox
import json
import os


# 定数定義
MACHINE_DATA = {
    "200": {
        "chain_data": 31.75,
        "inches": [10, 12, 14, 16, 18, 20],
        "rings": [8, 10, 11, 13, 14, 15],
        "data1": 20,
        "elb": 1240,
        "elc": 120,
        "erb": 1240,
        "erc": 120,
        "er": 162,
        "el": 107,
        "st": 73 + 217 + 2,
    },
    "201": {
        "chain_data": 31.75,
        "inches": [10, 12, 14, 16, 18, 20],
        "rings": [8, 10, 11, 13, 14, 15],
        "data1": 20,
        "elb": 1690,
        "elc": 120,
        "erb": 1690,
        "erc": 120,
        "er": 162,
        "el": 107,
        "st": 54 + 294 + 2,
    },
    "350": {
        "chain_data": 50.8,
        "inches": [10, 12, 14, 16, 18, 20],
        "rings": [5, 6, 7, 8, 9, 10],
        "data1": 14,
        "erb": 1282,
        "erc": 216,
        "elc": 140,
        "elb": 1282,
        "er": 221.6,
        "el": 90.6,
        "st": 57 + 297 + 2,
    },
    "351": {
        "chain_data": 50.8,
        "inches": [10, 12, 14, 16, 18, 20],
        "rings": [5, 6, 7, 8, 9, 10],
        "data1": 14,
        "erb": 1725,
        "erc": 216,
        "elb": 1725,
        "elc": 140,
        "er": 221.6,
        "el": 90.6,
        "st": 57 + 297 + 2,
    },
}

# 高さとB値の対応表
HEIGHT_B_MAPPING = {
    "200": {2750: 250, 2500: 225, 2250: 0},
    "201": {3000: 250, 2750: 300, 2500: 175, 2250: 0},
}

# H値の計算用データ（350, 351のみ使用）
H_VALUES = {
    "350": {"10": 388, "12": 388, "14": 288, "16": 288},
    "351": {"10": 388, "12": 388, "14": 288, "16": 288},
}

# C値の計算用データ
C_VALUES = {
    "200": {
        4: 630, 5: 1000, 6: 1000, 7: 1500, 8: 1500, 9: 1500, 10: 1500,
        11: 1500, 12: 1500, 13: 1500, 14: 1500, 15: 1500, 16: 1500, 17: 1500,
    },
    "201": {
        4: 730, 5: 980, 6: 1000, 7: 1480, 8: 1500, 9: 1500, 10: 1500,
        11: 1500, 12: 1500, 13: 1500, 14: 1500, 15: 1500, 16: 1500, 17: 1500,
    },
    "350": {
        3: 1000, 4: 1000, 5: 1000, 6: 1500, 7: 2000, 8: 2000, 9: 2000,
        10: 2000, 11: 2000, 12: 2000, 13: 2000, 14: 2000, 15: 2000, 16: 2000, 17: 2000,
    },
    "351": {
        3: 850, 4: 900, 5: 900, 6: 1500, 7: 1900, 8: 1900, 9: 1900,
        10: 1900, 11: 1900, 12: 1900, 13: 1900, 14: 1900, 15: 1900, 16: 1900, 17: 1900,
    },
}

# D値の計算用データ
D_VALUES = {
    "200": {11: 1000, 12: 1000, 13: 1500, 14: 1500, 15: 1500, 16: 1500, 17: 1500},
    "201": {11: 980, 12: 1000, 13: 1480, 14: 1500, 15: 1500, 16: 1500, 17: 1500},
    "350": {10: 500, 11: 500, 12: 1000, 13: 1000, 14: 1500, 15: 1500, 16: 2000, 17: 2000},
    "351": {10: 500, 11: 500, 12: 1000, 13: 1000, 14: 1500, 15: 1500, 16: 2000, 17: 2000},
}

# 高さの制約
MIN_HEIGHT = 2500
HEIGHT_RESOLUTION = 250


class RPASystem:
    """
    検査表自動計算システムのメインクラス
    
    チェーンガイドの検査表を自動計算するためのGUIアプリケーションを提供します。
    """
    
    def __init__(self):
        """RPASystemの初期化"""
        self.setup_gui()

    # ==================== GUI設定メソッド ====================
    
    def setup_gui(self):
        """GUIの設定"""
        self.root = tk.Tk()
        self.root.title("検査表自動計算")
        self.root.geometry("1000x800")

        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # タイトル
        title_label = ttk.Label(main_frame, text="検査表自動計算", font=("Arial", 20, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # 入力フィールド
        ttk.Label(main_frame, text="機種:", font=("Arial", 12)).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.machine_var = tk.StringVar()
        machine_combo = ttk.Combobox(
            main_frame, textvariable=self.machine_var, values=list(MACHINE_DATA.keys()), state="readonly", font=("Arial", 12)
        )
        machine_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        machine_combo.bind("<<ComboboxSelected>>", self.on_machine_change)

        ttk.Label(main_frame, text="棚数:", font=("Arial", 12)).grid(row=2, column=0, sticky=tk.W, pady=5)
        self.shelf_var = tk.StringVar()
        shelf_entry = ttk.Entry(main_frame, textvariable=self.shelf_var, font=("Arial", 12))
        shelf_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)

        ttk.Label(main_frame, text="インチ数:", font=("Arial", 12)).grid(row=3, column=0, sticky=tk.W, pady=5)
        self.inch_var = tk.StringVar()
        self.inch_combo = ttk.Combobox(main_frame, textvariable=self.inch_var, state="readonly", font=("Arial", 12))
        self.inch_combo.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5)

        ttk.Label(main_frame, text="高さ:", font=("Arial", 12)).grid(row=4, column=0, sticky=tk.W, pady=5)
        self.height_var = tk.StringVar()
        height_entry = ttk.Entry(main_frame, textvariable=self.height_var, font=("Arial", 12))
        height_entry.grid(row=4, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # 高さの入力例を表示
        height_hint = ttk.Label(main_frame, text=f"※{MIN_HEIGHT}mm以上、{HEIGHT_RESOLUTION}mm刻み (例: {MIN_HEIGHT}, {MIN_HEIGHT + HEIGHT_RESOLUTION}, {MIN_HEIGHT + HEIGHT_RESOLUTION * 2}...)", 
                               font=("Arial", 9), foreground="gray")
        height_hint.grid(row=4, column=2, sticky=tk.W, padx=(10, 0), pady=5)

        # 計算ボタン
        calc_button = ttk.Button(main_frame, text="計算実行", command=self.calculate)
        calc_button.grid(row=5, column=0, columnspan=2, pady=20)
        # ボタンのフォントサイズを設定
        style = ttk.Style()
        style.configure("Large.TButton", font=("Arial", 12, "bold"))
        calc_button.configure(style="Large.TButton")

        # 結果表示
        result_frame = ttk.LabelFrame(main_frame, text="計算結果", padding="10")
        result_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        # フレームタイトルのフォントサイズを設定
        style = ttk.Style()
        style.configure("Large.TLabelframe.Label", font=("Arial", 12, "bold"))
        result_frame.configure(style="Large.TLabelframe")

        # テキストエリアの設定を改善
        self.result_text = tk.Text(
            result_frame, 
            height=20, 
            width=90,
            font=("Consolas", 12),
            wrap=tk.WORD,
            bg="#f8f9fa",
            fg="#212529",
            relief=tk.FLAT,
            borderwidth=1
        )
        self.result_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # スクロールバー
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.result_text.configure(yscrollcommand=scrollbar.set)
        
        # グリッドの重み設定
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)

        # 初期化
        self.on_machine_change()

    # ==================== イベントハンドラメソッド ====================
    
    def on_machine_change(self, event=None):
        """機種が変更された時の処理"""
        machine = self.machine_var.get()
        if machine in MACHINE_DATA:
            inches = MACHINE_DATA[machine]["inches"]
            self.inch_combo["values"] = inches
            if inches:
                self.inch_combo.set(inches[0])

    # ==================== 計算メソッド ====================
    
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

            # 高さの例外処理
            if height < MIN_HEIGHT:
                messagebox.showerror("高さエラー", f"高さは最小{MIN_HEIGHT}mm以上で入力してください")
                return
            
            # 高さの分解能チェック
            if (height - MIN_HEIGHT) % HEIGHT_RESOLUTION != 0:
                messagebox.showerror("高さエラー", f"高さは{MIN_HEIGHT}mmから{HEIGHT_RESOLUTION}mm刻みで入力してください\n例: {MIN_HEIGHT}, {MIN_HEIGHT + HEIGHT_RESOLUTION}, {MIN_HEIGHT + HEIGHT_RESOLUTION * 2}...")
                return

            # 機種データの取得
            machine_info = MACHINE_DATA.get(machine)
            if not machine_info:
                messagebox.showerror("エラー", "無効な機種です")
                return

            # インチ数に対応するリング数の取得
            inch_index = machine_info["inches"].index(inch)
            ring_count = machine_info["rings"][inch_index]

            # A値の計算
            chain_data = machine_info["chain_data"]
            data1 = machine_info["data1"]
            A = ((shelf_count * ring_count - data1) / 2 * chain_data) + 60
            A1 = (shelf_count * ring_count - data1) / 2 * chain_data

            # C値とD値の計算
            if machine in ["350", "351"]:
                D = self.calculate_c_value(machine, height) - 160
                C = self.calculate_d_value(machine, height) - 160

            elif machine in ["200", "201"]:
                C = self.calculate_c_value(machine, height)
                D = self.calculate_d_value(machine, height)

            # H値の計算
            H = self.calculate_h_value(machine, inch)

            if machine in ["350", "351"]:
                if machine == "351":
                    if inch in [10, 12] and height < 4000:
                        if D <= 0:
                            B = A - C - 160 * 2 - H + 63
                        elif C <= 0:
                            B = A - D - 160 * 2 - H + 63
                    elif C <= 0 and D <= 0:
                        B = A - 160 - H - 37
                    else:
                        B = A - C - D - 160 * 3 - H - 37
                else:  # 350
                    if C <= 0 and D <= 0:
                        B = A - 160 - H - 37
                    elif C <= 0:
                        B = A - D - 160 * 2 - H - 37
                    elif D <= 0:
                        B = A - C - 160 * 2 - H - 37
                    else:
                        B = A - C - D - 160 * 3 - H - 37

            # B値の計算（200, 201の場合）
            elif machine in ["200", "201"]:
                B = self.calculate_b_value(machine, height)

            ST = A1 - machine_info["st"]
            ELB = machine_info["elb"]
            ELC = machine_info["elc"]
            ERB = machine_info["erb"]
            ERC = machine_info["erc"]
            EB = machine_info["er"]
            EC = machine_info["el"]

            # 出力値が0以下の場合は0に設定
            A = max(0, A)
            B = max(0, B)
            C = max(0, C)
            D = max(0, D)
            H = max(0, H)
            ST = max(0, ST)

            # B値の表示用テキストを準備
            b_display = f"{B:.2f}"
            if machine in ["350", "351"] and height <= 4000:
                b_display += " (H=4000以下の場合、図面確認)"
            
            result = f"""═══════════════════════════════════════════════════════════
                    検査表自動計算結果
═══════════════════════════════════════════════════════════

【入力条件】
機種: {machine}
棚数: {shelf_count}
インチ数: {inch}
高さ: {height}mm
リング数: {ring_count}
チェーンデータ: {chain_data}mm
データ1: {data1}

【チェーンガイド計算結果】
┌─────────────────────────────────────────────────────────┐
│  A値 = {A:>8.2f}mm                                      │
│  B値 = {b_display:>8}!!351の場合は図面確認                                   │
│  C値 = {C:>8.0f}mm                                      │
│  D値 = {D:>8.0f}mm                                      │
│  H値 = {H:>8.0f}mm                                      │
│  スタビ = {ST:>6.2f}mm!!図面確認                                  │
└─────────────────────────────────────────────────────────┘

【エンドフレーム仕様】
┌─────────────────────────────────────────────────────────┐
│  エンドフレームA = {height:>6.0f}mm                     │
│  エンドフレーム左B = {ELB:>6.0f}mm                      │
│  エンドフレーム左C = {ELC:>6.0f}mm                      │
│  エンドフレーム右B = {ERB:>6.0f}mm                      │
│  エンドフレーム右C = {ERC:>6.0f}mm                      │
│  溶接補強高さ右 = {EB:>6.0f}mm                          │
│  溶接補強高さ左 = {EC:>6.0f}mm                          │
└─────────────────────────────────────────────────────────┘

【計算詳細】
• H値: 機種{machine}、インチ数{inch}の場合
• C値: 機種{machine}、高さ{height}mmの場合  
• D値: 機種{machine}、高さ{height}mmの場合

═══════════════════════════════════════════════════════════
"""
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(1.0, result)

        except ValueError as e:
            error_msg = f"""入力エラーが発生しました。

以下の項目を確認してください：
• 棚数: 数値で入力してください
• 高さ: 数値で入力してください（{MIN_HEIGHT}mm以上、{HEIGHT_RESOLUTION}mm刻み）
• 機種: 選択してください
• インチ数: 選択してください

高さの例: {MIN_HEIGHT}, {MIN_HEIGHT + HEIGHT_RESOLUTION}, {MIN_HEIGHT + HEIGHT_RESOLUTION * 2}, {MIN_HEIGHT + HEIGHT_RESOLUTION * 3}, {MIN_HEIGHT + HEIGHT_RESOLUTION * 4}..."""
            messagebox.showerror("入力エラー", error_msg)
        except Exception as e:
            error_msg = f"""計算中にエラーが発生しました。

エラー詳細: {str(e)}

以下の点を確認してください：
• 入力値が正しい範囲内か
• 機種とインチ数の組み合わせが正しいか
• 高さの値が適切か"""
            messagebox.showerror("計算エラー", error_msg)

    # ==================== 個別計算メソッド ====================
    
    def calculate_b_value(self, machine, height):
        """
        B値の計算
        
        Args:
            machine (str): 機種コード ("200", "201", "350", "351")
            height (int): 高さ(mm)
            
        Returns:
            int: 計算されたB値
        """
        if machine in ["200", "201"]:
            if machine == "200":
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
        """
        H値の計算（350, 351のみ使用）
        
        Args:
            machine (str): 機種コード ("350", "351")
            inch (int): インチ数
            
        Returns:
            int: 計算されたH値
        """
        if machine in ["350", "351"]:
            if inch in [10, 12]:
                return 388
            elif inch in [14, 16]:
                return 288
            else:
                return 0  # その他のインチ数の場合
        else:
            return 0  # 200, 201の場合は使用しない

    def calculate_c_value(self, machine, height):
        """
        C値の計算
        
        Args:
            machine (str): 機種コード
            height (int): 高さ(mm)
            
        Returns:
            int: 計算されたC値
        """
        # 高さからHインデックスを計算（分解能250）
        h_index = self.height_to_h_index(height)

        if machine in C_VALUES and h_index in C_VALUES[machine]:
            # 機種351の場合はインチ数も考慮
            if machine == "351":
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
                    return C_VALUES[machine].get(h_index, 0)
            else:
                return C_VALUES[machine].get(h_index, 0)
        return 0

    def calculate_d_value(self, machine, height):
        """
        D値の計算
        
        Args:
            machine (str): 機種コード
            height (int): 高さ(mm)
            
        Returns:
            int: 計算されたD値
        """
        # 高さからHインデックスを計算（分解能250）
        h_index = self.height_to_h_index(height)

        if machine in D_VALUES and h_index in D_VALUES[machine]:
            return D_VALUES[machine][h_index]
        return 0

    def height_to_h_index(self, height):
        """
        高さからHインデックスを計算（分解能250）
        
        Args:
            height (int): 高さ(mm)
            
        Returns:
            int: Hインデックス (3-17)
        """
        # 高さ3000から始まり、分解能250
        if height < 3000:
            return 3  # 最小値
        elif height > 6750:  # 3000 + 250 * 15
            return 17  # 最大値
        else:
            # 3000から始まる高さを250で割ってインデックスを計算
            return 3 + ((height - 3000) // HEIGHT_RESOLUTION)

    # ==================== メイン実行メソッド ====================
    
    def run(self):
        """GUIの実行"""
        self.root.mainloop()


# ==================== メイン実行部分 ====================

if __name__ == "__main__":
    rpa = RPASystem()
    rpa.run()