import tkinter as tk
from tkinter import ttk, messagebox
import pyodbc
import json
import os

class ConnectionSettingsDialog:
    def __init__(self, parent, current_settings):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("接続設定")
        self.dialog.geometry("400x300")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.result = None
        self.current_settings = current_settings
        self.setup_variables()
        self.create_widgets()
        self.load_settings()
        
    def setup_variables(self):
        self.server = tk.StringVar(value=self.current_settings['server'])
        self.username = tk.StringVar(value=self.current_settings['username'])
        self.password = tk.StringVar(value=self.current_settings['password'])
        self.driver = tk.StringVar(value=self.current_settings['driver'])
        
    def get_available_drivers(self):
        try:
            return [driver for driver in pyodbc.drivers() if 'SQL Server' in driver]
        except:
            return ['ODBC Driver 17 for SQL Server']
            
    def create_widgets(self):
        # メインフレーム
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # サーバー設定
        ttk.Label(main_frame, text="サーバー名:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.server, width=40).grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        
        # ユーザー名設定
        ttk.Label(main_frame, text="ユーザー名:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.username, width=40).grid(row=1, column=1, sticky='ew', padx=5, pady=5)
        
        # パスワード設定
        ttk.Label(main_frame, text="パスワード:").grid(row=2, column=0, sticky='w', padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.password, show="*", width=40).grid(row=2, column=1, sticky='ew', padx=5, pady=5)
        
        # ドライバー選択
        ttk.Label(main_frame, text="ドライバー:").grid(row=3, column=0, sticky='w', padx=5, pady=5)
        driver_combo = ttk.Combobox(main_frame, textvariable=self.driver, values=self.get_available_drivers(), width=37)
        driver_combo.grid(row=3, column=1, sticky='ew', padx=5, pady=5)
        
        # テスト接続ボタン
        ttk.Button(main_frame, text="接続テスト", command=self.test_connection).grid(row=4, column=0, columnspan=2, pady=20)
        
        # 保存・キャンセルボタン
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=10)
        ttk.Button(button_frame, text="保存", command=self.save).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="キャンセル", command=self.cancel).pack(side=tk.LEFT, padx=5)
        
    def test_connection(self):
        try:
            connection_string = (
                f"DRIVER={{{self.driver.get()}}};"
                f"SERVER={self.server.get()};"
                f"UID={self.username.get()};"
                f"PWD={self.password.get()}"
            )
            conn = pyodbc.connect(connection_string)
            conn.close()
            messagebox.showinfo("成功", "接続テストに成功しました")
        except Exception as e:
            messagebox.showerror("エラー", f"接続テストに失敗しました: {str(e)}")
            
    def save(self):
        if not all([self.server.get(), self.username.get(), self.password.get(), self.driver.get()]):
            messagebox.showwarning("警告", "すべての項目を入力してください")
            return
            
        self.result = {
            'server': self.server.get(),
            'username': self.username.get(),
            'password': self.password.get(),
            'driver': self.driver.get()
        }
        self.dialog.destroy()
        
    def cancel(self):
        self.dialog.destroy()
        
    def load_settings(self):
        # 既存の設定をフィールドに設定
        self.server.set(self.current_settings.get('server', ''))
        self.username.set(self.current_settings.get('username', ''))
        self.password.set(self.current_settings.get('password', ''))
        self.driver.set(self.current_settings.get('driver', ''))
        
class ColumnDialog:
    def __init__(self, sql_manager, title, current_values=None):
        """
        Args:
            sql_manager: SQLTableManagerのインスタンス
            title: ダイアログのタイトル
            current_values: 現在の値（編集時）
        """
        self.sql_manager = sql_manager
        
        # Tkinterのルートウィンドウを取得
        self.dialog = tk.Toplevel(sql_manager.root)
        self.dialog.title(title)
        self.dialog.geometry("235x410+150+500")
        self.dialog.transient(sql_manager.root)
        self.dialog.grab_set()
        self.result = None
        self.setup_variables(current_values)
        self.create_widgets()
        
    def setup_variables(self, current_values):
        self.column_name = tk.StringVar()
        self.selected_type = tk.StringVar()
        self.length = tk.StringVar()
        self.precision = tk.StringVar()
        self.scale = tk.StringVar()
        self.is_primary = tk.BooleanVar(value=False)
        self.is_nullable = tk.BooleanVar(value=True)
        self.is_computed = tk.BooleanVar(value=False)
        self.computation_formula = tk.StringVar()
        self.is_foreign_key = tk.BooleanVar(value=False)
        self.ref_table = tk.StringVar()
        self.ref_column = tk.StringVar()
        
        if current_values:
            self.column_name.set(current_values[0])
            self.parse_data_type(current_values[1])
            self.is_primary.set(current_values[2] == "はい")
            self.is_nullable.set(current_values[3] == "はい")
            
            # 計算列の情報を設定
            if len(current_values) > 4:
                is_computed_val = current_values[4]
                computed_definition = current_values[5]
 
                # 計算列かどうかの判定を修正
                self.is_computed.set(isinstance(is_computed_val, bool) and is_computed_val or 
                                   isinstance(is_computed_val, str) and is_computed_val.upper() == 'YES')
                
                # 計算式が存在し、計算列として設定されている場合
                if computed_definition and self.is_computed.get():
                    # 計算式から余分な情報を除去
                    formula = computed_definition.strip()
                    if formula.startswith('(') and formula.endswith(')'):
                        formula = formula[1:-1].strip()
                    self.computation_formula.set(formula)
            
            if len(current_values) > 6:
                self.is_foreign_key.set(current_values[6])
                if current_values[6]:
                    self.ref_table.set(current_values[7])
                    self.ref_column.set(current_values[8])
    
    def parse_data_type(self, data_type):
        import re
        
        # DECIMAL(10,2)やVARCHAR(50)などのパターンをマッチ
        pattern = r'(\w+)(?:\((\d+)(?:,(\d+))?\))?'
        match = re.match(pattern, data_type)
        
        if match:
            base_type = match.group(1)
            self.selected_type.set(base_type)
            
            if match.group(2):
                if base_type in DataTypes.TYPES_WITH_LENGTH:
                    self.length.set(match.group(2))
                elif base_type in DataTypes.TYPES_WITH_PRECISION:
                    self.precision.set(match.group(2))
                    if match.group(3):
                        self.scale.set(match.group(3))
        
    def create_widgets(self):
        # メインフレーム
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # カラム名
        ttk.Label(main_frame, text="カラム名:").grid(row=0, column=0, padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.column_name).grid(row=0, column=1, columnspan=2, sticky='ew', padx=5, pady=5)
        
        # データ型
        ttk.Label(main_frame, text="データ型:").grid(row=1, column=0, padx=5, pady=5)
        type_combo = ttk.Combobox(main_frame, textvariable=self.selected_type, values=DataTypes.BASIC_TYPES)
        type_combo.grid(row=1, column=1, columnspan=2, sticky='ew', padx=5, pady=5)
        type_combo.bind('<<ComboboxSelected>>', self.on_type_selected)
        
        # 長さ、精度、スケール用フレーム
        self.type_details_frame = ttk.LabelFrame(main_frame, text="型の詳細", padding="5")
        self.type_details_frame.grid(row=2, column=0, columnspan=3, sticky='ew', padx=5, pady=5)
        
        # デフォルトでは非表示
        self.length_entry = ttk.Entry(self.type_details_frame, textvariable=self.length)
        self.precision_entry = ttk.Entry(self.type_details_frame, textvariable=self.precision)
        self.scale_entry = ttk.Entry(self.type_details_frame, textvariable=self.scale)
        
        # 制約フレーム
        constraints_frame = ttk.LabelFrame(main_frame, text="制約", padding="5")
        constraints_frame.grid(row=3, column=0, columnspan=3, sticky='ew', padx=5, pady=5)
        
        ttk.Checkbutton(constraints_frame, text="主キー", variable=self.is_primary).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(constraints_frame, text="NULL許可", variable=self.is_nullable).pack(side=tk.LEFT, padx=5)
        
        # 計算列フレーム
        computed_frame = ttk.LabelFrame(main_frame, text="計算列", padding="5")
        computed_frame.grid(row=4, column=0, columnspan=3, sticky='ew', padx=5, pady=5)
        
        ttk.Checkbutton(computed_frame, text="計算列として作成", 
                       variable=self.is_computed, 
                       command=self.toggle_computation_formula).pack(anchor='w', padx=5)
        
        self.formula_frame = ttk.Frame(computed_frame)
        ttk.Label(self.formula_frame, text="計算式:").pack(side=tk.LEFT)
        self.formula_entry = ttk.Entry(self.formula_frame, textvariable=self.computation_formula)
        self.formula_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # 初期状態の設定
        self.formula_frame.pack(fill=tk.X, padx=5, pady=5) if self.is_computed.get() else self.formula_frame.pack_forget()
        
        # 外部キー設定フレーム
        fk_frame = ttk.LabelFrame(main_frame, text="外部キー設定", padding="5")
        fk_frame.grid(row=5, column=0, columnspan=3, sticky='ew', padx=5, pady=5)
        
        ttk.Checkbutton(fk_frame, text="外部キーとして設定", 
                       variable=self.is_foreign_key,
                       command=self.toggle_fk_options).pack(anchor='w', padx=5)
        
        # 参照テーブル・カラム選択フレーム
        self.fk_options_frame = ttk.Frame(fk_frame)
        
        # 参照テーブル選択
        ttk.Label(self.fk_options_frame, text="参照テーブル:").pack(anchor='w', padx=5)
        self.ref_table_combo = ttk.Combobox(self.fk_options_frame, textvariable=self.ref_table)
        self.ref_table_combo.pack(fill=tk.X, padx=5, pady=2)
        self.ref_table_combo.bind('<<ComboboxSelected>>', self.on_ref_table_select)
        
        # 参照カラム選択
        ttk.Label(self.fk_options_frame, text="参照カラム:").pack(anchor='w', padx=5)
        self.ref_column_combo = ttk.Combobox(self.fk_options_frame, textvariable=self.ref_column)
        self.ref_column_combo.pack(fill=tk.X, padx=5, pady=2)
        
        # ボタン
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=3, pady=20)
        ttk.Button(button_frame, text="OK", command=self.ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="キャンセル", command=self.cancel).pack(side=tk.LEFT, padx=5)
        
        # 初期状態の設定
        self.on_type_selected(None)
        self.update_ref_tables()
        self.toggle_fk_options()
        
        # 計算列の場合は計算式入力欄を表示
        if self.is_computed.get():
            self.formula_frame.pack(fill=tk.X, padx=5, pady=5)
        else:
            self.formula_frame.pack_forget()

    def connect_to_server(self):
        return self.sql_manager.connect_to_server()
            
    def toggle_fk_options(self):
        if self.is_foreign_key.get():
            self.fk_options_frame.pack(fill=tk.X, padx=5, pady=5)
            self.update_ref_tables()
        else:
            self.fk_options_frame.pack_forget()
            
    def update_ref_tables(self):
        """参照可能なテーブル一覧を更新"""
        try:
            if not self.sql_manager.current_db:
                messagebox.showwarning("警告", "データベースが選択されていません")
                return
                
            with self.connect_to_server() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'")
                tables = [row[0] for row in cursor.fetchall()]
                self.ref_table_combo['values'] = tables
        except Exception as e:
            messagebox.showerror("エラー", f"テーブル一覧の取得に失敗しました: {str(e)}")

    def on_type_selected(self, event=None):
        selected_type = self.selected_type.get()
        
        # すべての入力フィールドを一旦非表示にする
        for widget in [self.length_entry, self.precision_entry, self.scale_entry]:
            widget.grid_forget()
        
        if selected_type in DataTypes.TYPES_WITH_LENGTH:
            # 長さを指定するデータ型の場合
            ttk.Label(self.type_details_frame, text="長さ:").grid(row=0, column=0, padx=5, pady=5)
            self.length_entry.grid(row=0, column=1, padx=5, pady=5)
            
        elif selected_type in DataTypes.TYPES_WITH_PRECISION:
            # 精度とスケールを指定するデータ型の場合
            ttk.Label(self.type_details_frame, text="精度:").grid(row=0, column=0, padx=5, pady=5)
            self.precision_entry.grid(row=0, column=1, padx=5, pady=5)
            
            ttk.Label(self.type_details_frame, text="スケール:").grid(row=1, column=0, padx=5, pady=5)
            self.scale_entry.grid(row=1, column=1, padx=5, pady=5)

    def on_ref_table_select(self, event):
        if not self.ref_table.get():
            return
            
        try:
            if not self.sql_manager.current_db:
                messagebox.showwarning("警告", "データベースが選択されていません")
                return
                
            with self.connect_to_server() as conn:
                cursor = conn.cursor()
                cursor.execute(f"""
                    SELECT COLUMN_NAME 
                    FROM INFORMATION_SCHEMA.COLUMNS 
                    WHERE TABLE_NAME = '{self.ref_table.get()}'
                    AND (COLUMNPROPERTY(OBJECT_ID(TABLE_SCHEMA + '.' + TABLE_NAME), COLUMN_NAME, 'IsIdentity') = 1
                    OR EXISTS (
                        SELECT 1 
                        FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE 
                        WHERE TABLE_NAME = '{self.ref_table.get()}' 
                        AND COLUMN_NAME = INFORMATION_SCHEMA.COLUMNS.COLUMN_NAME
                    ))
                """)
                columns = [row[0] for row in cursor.fetchall()]
                self.ref_column_combo['values'] = columns
        except Exception as e:
            messagebox.showerror("エラー", f"カラム一覧の取得に失敗しました: {str(e)}")

    def toggle_computation_formula(self):
        if self.is_computed.get():
            self.formula_frame.pack(fill=tk.X, padx=5, pady=5)
        else:
            self.formula_frame.pack_forget()
            
    def get_full_data_type(self):
        base_type = self.selected_type.get()
        if base_type in DataTypes.TYPES_WITH_LENGTH and self.length.get():
            return f"{base_type}({self.length.get()})"
        elif base_type in DataTypes.TYPES_WITH_PRECISION:
            precision = self.precision.get() or "18"
            scale = self.scale.get() or "0"
            return f"{base_type}({precision},{scale})"
        return base_type
        
    def validate(self):
        if not self.column_name.get().strip():
            messagebox.showwarning("警告", "カラム名を入力してください")
            return False
            
        if self.is_computed.get() and not self.computation_formula.get().strip():
            messagebox.showwarning("警告", "計算式を入力してください")
            return False
            
        return True
        
    def ok(self):
        if not self.validate():
            return
            
        self.result = {
            'name': self.column_name.get().strip(),
            'data_type': self.get_full_data_type(),
            'is_primary': self.is_primary.get(),
            'is_nullable': self.is_nullable.get(),
            'is_computed': self.is_computed.get(),
            'is_foreign_key' : self.is_foreign_key.get(),
            'ref_table' : self.ref_table.get(),
            'ref_column': self.ref_column.get(),
            'computation_formula': self.computation_formula.get().strip() if self.is_computed.get() else None
        }
        self.dialog.destroy()
        
    def cancel(self):
        self.dialog.destroy()

class SQLTableManager:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("SQLテーブル管理ツール")
        self.root.geometry("1100x800")
        
        # 設定ファイルのパス
        self.settings_file = "connection_settings.json"
        
        # 接続情報の読み込み
        self.connection_info = self.load_connection_settings()
        
        self.current_db = None # current_dbを初期化する
        self.current_table = None
        self.setup_ui()

    def check_and_install_driver(self):
        """SQL Server ODBCドライバの存在確認とインストール"""
        import subprocess
        import requests
        import tempfile
        import os
        import winreg

        def is_driver_installed():
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                                r"SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers", 
                                0, winreg.KEY_READ) as key:
                    i = 0
                    while True:
                        try:
                            driver = winreg.EnumValue(key, i)[0]
                            if "ODBC Driver 17 for SQL Server" in driver:
                                return True
                            i += 1
                        except WindowsError:
                            break
                return False
            except WindowsError:
                return False

        def download_and_install_driver():
            try:
                # アーキテクチャの確認
                import platform
                is_64bits = platform.machine().endswith('64')
                
                # ダウンロードURL（64bitと32bitで分岐）
                if is_64bits:
                    url = "https://go.microsoft.com/fwlink/?linkid=2239168"  # 64-bit version
                else:
                    url = "https://go.microsoft.com/fwlink/?linkid=2239169"  # 32-bit version

                # ユーザーに確認
                if not messagebox.askyesno("ドライバのインストール",
                                        "SQL Server ODBCドライバがインストールされていません。\n"
                                        "インストールしますか？"):
                    return False

                # 進捗バーダイアログの作成
                progress_dialog = tk.Toplevel(self.root)
                progress_dialog.title("ダウンロード中")
                progress_dialog.geometry("300x150")
                progress_dialog.transient(self.root)
                progress_dialog.grab_set()
                
                progress_label = ttk.Label(progress_dialog, text="ドライバをダウンロード中...")
                progress_label.pack(pady=20)
                
                progress_bar = ttk.Progressbar(progress_dialog, 
                                            length=200, 
                                            mode='indeterminate')
                progress_bar.pack(pady=10)
                progress_bar.start()

                # ダウンロードの実行
                response = requests.get(url, stream=True)
                temp_dir = tempfile.gettempdir()
                installer_path = os.path.join(temp_dir, "msodbcsql.msi")
                
                # ファイルの保存
                with open(installer_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)

                progress_dialog.destroy()

                # インストールの実行
                install_command = f'msiexec /i "{installer_path}" /quiet /norestart IACCEPTMSODBCSQLLICENSETERMS=YES'
                result = subprocess.run(install_command, shell=True)

                # インストール後の一時ファイルの削除
                try:
                    os.remove(installer_path)
                except:
                    pass

                if result.returncode == 0:
                    messagebox.showinfo("成功", 
                                    "SQL Server ODBCドライバのインストールが完了しました。")
                    return True
                else:
                    messagebox.showerror("エラー", 
                                    "ドライバのインストールに失敗しました。")
                    return False

            except Exception as e:
                messagebox.showerror("エラー", 
                                f"ドライバのインストール中にエラーが発生しました: {str(e)}")
                return False

        if not is_driver_installed():
            return download_and_install_driver()
        return True

    def load_connection_settings(self):
        default_settings = {
            'server': 'NTC18',
            'username': 'sa',
            'password': 'Ntc002611',
            'driver': 'ODBC Driver 17 for SQL Server'
        }
        
        # ドライバのチェックとインストール
        if not self.check_and_install_driver():
            if messagebox.askyesno("警告", 
                                "ドライバがインストールされていません。\n"
                                "代替のドライバで続行しますか？"):
                default_settings['driver'] = 'SQL Server'  # 代替ドライバを使用
            else:
                self.root.quit()
                return None
        
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    return json.load(f)
        except:
            pass
            
        return default_settings
        
    def save_connection_settings(self, settings):
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f)
        except Exception as e:
            messagebox.showerror("エラー", f"設定の保存に失敗しました: {str(e)}")

    def setup_ui(self):
        # メニューバーの追加
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 設定メニュー
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="設定", menu=settings_menu)
        settings_menu.add_command(label="接続設定", command=self.show_connection_settings)
        
        # メインフレームの作成
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 左側のフレーム（データベースとテーブル選択用）
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)

        # データベース選択部分
        db_frame = ttk.LabelFrame(left_frame, text="データベース設定")
        db_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(db_frame, text="データベース名:").pack(padx=5, pady=2)
        self.db_entry = ttk.Entry(db_frame)
        self.db_entry.pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(db_frame, text="登 録", command=self.register_database).pack(padx=5, pady=2)
        
        self.db_listbox = tk.Listbox(db_frame, height=5)
        self.db_listbox.pack(fill=tk.X, padx=5, pady=5)
        self.db_listbox.bind('<<ListboxSelect>>', self.on_db_select)

        # テーブル選択部分
        table_frame = ttk.LabelFrame(left_frame, text="テーブル設定")
        table_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(table_frame, text="テーブル名:").pack(padx=5, pady=2)
        self.table_entry = ttk.Entry(table_frame)
        self.table_entry.pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(table_frame, text="登 録", command=self.register_table).pack(padx=5, pady=2)

        self.table_listbox = tk.Listbox(table_frame, height=5)
        self.table_listbox.pack(fill=tk.X, padx=5, pady=5)
        self.table_listbox.bind('<<ListboxSelect>>', self.on_table_select)

        # カラム設定部分
        column_frame = ttk.LabelFrame(left_frame, text="カラム設定")
        column_frame.pack(fill=tk.X, padx=5, pady=5)

        # カラム設定のグリッドレイアウト
        settings_frame = ttk.Frame(column_frame)
        settings_frame.pack(fill=tk.X, padx=5, pady=5)

        # ボタンフレーム
        button_add = ttk.Frame(settings_frame)
        button_add.grid(row=1, column=0, columnspan=2, pady=10)
        ttk.Button(button_add, text="カラム 追加", command=self.add_column).pack(side=tk.LEFT, padx=5)
        button_edit = ttk.Frame(settings_frame)
        button_edit.grid(row=2, column=0, columnspan=2, pady=10)
        ttk.Button(button_edit, text="カラム 編集", command=self.edit_column).pack(side=tk.LEFT, padx=5)
        button_del = ttk.Frame(settings_frame)
        button_del.grid(row=3, column=0, columnspan=2, pady=10)
        ttk.Button(button_del, text="カラム 削除", command=self.delete_column).pack(side=tk.LEFT, padx=5)
        
        # 右側のフレーム（カラム設定用）
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        # カラム一覧表示
        self.column_tree = ttk.Treeview(right_frame, columns=("名前", "型", "主キー", "NULL許可"), show="headings")
        self.column_tree.heading("名前", text="カラム名")
        self.column_tree.heading("型", text="データ型")
        self.column_tree.heading("主キー", text="主キー")
        self.column_tree.heading("NULL許可", text="NULL許可")
        self.column_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 初期状態の設定
        self.refresh_database_list()

    def show_connection_settings(self):
        dialog = ConnectionSettingsDialog(self.root, self.connection_info)
        dialog.dialog.wait_window()
        
        if dialog.result:
            self.connection_info = dialog.result
            self.save_connection_settings(dialog.result)
            # 接続情報が変更されたので、データベース一覧を更新
            self.refresh_database_list()

    def connect_to_server(self):
        connection_string = (
            f"DRIVER={{{self.connection_info['driver']}}};"
            f"SERVER={self.connection_info['server']};"
            f"UID={self.connection_info['username']};"
            f"PWD={self.connection_info['password']}"
        )
        if self.current_db:
            connection_string += f";DATABASE={self.current_db}"
        return pyodbc.connect(connection_string)

    def refresh_database_list(self):
        try:
            with self.connect_to_server() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sys.databases WHERE database_id > 4")
                self.db_listbox.delete(0, tk.END)
                for db in cursor.fetchall():
                    self.db_listbox.insert(tk.END, db[0])
        except Exception as e:
            error_type = type(e).__name__
            error_message = str(e)
            messagebox.showerror("エラー", f"データベース一覧の取得に失敗しました: {str(e)}: {error_type}: {error_message}")

    def refresh_table_list(self):
        if not self.current_db:
            return
        try:
            with self.connect_to_server() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'")
                self.table_listbox.delete(0, tk.END)
                for table in cursor.fetchall():
                    self.table_listbox.insert(tk.END, table[0])
        except Exception as e:
            messagebox.showerror("エラー", f"テーブル一覧の取得に失敗しました: {str(e)}")

    def register_database(self):
        db_name = self.db_entry.get().strip()
        if not db_name:
            messagebox.showwarning("警告", "データベース名を入力してください")
            return

        try:
            with self.connect_to_server() as conn:
                cursor = conn.cursor()
                cursor.execute(f"CREATE DATABASE {db_name}")
                conn.commit()
            messagebox.showinfo("成功", f"データベース '{db_name}' を作成しました")
            self.refresh_database_list()
            self.db_entry.delete(0, tk.END)
        except Exception as e:
            messagebox.showerror("エラー", f"データベースの作成に失敗しました: {str(e)}")

    def register_table(self):
        if not self.current_db:
            messagebox.showwarning("警告", "データベースを選択してください")
            return

        table_name = self.table_entry.get().strip()
        if not table_name:
            messagebox.showwarning("警告", "テーブル名を入力してください")
            return

        try:
            with self.connect_to_server() as conn:
                cursor = conn.cursor()
                # IDENTITYを追加してAUTO_INCREMENTを実現
                cursor.execute(f"CREATE TABLE {table_name} (ID INT IDENTITY(1,1) PRIMARY KEY)")
                conn.commit()
            messagebox.showinfo("成功", f"テーブル '{table_name}' を作成しました")
            self.refresh_table_list()
            self.table_entry.delete(0, tk.END)
        except Exception as e:
            messagebox.showerror("エラー", f"テーブルの作成に失敗しました: {str(e)}")

    def add_column(self):
        if not self.current_table:
            messagebox.showwarning("警告", "テーブルを選択してください")
            return

        dialog = ColumnDialog(self, "カラムの追加")  # SQLTableManagerのインスタンスを渡す
        dialog.dialog.wait_window()
        
        if not dialog.result:
            return
            
        try:
            with self.connect_to_server() as conn:
                cursor = conn.cursor()
                
                if dialog.result['is_computed']:
    # 計算列の追加
                    sql = f"""ALTER TABLE {self.current_table} 
                            ADD {dialog.result['name']} AS {dialog.result['computation_formula']}"""
                else:
                    # 通常のカラム追加
                    # 主キーで数値型の場合はIDENTITYを追加
                    data_type = dialog.result['data_type']
                    is_numeric = any(type in data_type.upper() for type in ['INT', 'BIGINT', 'SMALLINT', 'TINYINT'])
                    
                    if dialog.result['is_primary'] and is_numeric:
                        # IDENTITY(1,1)を追加：開始値1, 増分値1
                        identity_clause = "IDENTITY(1,1)"
                    else:
                        identity_clause = ""
                    
                    constraint = "PRIMARY KEY" if dialog.result['is_primary'] else ""
                    null_constraint = "NULL" if dialog.result['is_nullable'] else "NOT NULL"
                    
                    sql = f"""ALTER TABLE {self.current_table} 
                            ADD {dialog.result['name']} {data_type} {identity_clause} {constraint} {null_constraint}"""
                    
                    cursor.execute(sql)
                    conn.commit()

                    # 外部キーの設定
                    if dialog.result.get('is_foreign_key', False):
                        ref_table = dialog.result.get('ref_table')
                        ref_column = dialog.result.get('ref_column')
                        if ref_table and ref_column:
                            fk_constraint_name = f"FK_{self.current_table}_{dialog.result['name']}"
                            cursor.execute(f"""
                                ALTER TABLE {self.current_table}
                                ADD CONSTRAINT {fk_constraint_name}
                                FOREIGN KEY ({dialog.result['name']})
                                REFERENCES {ref_table}({ref_column})
                            """)
                            conn.commit()

                
                self.column_tree.insert("", tk.END, values=(
                    dialog.result['name'],
                    dialog.result['data_type'],
                    "はい" if dialog.result['is_primary'] else "いいえ",
                    "はい" if dialog.result['is_nullable'] else "いいえ"
                ))
                
                messagebox.showinfo("成功", f"カラム '{dialog.result['name']}' を追加しました")
                
        except Exception as e:
            messagebox.showerror("エラー", f"カラムの追加に失敗しました: {str(e)}")

    def delete_column(self):
        if not self.current_table:
            messagebox.showwarning("警告", "データベース・テーブルを選択し、\n削除するカラムを選択してください")
            return

        selected_item = self.column_tree.selection()
        if not selected_item:
            messagebox.showwarning("警告", "削除するカラムを選択してください")
            return

        column_name = self.column_tree.item(selected_item)['values'][0]
        if messagebox.askyesno("確認", f"カラム '{column_name}' を削除しますか？"):
            try:
                with self.connect_to_server() as conn:
                    cursor = conn.cursor()
                    cursor.execute(f"ALTER TABLE {self.current_table} DROP COLUMN {column_name}")
                    conn.commit()
                self.column_tree.delete(selected_item)
                messagebox.showinfo("成功", f"カラム '{column_name}' を削除しました")
            except Exception as e:
                messagebox.showerror("エラー", f"カラムの削除に失敗しました: {str(e)}")

    def on_db_select(self, event):
        selection = self.db_listbox.curselection()
        if selection:
            self.current_db = self.db_listbox.get(selection[0])
            self.refresh_table_list()
            self.current_table = None
            self.column_tree.delete(*self.column_tree.get_children())

    def on_table_select(self, event):
        selection = self.table_listbox.curselection()
        if selection:
            self.current_table = self.table_listbox.get(selection[0])
            self.refresh_column_list()

    def refresh_column_list(self):
        if not self.current_table:
            return

        try:
            with self.connect_to_server() as conn:
                cursor = conn.cursor()
                cursor.execute(f"""
                    SELECT 
                        c.COLUMN_NAME, 
                        c.DATA_TYPE,
                        CASE WHEN COLUMNPROPERTY(OBJECT_ID(c.TABLE_SCHEMA + '.' + c.TABLE_NAME), c.COLUMN_NAME, 'IsIdentity') = 1 
                             OR EXISTS (
                                SELECT 1 FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE 
                                WHERE TABLE_NAME = '{self.current_table}' 
                                AND COLUMN_NAME = c.COLUMN_NAME
                             ) THEN 'はい' ELSE 'いいえ' END as IS_PRIMARY,
                        CASE WHEN c.IS_NULLABLE = 'YES' THEN 'はい' ELSE 'いいえ' END as IS_NULLABLE,
                        CASE WHEN cc.is_computed = 1 THEN 'YES' ELSE 'NO' END as IS_COMPUTED,
                        cc.definition as COMPUTED_DEFINITION  -- cc.definition を使用
                    FROM INFORMATION_SCHEMA.COLUMNS c
                    LEFT JOIN sys.columns sc 
                        ON OBJECT_ID(c.TABLE_SCHEMA + '.' + c.TABLE_NAME) = sc.object_id 
                        AND c.COLUMN_NAME = sc.name
                    LEFT JOIN sys.computed_columns cc 
                        ON sc.object_id = cc.object_id 
                        AND sc.column_id = cc.column_id
                    WHERE c.TABLE_NAME = '{self.current_table}'
                """)
                
                self.column_tree.delete(*self.column_tree.get_children())
                self.columns_data = []  # カラムデータを保存するリストをクリア
                
                for column in cursor.fetchall():
                    # タプルの各要素を個別に取得して表示
                    column_name = column[0]
                    data_type = column[1]
                    is_primary = column[2]
                    is_nullable = column[3]
                    is_computed = column[4]
                    computed_definition = column[5]

                    # カラムデータを保存
                    self.columns_data.append({
                        'name': column_name,
                        'data_type': data_type,
                        'is_primary': is_primary == 'はい',
                        'is_nullable': is_nullable == 'はい',
                        'is_computed': is_computed == 'YES',
                        'computed_definition': computed_definition if computed_definition else None
                    })
                    
                    self.column_tree.insert("", tk.END, values=(column_name, data_type, is_primary, is_nullable))

        except Exception as e:
            messagebox.showerror("エラー", f"カラム一覧の取得に失敗しました: {str(e)}")

    def edit_column(self):
        selected_item = self.column_tree.selection()
        if not selected_item:
            messagebox.showwarning("警告", "編集するカラムを選択してください")
            return

        # 選択されたアイテムのインデックスを取得
        selected_index = self.column_tree.index(selected_item)
        if selected_index >= len(self.columns_data):
            messagebox.showerror("エラー", "カラム情報の取得に失敗しました")
            return

        # 保存されているカラムデータを取得
        column_data = self.columns_data[selected_index]
        current_values = [
            column_data['name'],
            column_data['data_type'],
            "はい" if column_data['is_primary'] else "いいえ",
            "はい" if column_data['is_nullable'] else "いいえ",
            column_data['is_computed'],
            column_data['computed_definition']
        ]
        
        dialog = ColumnDialog(self, "カラムの編集", current_values)
        dialog.dialog.wait_window()
        
        if not dialog.result:
            return
            
        try:
            with self.connect_to_server() as conn:
                cursor = conn.cursor()
                old_column_name = current_values[0]
                
                # カラム名の変更
                if dialog.result['name'] != old_column_name:
                    cursor.execute(f"EXEC sp_rename '{self.current_table}.{old_column_name}', '{dialog.result['name']}', 'COLUMN'")
                
                # 計算列への変更または通常カラムへの変更
                if dialog.result['is_computed']:
                    sql = f"""ALTER TABLE {self.current_table} DROP COLUMN {dialog.result['name']}
                            ALTER TABLE {self.current_table} ADD {dialog.result['name']} 
                            AS {dialog.result['computation_formula']}"""
                else:
                    # データ型と制約の変更
                    null_constraint = "NULL" if dialog.result['is_nullable'] else "NOT NULL"
                    sql = f"""ALTER TABLE {self.current_table} ALTER COLUMN {dialog.result['name']} 
                            {dialog.result['data_type']} {null_constraint}"""
                
                cursor.execute(sql)
                
                # 主キー制約の変更
                if dialog.result['is_primary'] != (current_values[2] == "はい"):
                    if dialog.result['is_primary']:
                        cursor.execute(f"ALTER TABLE {self.current_table} ADD CONSTRAINT PK_{dialog.result['name']} PRIMARY KEY ({dialog.result['name']})")
                    else:
                        cursor.execute(f"ALTER TABLE {self.current_table} DROP CONSTRAINT PK_{old_column_name}")
                
                conn.commit()
                self.refresh_column_list()
                messagebox.showinfo("成功", "カラムを更新しました")
                
        except Exception as e:
            messagebox.showerror("エラー", f"カラムの更新に失敗しました: {str(e)}")

    def run(self):
        self.root.mainloop()

class DataTypes:
    # 長さを指定できるデータ型
    TYPES_WITH_LENGTH = [
        'VARCHAR', 'NVARCHAR', 'CHAR',
        'NCHAR', 'BINARY', 'VARBINARY'
    ]
    
    # 精度とスケールを指定できるデータ型
    TYPES_WITH_PRECISION = [
        'DECIMAL', 'NUMERIC'
    ]
    
    # 基本的なデータ型一覧
    BASIC_TYPES = [
        'INT', 'BIGINT', 'SMALLINT',
        'TINYINT', 'BIT', 'DECIMAL',
        'NUMERIC', 'MONEY', 'SMALLMONEY',
        'FLOAT', 'REAL', 'DATE',
        'TIME', 'DATETIME', 'DATETIME2',
        'DATETIMEOFFSET', 'SMALLDATETIME',
        'CHAR', 'VARCHAR', 'TEXT',
        'NCHAR', 'NVARCHAR', 'NTEXT',
        'BINARY', 'VARBINARY', 'IMAGE',
        'UNIQUEIDENTIFIER'
    ]

def main():
    app = SQLTableManager()
    app.run()

if __name__ == "__main__":
    main()