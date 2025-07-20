import tkinter as tk
# ttkとcsvモジュールをインポート
from tkinter import filedialog, messagebox, ttk
import json
import csv
import os
import re
import sys
import subprocess
# モダンなテーマを適用するためにttkthemesをインポート
try:
    from ttkthemes import ThemedTk
except ImportError:
    # ライブラリがない場合はNoneにしておく
    ThemedTk = None

# 設定ファイル名
CONFIG_FILE = "renamer_config.json"

class TreeviewToolTip:
    """Creates a tooltip for a given Treeview widget."""
    def __init__(self, treeview, delay=500):
        self.tree = treeview
        self.delay = delay
        self.tip_window = None
        self.after_id = None
        self.last_row = None
        self.last_col = None

        self.tree.bind('<Motion>', self.schedule_tip)
        self.tree.bind('<Leave>', '>')

    def schedule_tip(self, event):
        """Schedules the tooltip to appear after a delay."""
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)

        if not row_id or not col_id:
            self.hide_tip()
            return

        if row_id == self.last_row and col_id == self.last_col:
            return

        self.hide_tip()

        self.last_row = row_id
        self.last_col = col_id

        self.after_id = self.tree.after(self.delay, lambda: self.show_tip(event))

    def show_tip(self, event):
        """Displays the tooltip."""
        if not self.last_row or not self.last_col:
            return

        col_index = int(self.last_col.replace('#', '')) - 1
        column_id = self.tree['columns'][col_index]
        cell_text = self.tree.set(self.last_row, column_id)

        if not cell_text:
            return

        self.tip_window = tw = tk.Toplevel(self.tree)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{event.x_root + 20}+{event.y_root + 10}")

        label = tk.Label(tw, text=cell_text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hide_tip(self, event=None):
        """Hides the tooltip."""
        if self.after_id:
            self.tree.after_cancel(self.after_id)
            self.after_id = None

        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None
        
        self.last_row = None
        self.last_col = None

class RenamerApp:
    def __init__(self, master):
        """
        RenamerAppクラスのコンストラクタ。
        GUIの初期設定とウィジェットの作成を行います。
        """
        self.master = master
        master.title("ファイル・フォルダ改名アプリ") # アプリケーションのタイトルを設定

        # 選択されたフォルダのパスを保持するStringVar
        self.selected_folder_path = tk.StringVar()
        self.selected_folder_path.set("フォルダが選択されていません") # 初期メッセージ

        # 解析結果を保持するリスト
        self.analysis_results = []

        # リネーム履歴を保持するリスト
        self.rename_history = []

        # 置換対象の記号を保持するStringVar
        self.custom_symbols_to_replace = tk.StringVar(value="")

        # 置換文字を保持するStringVar
        self.replacement_char = tk.StringVar(value="_")

        # Treeviewソート用の変数
        self.treeview_headers = {
            "type": "種類",
            "original": "変更前の名前",
            "new": "変更後の名前",
            "path": "場所"
        }
        self.treeview_sort_column = None
        self.treeview_sort_reverse = False

        self.create_widgets()
        self.load_config()

    def create_widgets(self):
        """
        GUIウィジェットを作成し、配置します。
        Frameを使用してレイアウトを左右に分割します。
        """
        # メインフレーム
        main_frame = tk.Frame(self.master)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # --- 左側のコントロールフレーム ---
        left_frame = tk.Frame(main_frame, width=280)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left_frame.pack_propagate(False) # 幅を固定

        # --- 右側の結果表示フレーム ---
        right_frame = tk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # --- 左側ウィジェット ---
        # 選択されたフォルダパスを表示するラベル
        self.path_label = tk.Label(left_frame, textvariable=self.selected_folder_path, wraplength=270, justify=tk.LEFT)
        self.path_label.pack(pady=10, anchor='w')

        # フォルダ選択ボタン
        self.select_button = tk.Button(left_frame, text="フォルダを選択", command=self.select_folder)
        self.select_button.pack(pady=5, fill=tk.X)

        # 設定ボタン
        self.settings_button = tk.Button(left_frame, text="置換対象の設定", command=self.open_settings_window)
        self.settings_button.pack(pady=5, fill=tk.X)

        # --- 置換文字設定 ---
        replace_frame = tk.Frame(left_frame)
        replace_frame.pack(pady=5, fill=tk.X)
        replace_label = tk.Label(replace_frame, text="置換後の文字:")
        replace_label.pack(side=tk.LEFT, padx=(0, 5))
        self.replace_entry = tk.Entry(replace_frame, textvariable=self.replacement_char, width=10)
        self.replace_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 解析ボタン
        self.analyze_button = tk.Button(left_frame, text="解析 (プレビュー)", command=self.analyze_items, state=tk.DISABLED)
        self.analyze_button.pack(pady=5, fill=tk.X)

        # CSV出力ボタン
        self.export_csv_button = tk.Button(left_frame, text="CSVに出力", command=self.export_to_csv, state=tk.DISABLED)
        self.export_csv_button.pack(pady=5, fill=tk.X)

        # 実行ボタン
        self.execute_button = tk.Button(left_frame, text="リネーム実行", command=self.rename_items, state=tk.DISABLED)
        self.execute_button.pack(pady=20, fill=tk.X)

        # 切り戻しボタン (初期状態では無効)
        self.revert_button = tk.Button(left_frame, text="リネームを元に戻す", command=self.revert_rename, state=tk.DISABLED)
        self.revert_button.pack(pady=5, fill=tk.X)

        # プログレスバー (初期状態では非表示)
        self.progress_bar = ttk.Progressbar(left_frame, orient="horizontal", mode="determinate")
        # pack() は rename_items の中で実行時に行います

        # ログ表示エリア
        log_label = tk.Label(left_frame, text="ログ:")
        log_label.pack(pady=(10,0), anchor='w')
        self.log_text = tk.Text(left_frame, height=10, state=tk.DISABLED)
        self.log_text.pack(pady=5, fill=tk.BOTH, expand=True)
        self.log_message("ログ:\n")

        # --- 右側ウィジェット (解析結果リスト) ---
        result_label = tk.Label(right_frame, text="解析結果:")
        result_label.pack(pady=(10,0), anchor='w')

        tree_frame = tk.Frame(right_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.tree = ttk.Treeview(tree_frame, columns=tuple(self.treeview_headers.keys()), show="headings")
        for col, text in self.treeview_headers.items():
            self.tree.heading(col, text=text, anchor='w',
                              command=lambda _col=col: self.sort_treeview_column(_col))

        self.tree.column("type", width=80, anchor='w')
        self.tree.column("original", width=200, anchor='w')
        self.tree.column("new", width=200, anchor='w')
        self.tree.column("path", width=300, anchor='w')

        # スクロールバー
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Treeviewのダブルクリックイベントをバインド
        self.tree.bind("<Double-1>", self.on_tree_double_click)

        # Treeviewにツールチップ機能を追加
        TreeviewToolTip(self.tree)

    def log_message(self, message):
        """
        ログメッセージをテキストエリアに追記します。
        """
        self.log_text.config(state=tk.NORMAL) # 書き込み可能にする
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END) # 最新のメッセージまでスクロール
        self.log_text.config(state=tk.DISABLED) # 読み取り専用に戻す

    def clear_results(self):
        """解析結果のリストとTreeviewをクリアします。"""
        self.analysis_results.clear()
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.export_csv_button.config(state=tk.DISABLED)
        # ソート状態とヘッダーの矢印をリセット
        if hasattr(self, 'treeview_headers'):
             for c, text in self.treeview_headers.items():
                self.tree.heading(c, text=text)
        self.treeview_sort_column = None
        self.treeview_sort_reverse = False

    def select_folder(self):
        """
        フォルダ選択ダイアログを開き、選択されたフォルダパスを更新します。
        フォルダが選択された場合、実行ボタンを有効にします。
        """
        folder_selected = filedialog.askdirectory()
        self.clear_results() # 選択時に以前の結果をクリア
        self.rename_history.clear() # 履歴をクリア
        self.revert_button.config(state=tk.DISABLED) # 切り戻しボタンを無効化
        if folder_selected:
            self.selected_folder_path.set(folder_selected)
            self.analyze_button.config(state=tk.NORMAL) # フォルダが選択されたら解析ボタンを有効化
            self.execute_button.config(state=tk.NORMAL) # フォルダが選択されたら実行ボタンを有効化
            self.log_message(f"選択されたフォルダ: {folder_selected}\n")
            self.save_config()
        else:
            self.selected_folder_path.set("フォルダが選択されていません")
            self.analyze_button.config(state=tk.DISABLED) # 選択がキャンセルされたら解析ボタンを無効化
            self.execute_button.config(state=tk.DISABLED) # 選択がキャンセルされたら実行ボタンを無効化
            self.log_message("フォルダ選択がキャンセルされました。\n")

    def set_buttons_state(self, state):
        """主要なボタンの状態を一括で設定します。"""
        self.select_button.config(state=state)
        self.settings_button.config(state=state)
        self.analyze_button.config(state=state)
        self.execute_button.config(state=state)
        self.revert_button.config(state=state)
        self.replace_entry.config(state=state)
        # CSVボタンと切り戻しボタンは特定の条件下でのみ有効になるため、
        # stateがDISABLEDの場合にのみ一括で無効化する
        if state == tk.DISABLED:
            self.export_csv_button.config(state=tk.DISABLED)

    def open_settings_window(self):
        """
        置換対象の記号を設定するためのウィンドウを開きます。
        """
        settings_win = tk.Toplevel(self.master)
        settings_win.title("置換対象の設定")
        settings_win.geometry("400x150")
        settings_win.transient(self.master) # 親ウィンドウの上に表示
        settings_win.grab_set() # モーダルウィンドウにする

        # 設定ウィンドウ内でのみ使用する一時的な変数
        temp_custom_symbols = tk.StringVar(value=self.custom_symbols_to_replace.get())

        main_label = tk.Label(settings_win, text="絵文字に加えて置換したい記号を続けて入力してください。")
        main_label.pack(pady=10, padx=10)

        entry_frame = tk.Frame(settings_win)
        entry_frame.pack(pady=5, padx=20, fill=tk.X)

        symbols_entry = tk.Entry(entry_frame, textvariable=temp_custom_symbols)
        symbols_entry.pack(fill=tk.X, expand=True)

        def save_and_close():
            self.custom_symbols_to_replace.set(temp_custom_symbols.get())
            self.log_message(f"追加の置換対象記号が更新されました: '{temp_custom_symbols.get()}'\n")
            settings_win.destroy()

        button_frame = tk.Frame(settings_win)
        button_frame.pack(pady=20)

        save_button = tk.Button(button_frame, text="保存して閉じる", command=save_and_close)
        save_button.pack(side=tk.LEFT, padx=10)

        cancel_button = tk.Button(button_frame, text="キャンセル", command=settings_win.destroy)
        cancel_button.pack(side=tk.LEFT, padx=10)

    def analyze_items(self):
        """
        選択されたフォルダ内のファイルとフォルダを解析し、
        リネーム対象となるアイテムの一覧を右側のリストに表示します。
        """
        root_dir = self.selected_folder_path.get()
        if not os.path.isdir(root_dir):
            messagebox.showerror("エラー", "有効なフォルダが選択されていません。")
            return

        self.log_message(f"\n--- フォルダ '{root_dir}' の解析を開始します ---\n")
        self.set_buttons_state(tk.DISABLED)
        self.rename_history.clear() # 再解析時に履歴をクリア
        self.revert_button.config(state=tk.DISABLED) # 切り戻しボタンを無効化
        self.clear_results()

        try:
            # 置換文字を取得（空の場合は"_"を使用）
            replacement = self.replacement_char.get()
            if not replacement:
                replacement = "_"

            # os.walkでファイルとフォルダを探索
            for dirpath, dirnames, filenames in os.walk(root_dir):
                for filename in filenames:
                    new_filename = self.replace_invalid_chars(filename, replacement)
                    if new_filename != filename:
                        result = {'type': 'ファイル', 'original': filename, 'new': new_filename, 'path': dirpath}
                        self.analysis_results.append(result)
                        self.tree.insert("", tk.END, values=(result['type'], result['original'], result['new'], result['path']))
                for dirname in dirnames:
                    new_dirname = self.replace_invalid_chars(dirname, replacement)
                    if new_dirname != dirname:
                        result = {'type': 'フォルダ', 'original': dirname, 'new': new_dirname, 'path': dirpath}
                        self.analysis_results.append(result)
                        self.tree.insert("", tk.END, values=(result['type'], result['original'], result['new'], result['path']))

            # ルートフォルダ自体の解析
            root_basename = os.path.basename(root_dir)
            new_root_basename = self.replace_invalid_chars(root_basename, replacement)
            if new_root_basename != root_basename:
                result = {'type': 'ルートフォルダ', 'original': root_basename, 'new': new_root_basename, 'path': os.path.dirname(root_dir)}
                self.analysis_results.append(result)
                self.tree.insert("", tk.END, values=(result['type'], result['original'], result['new'], result['path']))

            if self.analysis_results:
                self.log_message(f"{len(self.analysis_results)} 件のリネーム対象が見つかりました。\n")
                self.export_csv_button.config(state=tk.NORMAL) # 結果があればCSV出力ボタンを有効化
            else:
                self.log_message("リネーム対象のアイテムは見つかりませんでした。\n")

            self.log_message(f"--- 解析が完了しました ---\n")
            messagebox.showinfo("完了", "解析が完了しました。右側のリストを確認してください。")
        except Exception as e:
            self.log_message(f"解析中に予期せぬエラーが発生しました: {e}\n")
            messagebox.showerror("エラー", f"解析中に予期せぬエラーが発生しました: {e}")
        finally:
            self.set_buttons_state(tk.NORMAL)
            # フォルダが選択されていなければ、解析・実行ボタンは無効のままにする
            if not os.path.isdir(self.selected_folder_path.get()):
                self.analyze_button.config(state=tk.DISABLED)
                self.execute_button.config(state=tk.DISABLED)

    def rename_items(self):
        """
        選択されたフォルダ内のファイルとフォルダの名前を絵文字を指定文字に変換して変更します。
        処理の進捗はプログレスバーで表示されます。
        """
        root_dir = self.selected_folder_path.get()
        if not os.path.isdir(root_dir):
            messagebox.showerror("エラー", "有効なフォルダが選択されていません。")
            return

        self.log_message(f"\n--- フォルダ '{root_dir}' の処理を開始します ---\n")
        self.set_buttons_state(tk.DISABLED)
        # 実行前に以前の履歴をクリアし、切り戻しボタンを無効化
        self.rename_history.clear()
        self.revert_button.config(state=tk.DISABLED)

        # プログレスバーをリセットして表示
        self.progress_bar.pack(pady=(0, 10), fill=tk.X)
        self.progress_bar['value'] = 0

        try:
            # 置換文字を取得（空の場合は"_"を使用）
            replacement = self.replacement_char.get() or "_"

            # --- リネーム対象のリストアップ ---
            # os.walkを一度だけ実行して、ファイルとフォルダのリストを効率的に作成します。
            files_to_rename = []
            dirs_to_process = []
            for dirpath, dirnames, filenames in os.walk(root_dir):
                for filename in filenames:
                    if self.replace_invalid_chars(filename, replacement) != filename:
                        files_to_rename.append(os.path.join(dirpath, filename))
                for dirname in dirnames:
                    # フォルダは後で深い順にソートするため、一旦すべて追加
                    dirs_to_process.append(os.path.join(dirpath, dirname))

            # フォルダをパスの長さで降順にソートし、深い階層から処理
            dirs_to_process.sort(key=len, reverse=True)
            dirs_to_rename = [d for d in dirs_to_process if self.replace_invalid_chars(os.path.basename(d), replacement) != os.path.basename(d)]

            # ルートフォルダが対象かチェック
            root_basename = os.path.basename(root_dir)
            is_root_to_rename = self.replace_invalid_chars(root_basename, replacement) != root_basename

            # --- プログレスバーの設定 ---
            total_items = len(files_to_rename) + len(dirs_to_rename) + (1 if is_root_to_rename else 0)
            if total_items == 0:
                self.log_message("リネーム対象のアイテムは見つかりませんでした。\n")
                messagebox.showinfo("情報", "リネーム対象のアイテムはありませんでした。")
                return

            self.progress_bar['maximum'] = total_items
            self.master.update_idletasks()

            # --- 処理実行 ---
            # Step 1: ファイルのリネーム
            for original_file_path in files_to_rename:
                try:
                    dirname, basename = os.path.split(original_file_path)
                    new_filename = self.replace_invalid_chars(basename, replacement)
                    new_file_path = os.path.join(dirname, new_filename)
                    os.rename(original_file_path, new_file_path)
                    self.rename_history.append({'original': original_file_path, 'new': new_file_path})
                    self.log_message(f"ファイル改名: '{basename}' -> '{new_filename}'\n")
                except FileNotFoundError:
                    self.log_message(f"警告: ファイル '{original_file_path}' は見つかりませんでした。スキップします。\n")
                except Exception as e:
                    self.log_message(f"エラー: ファイル '{original_file_path}' の改名中にエラーが発生しました: {e}\n")
                finally:
                    self.progress_bar.step()
                    self.master.update_idletasks()

            # Step 2: フォルダのリネーム
            for original_dir_path in dirs_to_rename:
                try:
                    if os.path.exists(original_dir_path):
                        parent_dir, current_dir_name = os.path.split(original_dir_path)
                        new_dir_name = self.replace_invalid_chars(current_dir_name, replacement)
                        new_dir_path = os.path.join(parent_dir, new_dir_name)
                        os.rename(original_dir_path, new_dir_path)
                        self.rename_history.append({'original': original_dir_path, 'new': new_dir_path})
                        self.log_message(f"フォルダ改名: '{current_dir_name}' -> '{new_dir_name}'\n")
                except Exception as e:
                    self.log_message(f"エラー: フォルダ '{original_dir_path}' の改名中にエラーが発生しました: {e}\n")
                finally:
                    self.progress_bar.step()
                    self.master.update_idletasks()

            # Step 3: ルートフォルダのリネーム
            if is_root_to_rename:
                try:
                    parent_of_root = os.path.dirname(root_dir)
                    new_root_basename = self.replace_invalid_chars(root_basename, replacement)
                    new_root_path = os.path.join(parent_of_root, new_root_basename)
                    os.rename(root_dir, new_root_path)
                    self.rename_history.append({'original': root_dir, 'new': new_root_path})
                    self.log_message(f"ルートフォルダ改名: '{root_basename}' -> '{new_root_basename}'\n")
                    self.selected_folder_path.set(new_root_path)
                    self.save_config()
                except Exception as e:
                    self.log_message(f"エラー: ルートフォルダ '{root_dir}' の改名中にエラーが発生しました: {e}\n")
                finally:
                    self.progress_bar.step()
                    self.master.update_idletasks()

            self.log_message(f"--- フォルダ '{root_dir}' の処理が完了しました ---\n")
            messagebox.showinfo("完了", "ファイルとフォルダの改名が完了しました。")
            self.clear_results() # 実行後は解析結果をクリア

        except Exception as e:
            self.log_message(f"予期せぬエラーが発生しました: {e}\n")
            messagebox.showerror("エラー", f"予期せぬエラーが発生しました: {e}")
        finally:
            self.progress_bar.pack_forget() # 処理完了後、プログレスバーを非表示にする
            self.set_buttons_state(tk.NORMAL)
            if self.rename_history:
                self.revert_button.config(state=tk.NORMAL) # 履歴があれば切り戻しボタンを有効化
            # フォルダが選択されていなければ、解析・実行ボタンは無効のままにする
            if not os.path.isdir(self.selected_folder_path.get()):
                self.analyze_button.config(state=tk.DISABLED)
                self.execute_button.config(state=tk.DISABLED)

    def revert_rename(self):
        """
        直前のリネーム処理を元に戻します。
        """
        if not self.rename_history:
            messagebox.showinfo("情報", "元に戻すリネーム履歴がありません。")
            return

        if not messagebox.askyesno("確認", f"{len(self.rename_history)}件のリネームを元に戻しますか？"):
            return

        self.log_message("\n--- リネームの切り戻し処理を開始します ---\n")
        self.set_buttons_state(tk.DISABLED)

        # プログレスバーの準備
        self.progress_bar.pack(pady=(0, 10), fill=tk.X)
        self.progress_bar['value'] = 0
        self.progress_bar['maximum'] = len(self.rename_history)
        self.master.update_idletasks()

        try:
            # 履歴を逆順に処理して、ルート -> フォルダ -> ファイルの順で安全に元に戻す
            for item in reversed(self.rename_history):
                original_path = item['original']
                new_path = item['new']
                try:
                    if os.path.exists(new_path) or os.path.isdir(new_path):
                        os.rename(new_path, original_path)
                        # もし現在の選択パスが切り戻し対象のパスだったら、表示を更新する
                        if self.selected_folder_path.get() == new_path:
                            self.selected_folder_path.set(original_path)
                            self.save_config()
                        self.log_message(f"切り戻し: '{os.path.basename(new_path)}' -> '{os.path.basename(original_path)}'\n")
                    else:
                        self.log_message(f"警告: パス '{new_path}' が見つかりません。スキップします。\n")
                except Exception as e:
                    self.log_message(f"エラー: '{new_path}' の切り戻し中にエラー: {e}\n")
                finally:
                    self.progress_bar.step()
                    self.master.update_idletasks()

            self.log_message("--- 切り戻し処理が完了しました ---\n")
            messagebox.showinfo("完了", "リネームを元に戻しました。")

        except Exception as e:
            self.log_message(f"切り戻し中に予期せぬエラーが発生しました: {e}\n")
            messagebox.showerror("エラー", f"切り戻し中に予期せぬエラーが発生しました: {e}")
        finally:
            self.progress_bar.pack_forget()
            self.rename_history.clear() # 履歴をクリア
            self.clear_results() # 解析結果もクリア
            self.set_buttons_state(tk.NORMAL)
            self.revert_button.config(state=tk.DISABLED) # 切り戻しは一度きり

            # フォルダが選択されていなければ、解析・実行ボタンは無効のままにする
            if not os.path.isdir(self.selected_folder_path.get()):
                self.analyze_button.config(state=tk.DISABLED)
                self.execute_button.config(state=tk.DISABLED)

    def export_to_csv(self):
        """
        解析結果をCSVファイルに出力します。
        """
        if not self.analysis_results:
            messagebox.showwarning("警告", "エクスポートするデータがありません。")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV ファイル", "*.csv"), ("すべてのファイル", "*.*")],
            initialfile="rename_preview.csv",
            title="解析結果をCSVに保存"
        )

        if not filepath:
            self.log_message("CSVエクスポートがキャンセルされました。\n")
            return

        try:
            with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(['種類', '変更前の名前', '変更後の名前', '場所']) # ヘッダー
                for result in self.analysis_results:
                    writer.writerow([result['type'], result['original'], result['new'], result['path']])
            self.log_message(f"解析結果をCSVファイルにエクスポートしました: {filepath}\n")
            messagebox.showinfo("成功", f"CSVファイルとして保存しました:\n{filepath}")
        except Exception as e:
            self.log_message(f"CSVファイルのエクスポート中にエラーが発生しました: {e}\n")
            messagebox.showerror("エラー", f"CSVファイルの保存中にエラーが発生しました: {e}")

    def replace_invalid_chars(self, text, replacement):
        """
        文字列内の絵文字と、ユーザーが設定した追加の記号を指定された文字に置換します。
        """
        # 1. 絵文字のパターン文字列
        emoji_pattern_str = (
            "["
            "\U0001F600-\U0001F64F"  # Emoticons
            "\U0001F300-\U0001F5FF"  # Symbols & Pictographs
            "\U0001F680-\U0001F6FF"  # Transport & Map Symbols
            "\U0001F1E0-\U0001F1FF"  # Flags (iOS)
            "\U00002600-\U000026FF"  # Miscellaneous Symbols
            "\U00002702-\U000027B0"  # Dingbats
            "\U0001F900-\U0001F9FF"  # Supplemental Symbols and Pictographs
            "\U0001FA00-\U0001FA6F"  # Chess Symbols
            "\U00002B00-\U00002BFF"  # Miscellaneous Symbols and Arrows
            "\U000020D0-\U000020FF"  # Combining Diacritical Marks for Symbols
            "\U0000FE00-\U0000FE0F"  # Variation Selectors
            "\U0001F000-\U0001F02F"  # Mahjong Tiles, Domino Tiles
            "\U0001F0A0-\U0001F0FF"  # Playing Cards
            "\U0001F700-\U0001F77F"  # Alchemical Symbols
            "\U0001F780-\U0001F7FF"  # Geometric Shapes Extended
            "\U0001F800-\U0001F8FF"  # Supplemental Arrows-C
            "\U0001FB00-\U0001FBFF"  # Symbols for Legacy Computing
            "\u200d"                  # Zero Width Joiner (for compound emojis)
            "]+"
        )

        # 2. ユーザー定義の記号のパターン文字列
        custom_symbols = self.custom_symbols_to_replace.get()
        custom_pattern_str = ""
        if custom_symbols:
            # 正規表現で特別な意味を持つ文字をエスケープする
            escaped_symbols = re.escape(custom_symbols)
            custom_pattern_str = f"[{escaped_symbols}]+"

        # 3. パターンを結合
        if custom_pattern_str:
            combined_pattern_str = f"{emoji_pattern_str}|{custom_pattern_str}"
        else:
            combined_pattern_str = emoji_pattern_str

        combined_pattern = re.compile(combined_pattern_str, flags=re.UNICODE)
        return combined_pattern.sub(replacement, text)

    def save_config(self):
        """現在の設定（最後に選択したフォルダパス）をファイルに保存します。"""
        config_data = {
            "last_folder_path": self.selected_folder_path.get()
        }
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            self.log_message(f"設定の保存中にエラーが発生しました: {e}\n")

    def load_config(self):
        """設定ファイルから最後に選択したフォルダパスを読み込みます。"""
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)

                last_path = config_data.get("last_folder_path")
                if last_path and os.path.isdir(last_path):
                    self.selected_folder_path.set(last_path)
                    self.analyze_button.config(state=tk.NORMAL)
                    self.execute_button.config(state=tk.NORMAL)
                    self.log_message(f"前回のフォルダを読み込みました: {last_path}\n")
        except (json.JSONDecodeError, IOError) as e:
            self.log_message(f"設定の読み込み中にエラーが発生しました: {e}\n")

    def on_tree_double_click(self, event):
        """
        解析結果リストのアイテムがダブルクリックされたときに、
        その場所をエクスプローラーで開きます。
        """
        # 選択されているアイテムのIDを取得
        item_id = self.tree.focus()
        if not item_id:
            return

        # アイテムの情報を取得
        item_data = self.tree.item(item_id)
        values = item_data.get('values')
        if not values:
            return

        _, item_name, _, parent_path = values

        # アイテムのフルパスを構築
        full_path = os.path.join(parent_path, item_name)

        try:
            # パスが存在するかチェック
            if not os.path.exists(full_path):
                messagebox.showwarning("警告", f"パスが見つかりません: {full_path}")
                return

            # OSに応じてファイル/フォルダを開く
            if sys.platform == "win32":
                subprocess.run(['explorer', '/select,', full_path], check=True)
            elif sys.platform == "darwin": # macOS
                subprocess.run(['open', '-R', full_path], check=True)
            else: # Linuxなど
                path_to_open = parent_path if os.path.isfile(full_path) else full_path
                subprocess.run(['xdg-open', path_to_open], check=True)

            self.log_message(f"'{full_path}' をファイルマネージャーで開きました。\n")

        except Exception as e:
            error_message = f"場所を開けませんでした: {e}"
            self.log_message(f"エラー: {error_message}\n")
            messagebox.showerror("エラー", error_message)

    def sort_treeview_column(self, col):
        """
        Treeviewの列ヘッダーがクリックされたときに、その列を基準にソートします。
        """
        # ソート順を決定（同じ列なら逆順、違う列なら昇順）
        reverse = self.treeview_sort_reverse if col == self.treeview_sort_column else False
        reverse = not reverse

        # Treeviewからデータを取得
        l = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]

        # データをソート（文字列として、大文字小文字を区別せずにソート）
        l.sort(key=lambda t: t[0].lower(), reverse=reverse)

        # Treeview内のアイテムを並べ替え
        for index, (val, k) in enumerate(l):
            self.tree.move(k, '', index)

        # ヘッダーにソート方向の矢印を追加
        for c, text in self.treeview_headers.items():
            arrow = ' ▼' if reverse else ' ▲'
            self.tree.heading(c, text=text + (arrow if c == col else ''))

        self.treeview_sort_column = col
        self.treeview_sort_reverse = reverse

def main():
    """
    アプリケーションのエントリーポイント。
    Tkinterのルートウィンドウを作成し、RenamerAppインスタンスを起動します。
    """
    if ThemedTk:
        # ttkthemesが利用可能な場合、モダンなテーマを適用
        # 利用可能なテーマ例: "arc", "breeze", "plastik", "equilux", "itft1", "clearlooks"
        root = ThemedTk(theme="arc")
    else:
        # 利用できない場合は通常のTkinterウィンドウを使用
        root = tk.Tk()

    app = RenamerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
