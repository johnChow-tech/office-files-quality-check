# ==============================================================================
# gui.py - グラフィカルユーザーインターフェース (GUI) 実装 (レイアウト最適化: 左側 4 列 + コントロール列 1 列 + ログ 1 列)
# ==============================================================================
# サードパーティライブラリ
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
import threading

# ビジネスロジック
# extractor.py と url_opener.py が存在することを確認してください
from core.extractor import DocExtractor
try:
    from core.url_opener import UrlOpener
except ImportError:
    # インポートエラー
    messagebox.showerror(
        "インポートエラー", "url_opener.py をインポートできません。ファイルが存在するか確認してください。")
    sys.exit(1)


# ==============================================================================
# コアスレッドクラス：DocExtractor.run_extraction をバックグラウンドで実行するために使用
# ==============================================================================
class ExtractionWorker(threading.Thread):

    def __init__(self, extractor, source_path, output_path, file_map, callback):

        super().__init__()
        self.extractor = extractor
        self.source_path = source_path
        self.output_path = output_path
        self.file_map = file_map
        self.callback = callback
        self.result = False
        self.exception = None

    def run(self):
        try:
            # 進捗コールバック付きの抽出メソッドを呼び出し
            self.result = self.extractor.run_extraction(
                source_path=self.source_path,
                output_path=self.output_path,
                file_map=self.file_map,
                progress_callback=self.callback
            )
        except Exception as e:
            self.exception = e
        finally:
            # 成功か失敗かに関わらず、最終完了コールバックを一度呼び出し
            self.callback(len(self.file_map), len(self.file_map),
                          finished=True, exception=self.exception)


# ==============================================================================
# ログリダイレクトクラス：print() の出力を tk.Text コンポーネントに転送
# ==============================================================================
class TextRedirector:

    def __init__(self, widget, tag="stdout"):

        self.widget = widget
        self.tag = tag

    def write(self, string):
        # メインスレッドで挿入操作が実行されるように保証
        self.widget.after(0, self._insert_text, string)

    def _insert_text(self, string):
        self.widget.configure(state='normal')
        self.widget.insert(tk.END, string, self.tag)
        self.widget.see(tk.END)   # 自動的に最下部までスクロール
        self.widget.configure(state='disabled')   # 編集を無効化

    def flush(self):
        pass

# ==============================================================================
# MainGUI クラス
# ==============================================================================


class MainGUI:

    def __init__(self, master):

        self.master = master
        master.title("棚卸しツール")
        master.geometry('1440x860')

        # コアビジネスオブジェクト
        try:
            self.extractor = DocExtractor()
        except RuntimeError as e:
            messagebox.showerror("初期化エラー", f"アプリケーションの初期化に失敗しました: {e}\n\n必要なライブラリがインストールされているか確認してください。")
            sys.exit(1)  # Exit application if DocExtractor cannot be initialized
        self.opener = UrlOpener()

        # 環境変数
        current_dir = os.getcwd()
        self.sourcePath = tk.StringVar(value=current_dir)
        self.outputPath = tk.StringVar(value=current_dir)
        self.outputTextPath = tk.StringVar()
        self.outputUrlsPath = tk.StringVar()

        # プログレスバー関連変数
        self.progress_value = tk.DoubleVar()

        # 独立した Listbox のパス対応付け辞書
        self.source_path_map = {}
        self.text_output_path_map = {}
        self.urls_output_path_map = {}

        # ログの表示/非表示ステータス変数
        self.log_visible = tk.BooleanVar(value=True)

        # toggleLogBtn のテキストを保存するための変数、デフォルト値は "<<"
        self.toggle_log_btn_text = tk.StringVar(value="<<")

        # 初期化
        self._configure_grid()
        self._create_widgets()
        self._layout_widgets()
        self._bind_events()

        # sys.stdout および sys.stderr のリダイレクト
        self._redirect_output()

        # 起動時にソースファイルリストを即座にロード
        if os.path.isdir(current_dir):
            self._load_files_to_listbox(
                folder_path=current_dir,
                target_listbox=self.sourceListbox,
                path_map=self.source_path_map,
                extension_filter=self.extractor.TARGET_EXTENSIONS
            )

    # Grid configuration

    def _configure_grid(self):
        """フレームのグリッド設定 (左側 4 列 + コントロール列 1 列 + 右側ログ 1 列)"""
        # C0, C1, C2, C3: 入力/出力エリア
        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_columnconfigure(1, weight=1)
        self.master.grid_columnconfigure(2, weight=1)
        self.master.grid_columnconfigure(3, weight=1)

        # C4: コントロール列
        self.master.grid_columnconfigure(4, weight=0)

        # C5: ログエリア
        self.master.grid_columnconfigure(5, weight=2)

        # R1-R4, R7-R8 はウェイトを維持
        self.master.grid_rowconfigure(1, weight=1)
        self.master.grid_rowconfigure(2, weight=1)
        self.master.grid_rowconfigure(3, weight=1)
        self.master.grid_rowconfigure(4, weight=1)
        self.master.grid_rowconfigure(7, weight=1)
        self.master.grid_rowconfigure(8, weight=1)

    # Widgets
    def _create_widgets(self):
        """コンポーネントを作成"""
        # fmt:off
        # Buttons
        self.sourceFolderSelectBtn = tk.Button(self.master, text="作業パスを選択")
        self.outputFolderSelectBtn = tk.Button(self.master, text="出力パスを選択")
        self.harvestBtn = tk.Button(self.master, text="抽出する")
        self.textFolderOpenBtn = tk.Button(self.master, text="フォルダを開く")
        self.urlsFolderOpenBtn = tk.Button(self.master, text="フォルダを開く")
        self.checkAllUrlsBtn = tk.Button(self.master, text="一括チェック") 

        # ログ切り替えボタン
        self.toggleLogBtn = tk.Button(self.master, 
                      textvariable=self.toggle_log_btn_text, 
                      command=self._toggle_log_visibility) 

        # Entries
        self.sourceFolder = tk.Entry(self.master,textvariable=self.sourcePath,state='readonly')
        self.outputFolder = tk.Entry(self.master,textvariable=self.outputPath,state='readonly')

        # プログレスバーコンポーネント
        self.progressBar = ttk.Progressbar(self.master, orient=tk.HORIZONTAL, mode='determinate', variable=self.progress_value)

        # ログウィンドウとそのスクロールバー (C5)
        self.logFrame = tk.LabelFrame(self.master, text="[即時ログ / Console Output]", foreground="gray33")
        self.logScrollbar = tk.Scrollbar(self.logFrame)
        self.logText = tk.Text(self.logFrame, wrap='word', height=40, width=40, state='disabled', 
                    yscrollcommand=self.logScrollbar.set,
                    font=("Consolas", 9)) 

        self.logScrollbar.config(command=self.logText.yview)
        # fmt:on

        # Frames and Listboxes
        # 1. ソースファイル Listbox
        self.sourceListboxFrame = tk.LabelFrame(
            self.master, text="[作業パス]/*.docx,*.xls*,*.pptx", foreground="gray33")
        self.sourceFrame, self.sourceListbox = self._create_listbox_with__scrollbar(
            self.sourceListboxFrame)

        # 2. PlainText 出力 Listbox
        self.outputTextListboxFrame = tk.LabelFrame(
            self.master, text="[出力パス]/PlainText/*", foreground="gray33")
        self.outputTextFrame, self.ouputTextListbox = self._create_listbox_with__scrollbar(
            self.outputTextListboxFrame)

        # 3. HyperLinks 出力 Listbox
        self.outputUrlsListboxFrame = tk.LabelFrame(
            self.master, text="[出力パス]/HyperLinks/*", foreground="gray33")
        self.outputUrlsFrame, self.outputUrlsListbox = self._create_listbox_with__scrollbar(
            self.outputUrlsListboxFrame)

    def _create_listbox_with__scrollbar(self, parent):
        """スクロールバー付きのListboxを作成。親コンテナは渡された parent"""
        frame = tk.Frame(parent)
        sb = tk.Scrollbar(frame, orient=tk.VERTICAL)
        listbox = tk.Listbox(frame, yscrollcommand=sb.set,
                             selectmode=tk.EXTENDED)
        sb.config(command=listbox.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        return frame, listbox

    def _redirect_output(self):
        """sys.stdout と sys.stderr をログウィンドウにリダイレクト"""
        sys.stdout = TextRedirector(self.logText, "stdout")
        sys.stderr = TextRedirector(self.logText, "stderr")
        self.logText.tag_config("stderr", foreground="red")

    def _bind_events(self):
        """イベントハンドラを対応するボタンにバインドする。"""
        self.sourceFolderSelectBtn.config(
            command=self._select_source_path_and_load_to_listbox)
        self.outputFolderSelectBtn.config(
            command=self._select_output_path)
        self.harvestBtn.config(
            command=self._harvest_btn_on_click)
        self.textFolderOpenBtn.config(
            command=lambda: self._open_folder_windows(self.outputTextPath.get()))
        self.urlsFolderOpenBtn.config(
            command=lambda: self._open_folder_windows(self.outputUrlsPath.get()))
        self.checkAllUrlsBtn.config(
            command=self._check_all_urls_btn_on_click)
        self.toggleLogBtn.config(command=self._toggle_log_visibility)

    def _toggle_log_visibility(self):
        """ログウィンドウの表示/非表示状態を切り替える"""
        is_visible = self.log_visible.get()
        C5, R0, PADY = 5, 0, 8

        # 現在表示されている場合、非表示にする
        if is_visible:
            self.logFrame.grid_forget()
            # ウェイトをゼロにし、スペースを解放
            self.master.grid_columnconfigure(C5, weight=0)
            self.toggle_log_btn_text.set(">>")
        else:
            # 現在非表示の場合、表示にする
            self.logFrame.grid(column=C5, row=R0, rowspan=10,
                               sticky='eswn', padx=8, pady=PADY)
            self.master.grid_columnconfigure(C5, weight=2)   # C5 のウェイトを復元
            self.toggle_log_btn_text.set("<<")

        # 状態変数を切り替え
        self.log_visible.set(not is_visible)

    def _open_folder_windows(self, path):
        """対象フォルダを開く (Windows のみ互換)"""
        if not path:
            messagebox.showwarning("警告", "パスが空のため、フォルダを開けません。")
            return

        if sys.platform == "win32":
            try:
                os.startfile(path)
            except FileNotFoundError:
                messagebox.showerror("エラー", f"パスが見つかりません: {path}")

    def _load_files_to_listbox(self, folder_path, target_listbox, path_map, extension_filter=None):
        """
        指定されたフォルダ内のファイルをスキャンし、そのファイル名を Listbox に設定し、同時にパス対応付けを作成します。
        """
        target_listbox.delete(0, tk.END)
        path_map.clear()

        if not os.path.isdir(folder_path):
            return

        try:
            items = os.listdir(folder_path)

            for item_name in sorted(items):
                absolute_path = os.path.join(folder_path, item_name)

                is_file = os.path.isfile(absolute_path)

                if extension_filter:
                    is_match = item_name.lower().endswith(extension_filter)
                else:
                    is_match = True

                if is_file and is_match:
                    listbox_index = target_listbox.size()
                    target_listbox.insert(tk.END, item_name)
                    path_map[listbox_index] = absolute_path

        except Exception as e:
            print(
                f"読み込みエラー: フォルダ内容 {folder_path} を読み込めません:\n{e}", file=sys.stderr)

    # --- A 部 ビジネスロジック ---

    def _select_source_path_and_load_to_listbox(self):
        """ダイアログを開いてソースフォルダを選択し、ソース Listbox に設定"""
        path = filedialog.askdirectory(title="ファイルを含むフォルダを選択してください")

        if path:
            self.sourcePath.set(path)
            self._load_files_to_listbox(
                folder_path=path,
                target_listbox=self.sourceListbox,
                path_map=self.source_path_map,
                extension_filter=self.extractor.TARGET_EXTENSIONS
            )
            print(f"作業パスが選択されました: {path}")

    def _select_output_path(self):
        """出力フォルダパスを選択"""
        path = filedialog.askdirectory(title="出力フォルダを選択してください")

        if path:
            self.outputPath.set(path)
            self.text_output_path_map.clear()
            self.urls_output_path_map.clear()
            self.ouputTextListbox.delete(0, tk.END)
            self.outputUrlsListbox.delete(0, tk.END)
            self._reset_progress()
            print(f"出力パスが選択されました: {path}")
            print()
            print("【抽出する】ボタンをクリックして抽出を開始してください。")
            print()

    def _reset_progress(self):
        """プログレスバーの状態をリセット"""
        self.progress_value.set(0)
        self.harvestBtn.config(state=tk.NORMAL)
        self.sourceFolderSelectBtn.config(state=tk.NORMAL)
        self.outputFolderSelectBtn.config(state=tk.NORMAL)

    def _update_progress(self, current, total, finished=False, exception=None):
        """
        進捗コールバック関数。メインスレッドで GUI の状態を更新するために使用されます。
        """
        self.master.after(0, self.__safe_update_progress,
                          current, total, finished, exception)

    def __safe_update_progress(self, current, total, finished, exception):
        """メインスレッドで実際に GUI の更新を実行"""
        if exception:
            self._reset_progress()
            messagebox.showerror("操作失敗", f"ファイル抽出中にエラーが発生しました: {exception}")
            return

        if finished:
            self.progress_value.set(100)
            self._refresh_output_listboxes()
            messagebox.showinfo("操作成功", "ファイル内容の抽出と保存が完了しました！")
            self._reset_progress()
            return

        if total > 0:
            percentage = (current / total) * 100
            self.progress_value.set(percentage)
            if current % 1 == 0 or current == total:
                # シンプルなログ進捗を追加
                print(f"進捗: {current}/{total} 番目のファイルを抽出中です...")
        else:
            self.progress_value.set(0)

    def _refresh_output_listboxes(self):
        """抽出完了後に出力ファイルリストを更新"""
        outputPath = self.outputPath.get()
        TEXT_FOLDER = self.extractor.TEXT_FOLDER
        URLS_FOLDER = self.extractor.URLS_FOLDER

        outputTextPath = os.path.join(outputPath, TEXT_FOLDER)
        outputUrlsPath = os.path.join(outputPath, URLS_FOLDER)

        self.outputTextPath.set(outputTextPath)
        self.outputUrlsPath.set(outputUrlsPath)

        self._load_files_to_listbox(
            folder_path=outputTextPath,
            target_listbox=self.ouputTextListbox,
            path_map=self.text_output_path_map,
            extension_filter=('.md',)
        )

        self._load_files_to_listbox(
            folder_path=outputUrlsPath,
            target_listbox=self.outputUrlsListbox,
            path_map=self.urls_output_path_map,
            extension_filter=('.csv',)
        )

    def _harvest_btn_on_click(self):
        """抽出ボタンのクリックイベントを処理し、バックグラウンドスレッドでコア抽出ロジックを呼び出す"""
        sourcePath = self.sourcePath.get()
        outputPath = self.outputPath.get()

        if not sourcePath or not outputPath:
            messagebox.showwarning("入力警告", "まず【作業パス】と【出力パス】を選択してください！")
            return

        if not self.source_path_map:
            messagebox.showwarning(
                "ファイル警告", "作業パス内に処理可能な Office ファイルが見つかりませんでした。")
            return

        # ボタンを無効化し、重複クリックを防止
        self.harvestBtn.config(state=tk.DISABLED)
        self.sourceFolderSelectBtn.config(state=tk.DISABLED)
        self.outputFolderSelectBtn.config(state=tk.DISABLED)
        self.progress_value.set(0)

        # ログウィンドウをクリア
        self.logText.configure(state='normal')
        self.logText.delete('1.0', tk.END)
        self.logText.configure(state='disabled')
        print("--- 抽出タスク開始 ---")

        # バックグラウンドスレッドを起動
        worker = ExtractionWorker(
            extractor=self.extractor,
            source_path=sourcePath,
            output_path=outputPath,
            file_map=self.source_path_map,
            callback=self._update_progress
        )
        worker.start()

        # --- B 部 ビジネスロジック ---

    def _check_all_urls_btn_on_click(self):
        """「一括チェック」ボタンのクリックイベントを処理し、リンクの一括オープンロジックを呼び出す"""
        urlsPath = self.outputUrlsPath.get()
        base_output_path = self.outputPath.get()

        if not urlsPath or not os.path.isdir(urlsPath):
            messagebox.showwarning(
                "警告", "まず【抽出する】をクリックしてハイパーリンクファイル (.csv) を生成してください！")
            return

        # リストを更新し、パス対応付けが最新であることを確認
        self._load_files_to_listbox(
            folder_path=urlsPath,
            target_listbox=self.outputUrlsListbox,
            path_map=self.urls_output_path_map,
            extension_filter=('.csv',)
        )

        if not self.urls_output_path_map:
            messagebox.showinfo(
                "ファイル警告", "【HyperLinks】ディレクトリに Urls_*.csv ファイルが見つかりませんでした。")
            return

        selected_indices = self.outputUrlsListbox.curselection()
        if selected_indices:
            files_to_open = {
                i: self.urls_output_path_map[i] for i in selected_indices if i in self.urls_output_path_map}
        else:
            files_to_open = self.urls_output_path_map

        if not files_to_open:
            messagebox.showwarning("警告", "開くためにハイパーリンクファイルを少なくとも一つ選択してください。")
            return

        confirm = messagebox.askyesno(
            "操作の確認",
            f"{len(files_to_open)} 個のハイパーリンクファイル内の全てのリンクを開こうとしています。\n\n続行しますか？"
        )
        if not confirm:
            return

        print("\n--- リンクチェックタスク開始 ---")
        success_count = self.opener.open_links_for_qc(
            urls_folder_path=urlsPath,
            file_map=files_to_open,
            base_output_path=base_output_path
        )
        print("--- リンクチェックタスク完了 ---")

        if success_count > 0:
            messagebox.showinfo(
                "操作完了", f"{success_count} 個のソースファイルのハイパーリンクが正常に開かれました！ブラウザに切り替えて確認してください。")
        else:
            messagebox.showwarning("操作完了", "開く必要のある有効なハイパーリンクは見つかりませんでした。")

        # Layout

    def _layout_widgets(self):
        """コンポーネントの位置決めと外観設定"""
        # Constants
        PADX = 8
        PADY = 8
        R0, R1, R2, R3, R4, R5, R6, R7, R8, R9 = 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
        # C0-C3 は左側コンテンツ (4列), C4 はコントロール列, C5 は右側ログ
        C0, C1, C2, C3, C4, C5 = 0, 1, 2, 3, 4, 5
        # fmt: off

        # ----------------------------------- 左側コンテンツ (C0-C3) -----------------------------------

        # R0: フォルダ選択ボタン & 入力ボックス
        self.sourceFolderSelectBtn.grid(column=C0,row=R0,sticky='ew',padx=PADX, pady=PADY)
        self.sourceFolder.grid(column=C1, columnspan=3,row=R0,sticky='ew',padx=PADX, pady=PADY) 

        # R1-R4: ソース Listbox フレーム (C0-C3 を占有)
        self.sourceListboxFrame.grid(column=C0, columnspan=4, row=R1, rowspan=4, sticky='eswn', padx=PADX, pady=PADY)
        self.sourceFrame.pack(fill=tk.BOTH, expand=True, padx=PADX, pady=(0, PADY)) 

        # R5: 出力パス選択ボタン & 入力ボックス
        self.outputFolderSelectBtn.grid(column=C0, row=R5,sticky='ew',padx=PADX, pady=PADY) 
        self.outputFolder.grid(column=C1, columnspan=3,row=R5,sticky='eswn',padx=PADX, pady=PADY) 

        # R6: 抽出ボタン / プログレスバー
        self.harvestBtn.grid(column=C0, row=R6, sticky='ew', padx=PADX, pady=PADY)
        # self.progressLabel.grid(column=C1, row=R6, sticky='w', padx=PADX, pady=PADY) 
        # プログレスバーは C1-C3 を占有 (3列)
        self.progressBar.grid(column=C1, columnspan=3, row=R6, sticky='ew', padx=PADX, pady=PADY) 

        # R7-R8: 出力 Listbox フレーム
        self.outputTextListboxFrame.grid(column=C0, columnspan=2, row=R7, rowspan=2, sticky='eswn', padx=PADX, pady=PADY) 
        self.outputTextFrame.pack(fill=tk.BOTH, expand=True, padx=PADX, pady=(0, PADY)) 
        self.outputUrlsListboxFrame.grid(column=C2, columnspan=2, row=R7, rowspan=2, sticky='eswn', padx=PADX, pady=PADY) 
        self.outputUrlsFrame.pack(fill=tk.BOTH, expand=True, padx=PADX, pady=(0, PADY)) 

        # R9: フォルダオープンボタンとリンクチェックボタン
        self.textFolderOpenBtn.grid(column=C0, row=R9, sticky='ew', padx=PADX, pady=PADY)
        self.urlsFolderOpenBtn.grid(column=C2, row=R9, sticky='ew', padx=PADX, pady=PADY)
        self.checkAllUrlsBtn.grid(column=C3, row=R9,sticky='ew',padx=PADX, pady=PADY)

        # ----------------------------------- 中央コントロール列 (C4) -----------------------------------
        self.toggleLogBtn.grid(column=C4, row=R0, rowspan=10, sticky='ns', padx=(1, 1), pady=PADY)

        # ----------------------------------- 右側ログ (C5) -----------------------------------
        self.logFrame.grid(column=C5, row=R0, rowspan=10, sticky='eswn', padx=PADX, pady=PADY)
        self.logScrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.logText.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=PADX, pady=(0, PADY))
        # fmt: on


if __name__ == '__main__':
    root = tk.Tk()
    # 一時的に元の stdout/stderr を保存
    original_stdout = sys.stdout
    original_stderr = sys.stderr

    app = MainGUI(root)
    root.mainloop()

    # 終了後に元の stdout/stderr を復元
    sys.stdout = original_stdout
    sys.stderr = original_stderr
