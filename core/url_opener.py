# ==============================================================================
# url_opener.py - DAT/CSV を読み込み、ハイパーリンクを一括で開くロジック (ファイル間でのグローバル重複排除を実現)
# ==============================================================================
import os
import csv
from datetime import datetime
import webbrowser
from tkinter import messagebox
import time


class UrlOpener:
    """
    Urls_*.csv ファイル（CSV形式）を読み込み、含まれるハイパーリンクを一括で開く役割を担います。
    """

    def __init__(self):
        print("UrlOpener の初期化が完了しました。")
        self.temp_html_path = None  # 一時HTMLディレクトリを保存し、外部からのクリーンアップに利用

    def _get_urls_from_dat(self, file_path):
            urls_tuples = []
            source_file_name = os.path.basename(file_path)
    
            try:
                # 使用 utf-8-sig 兼容带 BOM 的文件
                with open(file_path, 'r', encoding='utf-8-sig', newline='') as datafile:
                    reader = csv.DictReader(datafile)
                    
                    # 兼容性处理：防止标题行存在空格或大小写不一致
                    headers = [h.strip() if h else "" for h in reader.fieldnames] if reader.fieldnames else []
                    url_col = next((h for h in headers if h.upper() == "URL"), None)
                    src_col = next((h for h in headers if h.upper() == "SOURCE FILE"), None)
    
                    if not url_col:
                        print(f"警告: 文件 {source_file_name} 缺失 'URL' 列。")
                        return source_file_name, []
    
                    for row in reader:
                        # 获取 URL 并去除空白字符
                        raw_url = row.get(url_col, "").strip()
                        
                        # 修正过滤逻辑：只要包含 '.' 且不是纯数字，或者以协议开头，即视为有效 URL
                        if raw_url:
                            # 自动补全缺失协议的链接
                            if not raw_url.lower().startswith(('http://', 'https://', 'mailto:', 'file:')):
                                if '.' in raw_url:
                                    full_url = "https://" + raw_url
                                else:
                                    continue # 可能是内部锚点，跳过
                            else:
                                full_url = raw_url
                            
                            urls_tuples.append(full_url)
    
                        # 更新源文件名（仅取第一行）
                        if src_col and row.get(src_col):
                            source_file_name = row.get(src_col)
    
            except Exception as e:
                print(f"读取错误: {file_path} -> {e}")
                return os.path.basename(file_path), []
    
            unique_urls_in_file = sorted(list(set(urls_tuples)))
            return source_file_name, unique_urls_in_file

    def _create_and_open_qc_prompt(self, base_output_path, source_file_name, url_count):
        """
        QC (品質チェック) 案内ページを作成し、ブラウザで開きます。
        重要な修正: ファイル名に対して不正な文字をフィルタリングします。
        """
        # フィルタリングが必要な不正なファイル名文字を定義
        # 一般的な Windows/Unix の不正文字と、言及された中国語文字を含む
        ILLEGAL_CHARS = r'/\?%*:|"<>' + '【】'

        # 1. 元ファイル名のクリーンアップ
        cleaned_source_file_name = source_file_name
        for char in ILLEGAL_CHARS:
            # 不正な文字を空文字列に置換
            cleaned_source_file_name = cleaned_source_file_name.replace(
                char, '_')  # アンダーバーに置換する方が安全

        # 2. ターゲットパスの構築
        formatted_time = datetime.now().strftime("%Y%m%d%H%M%S")
        html_output_dir = os.path.join(
            base_output_path, f"temp_html_{formatted_time}")
        self.temp_html_path = html_output_dir  # パスを記録し、外部 (gui.py) のクリーンアップに利用

        if not os.path.exists(html_output_dir):
            os.makedirs(html_output_dir)

        # HTML コンテンツは変更なし (コードの簡潔さのため、内部の HTML_CONTENT 定義は省略)
        html_content = f"""
        <!DOCTYPE html>
        <html lang="ja">
        <head>
            <meta charset="UTF-8">
            <title>{source_file_name}-QC</title>
            <style>
                body {{ font-family: sans-serif; margin: 40px; background-color: #f4f4f9; }}
                .container {{ background-color: #ffffff; padding: 30px; border-radius: 8px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); }}
                h1 {{ color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 10px; }}
                p {{ font-size: 1.1em; color: #34495e; }}
                .highlight {{ color: #e74c3c; font-weight: bold; font-size: 1.2em; }}
                .count {{ color: #27ae60; font-weight: bold; }}
            </style>
        </head>
        <body>
            <div class="container">
                <h1>ハイパーリンク品質チェック (QC) 開始</h1>
                <p>現在処理中の元ファイルは：<span class="highlight">{source_file_name}</span></p>
                <p>これから <span class="count">{url_count}</span> 個の**重複排除後**のハイパーリンクを開きます。</p>
                <p>続けて開かれる各タブのリンクの有効性と内容を注意深く確認してください。</p>
            </div>
        </body>
        </html>
        """

        # 3. クリーンアップされたファイル名を使用して一時ファイルパスを構築
        # 注意: ここではクリーンアップされたファイル名を使用し、パスの問題を防ぐためにドットを削除しています
        temp_file_name = f"QC_Prompt_{cleaned_source_file_name.replace('.', '_')}.html"
        temp_file_path = os.path.join(html_output_dir, temp_file_name)

        try:
            with open(temp_file_path, 'w', encoding='utf-8') as f:
                f.write(html_content)

            print(f"    -> 案内ページが生成されました: {temp_file_path}")
            webbrowser.open_new_tab('file://' + temp_file_path)
            time.sleep(1)

        except Exception as e:
            print(f"    [エラー] HTML 案内ページを作成または開けませんでした: {e}")

    def open_links_for_qc(self, urls_folder_path: str, file_map: dict, base_output_path: str):
        """
        メイン制御メソッド: 全てのハイパーリンク DAT ファイルを走査し、まずグローバル重複排除を実行してからリンクを開きます。
        """
        if not urls_folder_path or not file_map or not base_output_path:
            print("エラー：パスまたはファイルリストが空です。")
            return 0

        # ------------------------------------------------------------------
        # 重要な修正: グローバル重複排除メカニズムの導入
        # ------------------------------------------------------------------
        global_seen_urls = set()

        # 開く URL のリストを、処理するファイル順にグループ化して格納
        urls_to_open_by_file = []

        # フェーズ 1: 全てのファイルリンクを収集し、グローバル重複排除を実行
        print("\n--- フェーズ 1: 全てのファイルリンクを収集し、グローバル重複排除を実行 ---")

        for index, dat_abs_path in file_map.items():
            source_file, unique_urls_in_file = self._get_urls_from_dat(
                dat_abs_path)

            # 現在のファイル内で、まだグローバルセットに出現していない URL をフィルタリング
            new_urls = []
            for url in unique_urls_in_file:
                if url not in global_seen_urls:
                    new_urls.append(url)
                    global_seen_urls.add(url)
                # else:
                #    # オプションとしてログを出力し、ファイル間での重複であることを説明
                #    # print(f"       [グローバル重複排除スキップ] リンクは既に他のファイルに出現しています: {url[:50]}...")

            if new_urls:
                urls_to_open_by_file.append({
                    "source_file": source_file,
                    "urls": new_urls
                })

            print(
                f"   ファイル {os.path.basename(dat_abs_path)}: {len(unique_urls_in_file)} 個のリンクが見つかりました (新規 {len(new_urls)} 個)")

        # ------------------------------------------------------------------
        # フェーズ 2: グローバル重複排除後のリストに基づき、ファイルごとにリンクを開く
        # ------------------------------------------------------------------
        print("\n--- フェーズ 2: グローバル重複排除後のリンクを開く ---")

        opened_file_count = 0
        total_unique_links = len(global_seen_urls)

        if total_unique_links == 0:
            print("グローバル重複排除後、開く必要のある固有のリンクは見つかりませんでした。")
            return 0

        for item in urls_to_open_by_file:
            source_file = item["source_file"]
            urls_to_open = item["urls"]

            opened_file_count += 1
            print(
                f"\n--- 元ファイルを処理中 ({opened_file_count}/{len(urls_to_open_by_file)}): {source_file}、{len(urls_to_open)} 個の**新規**リンクを開きます ---")

            # base_output_path を渡す
            self._create_and_open_qc_prompt(
                base_output_path, source_file, len(urls_to_open))

            # リンクを開く
            for i, url in enumerate(urls_to_open):
                print(f"    -> リンク {i+1}/{len(urls_to_open)} を開いています: {url}")
                webbrowser.open_new_tab(url)
                time.sleep(0.5)

        messagebox.showinfo(
            "重複排除結果",
            f"全てのファイル内のリンクの重複排除が完了しました。合計 {total_unique_links} 個の固有リンクが見つかり、全てブラウザで開かれました。"
        )

        return opened_file_count
