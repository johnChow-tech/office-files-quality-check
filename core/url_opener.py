import os
import csv
from datetime import datetime
import webbrowser
from tkinter import messagebox
import time

class UrlOpener:
    """
    Urls_*.csv ファイル（CSV形式）を読み込み、含まれるハイパーリンクを一括で開く役割を担います。
    グローバル重複排除と、プロセスごとの単一一時フォルダ管理を実装済み。
    """

    def __init__(self):
        print("UrlOpener の初期化が完了しました。")
        self.temp_html_path = None  # クリーンアップ用に唯一の一時ディレクトリを保持

    def _get_urls_from_dat(self, file_path):
        """
        単一の .csv ファイルを読み込み、URLを抽出・補完します。
        """
        urls_tuples = []
        source_file_name = os.path.basename(file_path)

        try:
            # utf-8-sig を使用して BOM 混入を防止
            with open(file_path, 'r', encoding='utf-8-sig', newline='') as datafile:
                reader = csv.DictReader(datafile)
                
                # ヘッダーの正規化（大文字小文字/空白を無視）
                headers = [h.strip() if h else "" for h in reader.fieldnames] if reader.fieldnames else []
                url_col = next((h for h in headers if h.upper() == "URL"), None)
                src_col = next((h for h in headers if h.upper() == "SOURCE FILE"), None)

                if not url_col:
                    print(f"警告: {source_file_name} に 'URL' 列が見つかりません。")
                    return source_file_name, []

                for row in reader:
                    raw_url = row.get(url_col, "").strip() if url_col in row else ""
                    if not raw_url:
                        continue

                    # プロトコルの自動補完 (http/https/mailto/file がない場合は https を付与)
                    if not raw_url.lower().startswith(('http://', 'https://', 'mailto:', 'file:')):
                        if '.' in raw_url:
                            full_url = "https://" + raw_url
                        else:
                            continue  # 無効な文字列はスキップ
                    else:
                        full_url = raw_url
                    
                    urls_tuples.append(full_url)

                    # ソースファイル名の更新（最初の有効な値を使用）
                    if src_col and row.get(src_col):
                        source_file_name = row.get(src_col)

        except Exception as e:
            print(f"エラー: {file_path} の読み込みに失敗しました: {e}")
            return os.path.basename(file_path), []

        return source_file_name, sorted(list(set(urls_tuples)))

    def _create_and_open_qc_prompt(self, target_dir, source_file_name, url_count):
        """
        指定されたフォルダ内に QC ページを生成し、ブラウザで開きます。
        """
        ILLEGAL_CHARS = r'/\?%*:|"<>' + '【】'
        cleaned_name = source_file_name
        for char in ILLEGAL_CHARS:
            cleaned_name = cleaned_name.replace(char, '_')

        html_content = f"""
        <!DOCTYPE html>
        <html lang="ja">
        <head>
            <meta charset="UTF-8">
            <title>QC - {source_file_name}</title>
            <style>
                body {{ font-family: sans-serif; margin: 40px; background-color: #f4f4f9; }}
                .container {{ background-color: #ffffff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
                h1 {{ color: #2c3e50; border-bottom: 3px solid #3498db; }}
                .highlight {{ color: #e74c3c; font-weight: bold; }}
            </style>
        </head>
        <body>
            <div class="container">
                <h1>品質チェック (QC) 開始</h1>
                <p>対象ファイル：<span class="highlight">{source_file_name}</span></p>
                <p>新規リンク数：<strong>{url_count}</strong> 個</p>
                <p>このタブの後に開かれるページを確認してください。</p>
            </div>
        </body>
        </html>
        """
        
        file_name = f"QC_Prompt_{cleaned_name.replace('.', '_')}.html"
        file_path = os.path.join(target_dir, file_name)

        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            webbrowser.open_new_tab('file://' + os.path.abspath(file_path))
            time.sleep(0.8)
        except Exception as e:
            print(f"QC ページの生成失敗: {e}")

    def open_links_for_qc(self, urls_folder_path: str, file_map: dict, base_output_path: str):
        """
        メイン制御：一貫した一時ディレクトリ内で全プロセスの処理を行います。
        """
        if not file_map:
            return 0

        # --- プロセスごとに唯一の一時ディレクトリを生成 ---
        formatted_time = datetime.now().strftime("%Y%m%d%H%M%S")
        unique_html_dir = os.path.join(base_output_path, f"temp_html_{formatted_time}")
        os.makedirs(unique_html_dir, exist_ok=True)
        self.temp_html_path = unique_html_dir

        global_seen_urls = set()
        urls_to_open_by_file = []

        print("\n--- フェーズ 1: 重複排除実行 ---")
        for dat_abs_path in file_map.values():
            source_file, unique_urls = self._get_urls_from_dat(dat_abs_path)
            
            new_urls = [url for url in unique_urls if url not in global_seen_urls]
            for url in new_urls:
                global_seen_urls.add(url)

            if new_urls:
                urls_to_open_by_file.append({"source_file": source_file, "urls": new_urls})
            
            print(f"  {os.path.basename(dat_abs_path)}: 新規 {len(new_urls)} 個")

        print("\n--- フェーズ 2: リンク展開 ---")
        if not global_seen_urls:
            messagebox.showinfo("結果", "開くべき新しいリンクは見つかりませんでした。")
            return 0

        for item in urls_to_open_by_file:
            # 同一フォルダ内に HTML を生成
            self._create_and_open_qc_prompt(unique_html_dir, item["source_file"], len(item["urls"]))
            
            for url in item["urls"]:
                webbrowser.open_new_tab(url)
                time.sleep(0.4)

        return len(urls_to_open_by_file)