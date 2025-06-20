# -*- coding: utf-8 -*-
# 終極版本 V5.1：Calibre解包 + 自動偵測可攜式KCC重新打包

import sys
import os
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

# --- 使用者設定區 ---

# 1. KCC 裝置設定檔 (KPW5 = Kindle Paperwhite 5 / 11代，是很好的通用高解析度選項)
#    其他選項如：KOA3 (Kindle Oasis 3), KV (Kindle Voyage)
KCC_DEVICE_PROFILE = "KPW5"

# 2. 是否在解包後自動執行重新打包 (True = 是, False = 否)
REPACK_AFTER_UNPACK = True

# 3. KCC命令列工具的檔名
KCC_EXE_FILENAME = "kcc-c2e.exe"

# --- 設定區結束 ---


def repack_with_kcc(image_folder_path: Path):
    """
    Uses a portable Kindle Comic Converter (KCC) to repack a folder of images
    into an optimized .mobi file.
    """
    print("\n--- KCC 重新打包階段 ---")
    
    # 自動偵測 KCC 路徑：在腳本所在的同一個資料夾尋找 KCC 執行檔
    try:
        script_dir = Path(__file__).parent
        kcc_path = script_dir / KCC_EXE_FILENAME
    except NameError:
        # 如果在非標準 Python 環境中 __file__ 未定義，則使用當前工作目錄
        script_dir = Path.cwd()
        kcc_path = script_dir / KCC_EXE_FILENAME

    if not kcc_path.is_file():
        print(f"❌ 致命錯誤: 在腳本所在的目錄找不到 KCC 執行檔 '{KCC_EXE_FILENAME}'")
        print("   請確保 kcc-c2e.exe 和您的 Python 腳本放在同一個資料夾中。")
        return

    # 輸出檔案的路徑和名稱
    output_mobi_path = image_folder_path.parent
    book_title = image_folder_path.name

    print(f"🔧 自動偵測到 KCC: '{kcc_path}'")
    print(f"   準備處理資料夾: '{image_folder_path.name}'")
    print(f"   裝置設定檔: {KCC_DEVICE_PROFILE}")

    try:
        command = [
            str(kcc_path),
            "-p", KCC_DEVICE_PROFILE,
            "-m",
            "--manga-style",
            "-t", book_title,
            "-o", str(output_mobi_path),
            str(image_folder_path)
        ]
        
        result = subprocess.run(
            command,
            check=True,
            capture_output=True,
            text=True,
            encoding='utf-8',
            shell=True
        )

        final_file_name = f"{book_title}.mobi"
        final_file_path = output_mobi_path / final_file_name
        
        if final_file_path.exists():
             print(f"✅ KCC 打包成功！優化後的檔案已儲存為: '{final_file_path.name}'")
        else:
             print(f"⚠️ KCC 執行完畢，但未找到預期的輸出檔案 '{final_file_name}'。請檢查輸出目錄。")

    except subprocess.CalledProcessError as e:
        print("❌ 錯誤: KCC 執行失敗。")
        print(f"   返回碼: {e.returncode}")
        print(f"   KCC 輸出訊息: {e.stdout}{e.stderr}")
    except Exception as e:
        print(f"❌ 執行 KCC 時發生未知錯誤: {e}")


def unpack_comic_with_calibre(mobi_path: Path) -> Path | None:
    # ... (此函式與 V5 版本完全相同) ...
    if not shutil.which("ebook-convert"):
        print("❌ 致命錯誤: 找不到 'ebook-convert' 指令。")
        print("   請先安裝 Calibre，並確保其命令列工具已加入系統 PATH。")
        return None

    output_dir = mobi_path.parent / mobi_path.stem
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        print(f"📂 輸出資料夾已建立: '{output_dir}'")
    except OSError:
        return None

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_zip_path = Path(temp_dir) / "converted.zip"
        print("🔧 正在呼叫 Calibre 將 mobi 轉換為 zip...")
        try:
            command = ["ebook-convert", str(mobi_path), str(temp_zip_path)]
            subprocess.run(command, check=True, capture_output=True, text=True, encoding='utf-8')
            print("✅ Calibre 轉換成功！")
        except Exception as e:
            print(f"❌ 錯誤: Calibre 的 ebook-convert 工具執行失敗。{e}")
            return None
        
        print("🔧 正在從 zip 檔案中提取圖片...")
        try:
            with zipfile.ZipFile(temp_zip_path, 'r') as zip_ref:
                image_files = [f for f in zip_ref.namelist() if f.lower().endswith(('.jpg', '.jpeg', '.png', '.gif'))]
                if not image_files:
                    print("⚠️ 警告: 在轉換後的檔案中找不到任何圖片。")
                    return None
                
                for i, image_file_in_zip in enumerate(sorted(image_files), 1):
                    image_data = zip_ref.read(image_file_in_zip)
                    extension = Path(image_file_in_zip).suffix
                    dest_filename = f"{i:04d}{extension}"
                    dest_path = output_dir / dest_filename
                    with open(dest_path, 'wb') as f: f.write(image_data)
                
                print(f"✅ 提取成功！共找到並儲存 {len(image_files)} 張圖片。")
                return output_dir

        except zipfile.BadZipFile:
            print("❌ 錯誤: Calibre 輸出的不是一個有效的 zip 檔案。")
            return None

    return None

if __name__ == "__main__":
    # ... (主程式區塊維持不變) ...
    files_to_process = sys.argv[1:]

    if not files_to_process:
        print("="*60)
        print(" MOBI 漫畫處理工作流 v5.1 (Calibre + 可攜式KCC)")
        print("="*60)
        print(f"依賴工具：")
        print(f"  1. Calibre (需安裝並加入系統 PATH)")
        print(f"  2. KCC (kcc-c2e.exe 需與本腳本在同一資料夾)")
        print("\n用法說明：")
        print("  將一個或多個 .mobi 檔案拖曳到 .bat 批次檔上即可。\n")
        print("="*60)
    else:
        print(f"▶️ 偵測到 {len(files_to_process)} 個檔案，準備開始處理...\n")
        for i, file_path_str in enumerate(files_to_process, 1):
            mobi_file_path = Path(file_path_str)
            print(f"--- [ 處理中: {i}/{len(files_to_process)} ] ---")
            print(f"📄 檔案: {mobi_file_path.name}")
            
            unpacked_folder = unpack_comic_with_calibre(mobi_file_path)
            
            if unpacked_folder and REPACK_AFTER_UNPACK:
                repack_with_kcc(unpacked_folder)

            print("-" * (21 + len(str(i)) + len(str(len(files_to_process)))))
            print()

        print("✅ 所有任務皆已完成！")

    print("\n按下 Enter 鍵結束程式...")
    input()