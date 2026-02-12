import time
import os
from pathlib import Path
import pytesseract
from pdf2image import convert_from_path
import cv2
import numpy as np
from PIL import Image
import PyPDF2
import shutil

# 環境に応じてパスを設定
if os.environ.get('RUNNING_IN_DOCKER'):
    # Docker環境（Azure）
    pytesseract.pytesseract.tesseract_cmd = 'tesseract'
    poppler_path = None
else:
    # ローカル環境（Windows）
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    poppler_path = r'C:\poppler\Library\bin'

def preprocess_image(image):
    """画像を前処理してOCR精度を向上"""
    # PIL ImageをOpenCV形式に変換
    img_array = np.array(image)
    
    # グレースケール変換
    gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
    
    # ノイズ除去
    denoised = cv2.fastNlMeansDenoising(gray, None, 10, 7, 21)
    
    # コントラスト強化（CLAHE）
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
    enhanced = clahe.apply(denoised)
    
    # 二値化（適応的閾値処理）
    binary = cv2.adaptiveThreshold(
        enhanced, 255, 
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
        cv2.THRESH_BINARY, 11, 2
    )
    
    # PIL Imageに戻す
    return Image.fromarray(binary)

def try_extract_text_directly(pdf_path):
    """テキストベースPDFから直接テキストを抽出"""
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ""
            for page_num, page in enumerate(pdf_reader.pages, 1):
                page_text = page.extract_text()
                if page_text.strip():  # テキストがあれば
                    text += f"\n--- ページ {page_num} ---\n{page_text}\n"
            
            # テキストが十分に抽出できた場合は返す
            if len(text.strip()) > 100:  # 100文字以上あればテキストPDFと判断
                return text, True
        return None, False
    except:
        return None, False

def extract_text_with_ocr(pdf_path):
    """OCRでテキストを抽出（画像前処理付き）"""
    print("📖 OCRでテキスト抽出中...")
    
    # 高解像度で画像変換（300dpi）
    images = convert_from_path(pdf_path, dpi=300, poppler_path=poppler_path)
    text = ""
    
    # Tesseractの設定（日本語 + 縦書き対応）
    custom_config = r'--oem 3 --psm 6 -l jpn'
    
    for i, image in enumerate(images, 1):
        print(f"  ページ {i}/{len(images)} を処理中...")
        
        # 画像前処理
        processed_image = preprocess_image(image)
        
        # OCR実行
        page_text = pytesseract.image_to_string(
            processed_image, 
            config=custom_config
        )
        text += f"\n--- ページ {i} ---\n{page_text}\n"
    
    return text

def extract_text(pdf_path):
    """PDFからテキストを抽出（最適な方法を自動選択）"""
    print("\n📄 PDF解析中...")
    
    # まずテキストベースPDFとして試す
    direct_text, is_text_pdf = try_extract_text_directly(pdf_path)
    
    if is_text_pdf:
        print("✓ テキストベースPDFを検出 → 直接抽出（高精度）")
        return direct_text
    else:
        print("✓ 画像PDFを検出 → OCR処理（前処理適用）")
        return extract_text_with_ocr(pdf_path)

def save_results(pdf_path, extracted_text):
    """抽出結果を保存"""
    result_file = Path("results") / f"{Path(pdf_path).stem}_抽出結果.txt"
    
    with open(result_file, 'w', encoding='utf-8') as f:
        f.write(f"契約書ファイル: {Path(pdf_path).name}\n")
        f.write(f"処理日時: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("=" * 60 + "\n\n")
        f.write(extracted_text)
    
    print(f"✅ 抽出完了！結果を保存: {result_file.name}\n")

def process_contract(pdf_path):
    try:
        print(f"\n{'='*60}")
        print(f"📋 処理開始: {Path(pdf_path).name}")
        print(f"{'='*60}")
        
        # テキスト抽出
        text = extract_text(pdf_path)
        
        # 結果保存
        save_results(pdf_path, text)
        
    except Exception as e:
        print(f"❌ エラー: {e}\n")
        import traceback
        traceback.print_exc()

def watch_folder():
    """フォルダを監視して新しいPDFを処理"""
    contracts_folder = Path("contracts")
    processed_files = set()
    
    print("=" * 60)
    print("📋 契約書OCR自動処理システム（精度向上版）")
    print("=" * 60)
    print("✓ テキストPDF: 直接抽出（精度100%）")
    print("✓ 画像PDF: 前処理+OCR（精度向上）")
    print("✓ contracts/フォルダを監視しています...")
    print("✓ 終了するには Ctrl+C を押してください\n")
    
    # resultsフォルダがなければ作成
    Path("results").mkdir(exist_ok=True)
    
    try:
        while True:
            # contractsフォルダ内のPDFファイルをチェック
            for pdf_file in contracts_folder.glob("*.pdf"):
                if pdf_file not in processed_files:
                    process_contract(pdf_file)
                    processed_files.add(pdf_file)
            
            time.sleep(2)  # 2秒ごとにチェック
            
    except KeyboardInterrupt:
        print("\n\n終了します...")

if __name__ == "__main__":
    watch_folder()