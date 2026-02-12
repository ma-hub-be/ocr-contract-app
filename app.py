from flask import Flask, render_template, request, redirect, url_for, send_file
from contract_auto import extract_text, preprocess_image
from pathlib import Path
import time
import PyPDF2
import os
from functiontools import wraps
from flask import request, Response

def check_auth(username, password):
    return username == 'maika' and password == 'perogostini'

def authenticate():
    return Response(
        'ログインが必要です', 401,
        {'WWW-Authenticate': 'Basic realm="Login Required"'}
    )        
def requires_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return authenticate()
        return f(*args, **kwargs)
    return decorated

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'

# アップロードフォルダの設定
UPLOAD_FOLDER = Path('uploads')
UPLOAD_FOLDER.mkdir(exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB上限

def extract_text_from_file(filepath):
    """
    ファイルからテキストを抽出する共通関数
    PDF、画像、Word、Excelに対応
    """
    file_ext = Path(filepath).suffix.lower()
    
    if file_ext == '.pdf':
        # PDF処理
        return extract_text(str(filepath))
    
    elif file_ext in ['.jpg', '.jpeg', '.png', '.tiff', '.tif', '.bmp']:
        # 画像ファイル処理
        from PIL import Image
        import pytesseract
        
        image = Image.open(filepath)
        processed_image = preprocess_image(image)
        
        custom_config = r'--oem 3 --psm 6 -l jpn'
        return pytesseract.image_to_string(processed_image, config=custom_config)
    
    elif file_ext == '.docx':
        # Word文書処理
        from docx import Document
        
        doc = Document(filepath)
        extracted_text = ""
        
        for para in doc.paragraphs:
            extracted_text += para.text + "\n"
        
        for table in doc.tables:
            for row in table.rows:
                row_text = "\t".join([cell.text for cell in row.cells])
                extracted_text += row_text + "\n"
        
        return extracted_text
    
    elif file_ext in ['.xlsx', '.xls']:
        # Excel処理
        from openpyxl import load_workbook
        
        wb = load_workbook(filepath, data_only=True)
        extracted_text = ""
        
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            extracted_text += f"\n=== シート: {sheet_name} ===\n\n"
            
            for row in ws.iter_rows(values_only=True):
                row_text = "\t".join([str(cell) if cell is not None else "" for cell in row])
                if row_text.strip():
                    extracted_text += row_text + "\n"
        
        return extracted_text
    
    else:
        raise ValueError(f"未対応のファイル形式です: {file_ext}")

@app.route('/')
@requires_auth
def index():
    """アップロード画面"""
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
@requires_auth
def upload_file():
    """ファイルをアップロードしてOCR/テキスト抽出処理"""
    if 'file' not in request.files:
        return "ファイルが選択されていません", 400
    
    file = request.files['file']
    
    if file.filename == '':
        return "ファイルが選択されていません", 400
    
    # 対応ファイル形式の確認
    allowed_extensions = ['.pdf', '.jpg', '.jpeg', '.png', '.tiff', '.tif', '.bmp', '.docx', '.xlsx', '.xls']
    file_ext = Path(file.filename).suffix.lower()

    if file_ext not in allowed_extensions:
        return f"対応していないファイル形式です。対応形式: {', '.join(allowed_extensions)}", 400

    # ファイルを保存
    filename = f"{int(time.time())}_{file.filename}"
    filepath = app.config['UPLOAD_FOLDER'] / filename
    file.save(filepath)
    
    # OCR/テキスト抽出処理
    try:
        # 共通関数を使用してテキスト抽出
        extracted_text = extract_text_from_file(filepath)
        
        # ファイル種類の表示名を設定
        if file_ext == '.pdf':
            try:
                with open(filepath, 'rb') as f:
                    pdf_reader = PyPDF2.PdfReader(f)
                    text_check = pdf_reader.pages[0].extract_text()
                    file_type = "テキストPDF" if len(text_check.strip()) > 100 else "画像PDF"
            except:
                file_type = "画像PDF"
        elif file_ext in ['.jpg', '.jpeg', '.png', '.tiff', '.tif', '.bmp']:
            file_type = f"画像ファイル ({file_ext.upper()})"
        elif file_ext == '.docx':
            file_type = "Word文書 (.docx)"
        elif file_ext in ['.xlsx', '.xls']:
            file_type = f"Excelファイル ({file_ext.upper()})"
        
        print(f"ファイル種類: {file_type}")
        print(f"抽出文字数: {len(extracted_text)}")
        print(f"最初の100文字: {extracted_text[:100] if len(extracted_text) > 0 else '(空)'}")
        
        # 結果をテキストファイルに保存
        Path("results").mkdir(exist_ok=True)
        result_file = Path("results") / f"抽出結果_{int(time.time())}.txt"
        
        with open(result_file, 'w', encoding='utf-8') as f:
            f.write(f"ファイル名: {file.filename}\n")
            f.write(f"ファイル種類: {file_type}\n")
            f.write(f"処理日時: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 60 + "\n\n")
            f.write(extracted_text)
        
        print(f"結果ファイル: {result_file.name}")
        
        # アップロードファイルを削除
        if filepath.exists():
            os.remove(filepath)
        
        return render_template('result.html', 
                             text=extracted_text, 
                             pdf_type=file_type,
                             filename=result_file.name)
    
    except Exception as e:
        # エラー時もファイル削除
        if filepath.exists():
            os.remove(filepath)
        
        print(f"エラー詳細: {str(e)}")
        import traceback
        traceback.print_exc()
        
        return f"処理エラー: {str(e)}", 500

@app.route('/download/<filename>')
@requires_auth
def download_file(filename):
    """結果ファイルをダウンロード"""
    file_path = Path("results") / filename
    if file_path.exists():
        return send_file(file_path, as_attachment=True)
    return "ファイルが見つかりません", 404

@app.route('/compare')
@requires_auth
def compare_page():
    """比較画面"""
    return render_template('compare.html')

@app.route('/compare_upload', methods=['POST'])
@requires_auth
def compare_upload():
    """2つのファイルを比較（全ファイル形式対応）"""
    if 'file1' not in request.files or 'file2' not in request.files:
        return "2つのファイルを選択してください", 400
    
    file1 = request.files['file1']
    file2 = request.files['file2']
    
    if file1.filename == '' or file2.filename == '':
        return "2つのファイルを選択してください", 400
    
    # ファイル保存
    filepath1 = app.config['UPLOAD_FOLDER'] / f"{int(time.time())}_1_{file1.filename}"
    filepath2 = app.config['UPLOAD_FOLDER'] / f"{int(time.time())}_2_{file2.filename}"
    file1.save(filepath1)
    file2.save(filepath2)
    
    try:
        # テキスト抽出（共通関数を使用）
        text1 = extract_text_from_file(filepath1)
        text2 = extract_text_from_file(filepath2)
        
        # 行ごとに分割
        lines1 = text1.splitlines()
        lines2 = text2.splitlines()
        
        import difflib
        matcher = difflib.SequenceMatcher(None, lines1, lines2)
        
        # 行ごとの差分情報を作成
        highlighted_lines1 = []
        highlighted_lines2 = []
        
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                # 同じ行
                for i in range(i1, i2):
                    highlighted_lines1.append({
                        'type': 'normal',
                        'html': lines1[i]
                    })
                for j in range(j1, j2):
                    highlighted_lines2.append({
                        'type': 'normal',
                        'html': lines2[j]
                    })
            
            elif tag == 'delete':
                # 削除された行
                for i in range(i1, i2):
                    highlighted_lines1.append({
                        'type': 'delete',
                        'html': lines1[i]
                    })
            
            elif tag == 'insert':
                # 追加された行
                for j in range(j1, j2):
                    highlighted_lines2.append({
                        'type': 'insert',
                        'html': lines2[j]
                    })
            
            elif tag == 'replace':
                # 変更された行 - 文字単位で比較
                old_lines = lines1[i1:i2]
                new_lines = lines2[j1:j2]
                
                # 行ペアごとに文字単位の差分を計算
                for idx in range(max(len(old_lines), len(new_lines))):
                    if idx < len(old_lines):
                        old_line = old_lines[idx]
                    else:
                        old_line = ""
                    
                    if idx < len(new_lines):
                        new_line = new_lines[idx]
                    else:
                        new_line = ""
                    
                    # 文字単位の差分
                    char_matcher = difflib.SequenceMatcher(None, old_line, new_line)
                    
                    old_html = ""
                    new_html = ""
                    
                    for ctag, ci1, ci2, cj1, cj2 in char_matcher.get_opcodes():
                        if ctag == 'equal':
                            old_html += old_line[ci1:ci2]
                            new_html += new_line[cj1:cj2]
                        elif ctag == 'delete':
                            old_html += f'<span class="char-delete">{old_line[ci1:ci2]}</span>'
                        elif ctag == 'insert':
                            new_html += f'<span class="char-insert">{new_line[cj1:cj2]}</span>'
                        elif ctag == 'replace':
                            old_html += f'<span class="char-change">{old_line[ci1:ci2]}</span>'
                            new_html += f'<span class="char-change">{new_line[cj1:cj2]}</span>'
                    
                    if old_html:
                        highlighted_lines1.append({
                            'type': 'change',
                            'html': old_html
                        })
                    
                    if new_html:
                        highlighted_lines2.append({
                            'type': 'change',
                            'html': new_html
                        })
        
        return render_template('compare_result_char.html',
                             file1_name=file1.filename,
                             file2_name=file2.filename,
                             highlighted_lines1=highlighted_lines1,
                             highlighted_lines2=highlighted_lines2)
    
    except Exception as e:
        print(f"比較エラー: {str(e)}")
        import traceback
        traceback.print_exc()
        return f"比較処理エラー: {str(e)}", 500
    
    finally:
        # ファイル削除
        if filepath1.exists():
            os.remove(filepath1)
        if filepath2.exists():
            os.remove(filepath2)

if __name__ == '__main__':

    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)


