import os
import io
from flask import Flask, request, render_template, send_file, redirect, url_for, flash
from pptx import Presentation
from deep_translator import GoogleTranslator, DeeplTranslator
import tempfile
import logging

# ロギング設定
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = "ppt_translator_secret_key"

# DeepL APIキー（環境変数から取得する例）
DEEPL_API_KEY = os.environ.get("DEEPL_API_KEY", "")

# 翻訳関数（DeepL APIを使用）
def translate_text_deepl(text, source_lang, target_lang):
    if not text or text.strip() == "":
        return text
    
    # 入力が確実に文字列型であることを確認
    text = str(text).strip()
    
    try:
        # DeepL APIキーが設定されている場合はDeepLを使用
        if DEEPL_API_KEY:
            translator = DeeplTranslator(api_key=DEEPL_API_KEY)
            result = translator.translate(text, source=source_lang, target=target_lang)
        else:
            # APIキーがない場合はGoogleを使用
            translator = GoogleTranslator(source=source_lang, target=target_lang)
            result = translator.translate(text)
        
        # 結果が文字列型であることを確認
        if result is None:
            logger.warning(f"翻訳結果がNoneです。元のテキストを返します: {text}")
            return text
        
        return str(result)
    except Exception as e:
        logger.error(f"翻訳エラー: {e}, テキスト: {text}")
        return text  # エラーの場合は元のテキストを返す

# PowerPointファイルの翻訳処理
def translate_pptx(input_file, source_lang, target_lang):
    try:
        # プレゼンテーションの読み込み
        logger.info(f"プレゼンテーションを読み込み中: {input_file}")
        prs = Presentation(input_file)
        
        # スライド数のログ
        logger.info(f"スライド数: {len(prs.slides)}")
        
        # 各スライドを処理
        for i, slide in enumerate(prs.slides):
            logger.info(f"スライド {i+1} を処理中...")
            
            # テキストフレームを持つすべての図形を処理
            for shape in slide.shapes:
                # テキストフレームがある場合
                if hasattr(shape, "text_frame") and shape.text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        # 各テキスト実行を処理
                        for run in paragraph.runs:
                            if run.text and run.text.strip():
                                # 翻訳
                                original_text = run.text
                                translated_text = translate_text_deepl(original_text, source_lang, target_lang)
                                logger.debug(f"翻訳: '{original_text}' -> '{translated_text}'")
                                run.text = translated_text
                
                # 表がある場合
                if hasattr(shape, "table"):
                    try:
                        for row in shape.table.rows:
                            for cell in row.cells:
                                if cell.text_frame:
                                    for paragraph in cell.text_frame.paragraphs:
                                        for run in paragraph.runs:
                                            if run.text and run.text.strip():
                                                # 翻訳
                                                original_text = run.text
                                                translated_text = translate_text_deepl(original_text, source_lang, target_lang)
                                                logger.debug(f"表の翻訳: '{original_text}' -> '{translated_text}'")
                                                run.text = translated_text
                    except Exception as table_e:
                        logger.error(f"表の処理中にエラーが発生しました: {table_e}")
        
        # 翻訳済みプレゼンテーションを一時ファイルに保存
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
        temp_file.close()  # 明示的にクローズしてからsave
        
        logger.info(f"翻訳済みプレゼンテーションを保存中: {temp_file.name}")
        prs.save(temp_file.name)
        
        return temp_file.name
    except Exception as e:
        logger.error(f"翻訳処理中にエラーが発生しました: {e}")
        raise

# Webインターフェース
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # ファイルが提供されているか確認
        if 'file' not in request.files:
            flash('ファイルがありません')
            return redirect(request.url)
            
        file = request.files['file']
        
        # ファイル名が空でないことを確認
        if file.filename == '':
            flash('ファイルが選択されていません')
            return redirect(request.url)
            
        # 有効なファイル拡張子を確認
        if not file.filename.endswith('.pptx'):
            flash('PowerPointファイル(.pptx)のみ対応しています')
            return redirect(request.url)
            
        # 翻訳方向の取得
        translation_direction = request.form.get('direction', 'ja-en')
        source_lang, target_lang = translation_direction.split('-')
        
        try:
            # 一時ファイルに保存
            input_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
            input_temp_path = input_temp.name
            input_temp.close()
            
            file.save(input_temp_path)
            logger.info(f"アップロードされたファイルを保存しました: {input_temp_path}")
            
            # 翻訳処理
            output_file = translate_pptx(input_temp_path, source_lang, target_lang)
            
            # 一時入力ファイルの削除
            try:
                os.unlink(input_temp_path)
                logger.info(f"一時ファイルを削除しました: {input_temp_path}")
            except Exception as unlink_e:
                logger.warning(f"一時ファイルの削除に失敗しました: {unlink_e}")
            
            # 出力ファイル名の設定
            language_prefix = "en" if target_lang == "en" else "ja"
            output_filename = f"{language_prefix}_{file.filename}"

            logger.info(f"翻訳済みファイルを返します: {output_file} -> {output_filename}")

            # ファイルを返す際に、特別なヘッダーを設定
            response = send_file(output_file, 
                            as_attachment=True,
                            download_name=output_filename,
                            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation')

            # カスタムヘッダーを追加（これをJavaScriptで検知する）
            response.headers["X-Translation-Complete"] = "true"
            return response
                            
        except Exception as e:
            error_message = f'エラーが発生しました: {str(e)}'
            flash(error_message)
            logger.error(f"処理エラー: {e}")
            return redirect(request.url)
            
    return render_template('index.html')

if __name__ == '__main__':
    # テンプレートディレクトリの作成
    os.makedirs('templates', exist_ok=True)
    
    # HTMLテンプレートの作成（シンプルなローディングエフェクト付き）
    with open('templates/index.html', 'w') as f:
        f.write('''
<!DOCTYPE html>
<html>
<head>
    <title>PowerPoint翻訳ツール</title>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
        }
        .container {
            background-color: #f9f9f9;
            border-radius: 10px;
            padding: 20px;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
            position: relative;
        }
        h1 {
            color: #333;
            text-align: center;
        }
        .form-group {
            margin-bottom: 15px;
        }
        label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
        }
        input[type="file"], select {
            width: 100%;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
        }
        button {
            background-color: #4CAF50;
            color: white;
            padding: 10px 15px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
            transition: background-color 0.3s;
        }
        button:hover {
            background-color: #45a049;
        }
        button:disabled {
            background-color: #cccccc;
            cursor: not-allowed;
        }
        .flash-messages {
            margin-bottom: 20px;
        }
        .flash-message {
            background-color: #f8d7da;
            color: #721c24;
            padding: 10px;
            border-radius: 4px;
            margin-bottom: 10px;
        }
        
        /* ロードのアニメーション */
        #loading {
            width: 100vw;
            height: 100vh;
            transition: all 1s;
            background-color: rgba(0, 187, 221, 0.7);
            position: fixed;
            top: 0;
            left: 0;
            z-index: 9999;
            display: none;
        }
        .spinner {
            width: 100px;
            height: 100px;
            margin: 200px auto;
            background-color: #fff;
            border-radius: 100%;
            animation: sk-scaleout 1s infinite ease-in-out;
        }
        /* ローディングアニメーション */
        @keyframes sk-scaleout {
            0% {
                transform: scale(0);
            }
            100% {
                transform: scale(1);
                opacity: 0;
            }
        }
        
        /* ローディングのテキスト */
        .loading-text {
            text-align: center;
            color: white;
            font-size: 24px;
            font-weight: bold;
            margin-top: 20px;
        }
    </style>
</head>
<body>
    <!-- ローディング表示 -->
    <div id="loading">
        <div class="spinner"></div>
        <div class="loading-text">翻訳処理中...</div>
    </div>
    
    <div class="container">
        <h1>PowerPoint翻訳ツール</h1>
        
        {% if get_flashed_messages() %}
        <div class="flash-messages">
            {% for message in get_flashed_messages() %}
            <div class="flash-message">{{ message }}</div>
            {% endfor %}
        </div>
        {% endif %}
        
        <form method="post" enctype="multipart/form-data" id="translationForm">
            <div class="form-group">
                <label for="file">PowerPointファイル (.pptx):</label>
                <input type="file" id="file" name="file" accept=".pptx" required>
            </div>
            
            <div class="form-group">
                <label for="direction">翻訳方向:</label>
                <select id="direction" name="direction">
                    <option value="ja-en">日本語 → 英語</option>
                    <option value="en-ja">英語 → 日本語</option>
                </select>
            </div>
            
            <button type="submit" id="submitButton">翻訳</button>
        </form>
    </div>
    
    <script>
        $(document).ready(function() {
            // フォーム送信時の処理
            $('#translationForm').on('submit', function(e) {
                e.preventDefault(); // 通常のフォーム送信を防止
                
                // ローディング表示
                $('#loading').show();
                
                // FormDataオブジェクトの作成
                const formData = new FormData(this);
                
                // FetchAPIでフォームを送信
                fetch('/', {
                    method: 'POST',
                    body: formData
                })
                .then(response => {
                    // 完了ヘッダーがあるか確認
                    if (response.headers.get('X-Translation-Complete') === 'true') {
                        // ローディングを非表示
                        $('#loading').hide();
                    }
                    
                    // ファイルのダウンロード処理
                    return response.blob();
                })
                .then(blob => {
                    // ダウンロードリンクの作成と自動クリック
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = formData.get('file').name;
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                })
                .catch(error => {
                    // エラー処理
                    $('#loading').hide();
                    console.error('Error:', error);
                });
            });
        });
    </script>
</body>
</html>
        ''')
    
    app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 5001)))