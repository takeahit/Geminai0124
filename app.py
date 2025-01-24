import pandas as pd
from rapidfuzz import fuzz, process
from docx import Document
from docx.shared import RGBColor
from io import BytesIO
from pydocx import PyDocX
import streamlit as st
from PyPDF2 import PdfReader
import re
import streamlit.components.v1 as components


# --- エラーハンドリングとログ記録 ---
def log_error(message):
    st.error(message)
    # 必要であればログファイルにも出力する

# --- データクリーニング ---
def clean_strings(df):
    def clean_cell(value):
        if isinstance(value, str):
            return re.sub(r'[\x00-\x1F\x7F]', '', value)
        return value
    return df.applymap(clean_cell)

def find_invalid_chars(df):
    invalid_rows = []
    for col in df.columns:
        for idx, value in df[col].items():
            if isinstance(value, str) and re.search(r'[\x00-\x1F\x7F]', value):
                invalid_rows.append((col, idx, value))
    return invalid_rows

# --- ファイル読み込み ---
def load_excel(file):
    try:
        df = pd.read_excel(file, engine="openpyxl")
        if df.columns.size < 1:
            raise ValueError("Excelファイルには少なくとも1列以上の用語が必要です。")
        return df
    except Exception as e:
        log_error(f"Excelファイルの読み込み中にエラーが発生しました: {e}")
        return None

def extract_text_from_file(file, file_type):
    try:
        if file_type == "docx":
            doc = Document(file)
            return "\n".join([paragraph.text for paragraph in doc.paragraphs])
        elif file_type == "doc":
            return PyDocX.to_text(file)
        elif file_type == "pdf":
            reader = PdfReader(file)
            text = ""
            for page in reader.pages:
                page_text = page.extract_text()
                page_text = page_text.replace("\n", " ").replace("\r", " ")
                page_text = " ".join(page_text.split())
                text += page_text + " "
            text = text.strip()
            return text
        else:
            return ""
    except Exception as e:
        log_error(f"ファイルからのテキスト抽出中にエラーが発生しました: {e}")
        return ""

# --- Fuzzy Matching ---
def find_similar_terms(text, terms, threshold):
    words = text.split()
    detected_terms = []
    for word in words:
        matches = process.extract(word, terms, scorer=fuzz.partial_ratio, limit=10)
        for match in matches:
            if match[1] >= threshold and match[1] < 100:
                detected_terms.append((word, match[0], match[1]))
    return detected_terms

# --- 修正処理 ---
def apply_corrections(text, corrections):
    total_replacements = 0
    for incorrect, correct in corrections:
        max_replacements = text.count(incorrect)
        for _ in range(max_replacements):
            if incorrect in text:
                text = text.replace(incorrect, correct, 1)
                total_replacements += 1
            else:
                break
    return text, total_replacements

def create_corrected_word_file_with_formatting(original_text, corrections):
    doc = Document()
    for paragraph_text in original_text.split("\n"):
        paragraph = doc.add_paragraph()
        start_index = 0
        for incorrect, correct in corrections:
            while incorrect in paragraph_text[start_index:]:
                start_index = paragraph_text.find(incorrect, start_index)
                end_index = start_index + len(incorrect)
                paragraph.add_run(paragraph_text[:start_index])
                run = paragraph.add_run(correct)
                run.font.highlight_color = 6
                paragraph_text = paragraph_text[end_index:]
                start_index = 0
        paragraph.add_run(paragraph_text)
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# --- データ表示とダウンロード ---
def create_correction_table(detected):
    return pd.DataFrame(detected, columns=["原稿内の語", "類似する用語", "類似度"])

def download_excel(df, file_name, sheet_name):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    st.download_button(
        label=f"{sheet_name}をダウンロード",
        data=output.getvalue(),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

def download_word(file, file_name):
     st.download_button(
        label="修正済みファイルをダウンロード",
        data=file.getvalue(),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

# --- メイン処理関数 ---
def process_file(word_file, terms_file, correction_file, kanji_file):
    file_type = word_file.name.split(".")[-1]
    original_text = extract_text_from_file(word_file, file_type)

    if not original_text:
        return

    all_corrections = []

    # 用語集の処理
    if terms_file:
        terms_df = load_excel(terms_file)
        if terms_df is not None:
            terms_df = clean_strings(terms_df)
            terms = terms_df.iloc[:, 0].dropna().astype(str).tolist()
            threshold = st.slider("類似度の閾値を設定してください (50-99):", min_value=50, max_value=99, value=65)
            detected = find_similar_terms(original_text, terms, threshold)
            if detected:
                st.success(f"類似語が{len(detected)}件検出されました！")
                correction_table = create_correction_table(detected)
                st.dataframe(correction_table)
                download_excel(correction_table, "修正箇所一覧.xlsx", "修正箇所一覧")

    # 正誤表の処理
    if correction_file:
        correction_df = load_excel(correction_file)
        if correction_df is not None:
            correction_df = clean_strings(correction_df)
            invalid_chars = find_invalid_chars(correction_df)
            if invalid_chars:
                st.error(f"非互換文字が検出されました: {invalid_chars}")
            else:
                corrections = list(correction_df.itertuples(index=False, name=None))
                updated_text, total_replacements = apply_corrections(original_text, corrections)
                all_corrections.extend(corrections)
                st.success(f"正誤表を適用し、{total_replacements}回の修正を行いました！")

                corrections_df = pd.DataFrame(corrections, columns=["誤った用語", "正しい用語"])
                st.dataframe(corrections_df)
                download_excel(corrections_df, "正誤表修正箇所.xlsx", "正誤表修正箇所")

                corrected_file = create_corrected_word_file_with_formatting(original_text, corrections)
                download_word(corrected_file, "正誤表修正済み.docx")


    # 利用漢字表の処理
    if kanji_file:
        kanji_df = load_excel(kanji_file)
        if kanji_df is not None:
            kanji_df = clean_strings(kanji_df)
            corrections = list(kanji_df.itertuples(index=False, name=None))
            updated_text, total_replacements = apply_corrections(original_text, corrections)
            all_corrections.extend(corrections)
            st.success(f"利用漢字表を適用し、{total_replacements}回の修正を行いました！")

            kanji_corrections_df = pd.DataFrame(corrections, columns=["ひらがな", "漢字"])
            st.dataframe(kanji_corrections_df)
            download_excel(kanji_corrections_df, "利用漢字表修正箇所.xlsx", "漢字修正箇所")

            corrected_file = create_corrected_word_file_with_formatting(original_text, corrections)
            download_word(corrected_file, "利用漢字表修正済み.docx")


# --- Streamlit アプリケーション ---
import streamlit as st
import streamlit.components.v1 as components

st.markdown("<h1 style='text-align: center;'>南江堂用用語チェッカー（笑）</h1>", unsafe_allow_html=True)
if "dify_initialized" not in st.session_state:
    dify_html = """
        <script>
            window.difyChatbotConfig = {
             token: 'rGMuWhHEu9Hcwbqe'
            };
            document.addEventListener('DOMContentLoaded', function() {
              var script = document.createElement('script');
              script.src = 'https://udify.app/embed.min.js';
              script.id = 'rGMuWhHEu9Hcwbqe';
              script.defer = true;
              document.head.appendChild(script);
            });
          </script>
          <style>
            #dify-chatbot-bubble-button {
              background-color: #1C64F2 !important;
            }
            #dify-chatbot-bubble-window {
              width: 24rem !important;
              height: 40rem !important;
            }
          </style>
    """
    components.html(dify_html, height=0)
    st.session_state["dify_initialized"] = True

st.write("以下のファイルを個別にアップロードしてください:")
word_file = st.file_uploader("原稿ファイル (Word, DOC, PDF):", type=["docx", "doc", "pdf"])
terms_file = st.file_uploader("用語集ファイル (A列に正しい用語を記載したExcel):", type=["xlsx"])
correction_file = st.file_uploader("正誤表ファイル (A列に誤った用語、B列に正しい用語を記載したExcel):", type=["xlsx"])
kanji_file = st.file_uploader("利用漢字表ファイル (A列にひらがな、B列に漢字を記載したExcel):", type=["xlsx"])

# ファイルサイズの制限 (100MB以下)
max_size = 100 * 1024 * 1024
for file, name in [(word_file, "原稿ファイル"), (terms_file, "用語集ファイル"), (correction_file, "正誤表ファイル"), (kanji_file, "利用漢字表ファイル")]:
    if file and file.size > max_size:
        st.error(f"{name}のサイズが大きすぎます（100MB以下にしてください）。")
        st.stop()

if word_file and (terms_file or correction_file or kanji_file):
    process_file(word_file, terms_file, correction_file, kanji_file)
else:
    st.warning("原稿ファイルと、用語集、正誤表、利用漢字表のいずれかをアップロードしてください！")
