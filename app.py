import streamlit as st
import pandas as pd
import re
import io
import difflib

def count_fullwidth_chars(text):
    """全角文字数としてカウントする（1000バイト＝500文字相当のチェック）"""
    if not isinstance(text, str):
        if pd.isna(text):
            return 0
        text = str(text)
    try:
        # cp932でエンコードしてバイト数を取得し、2で割って文字数相当とする
        byte_len = len(text.encode('cp932', errors='replace'))
        return byte_len // 2
    except Exception:
        return len(text)

def generate_ys_data(df, store_id):
    """店舗別のデータを生成する"""
    df_out = df.copy()
    
    # 対象列の取得: item-image-urls は置換の対象外
    replace_cols = [col for col in df_out.columns if col != 'item-image-urls']
    
    if store_id == 'YS1':
        return df_out
        
    if store_id == 'YS5':
        for col in replace_cols:
            mask = df_out[col].notna()
            if mask.any():
                df_out.loc[mask, col] = df_out.loc[mask, col].astype(str).str.replace('yahoo.co.jp/solltd', 'yahoo.co.jp/solltd5', regex=False)
        return df_out
        
    if store_id in ['YS2', 'YS3', 'YS4']:
        # common rule
        for col in replace_cols:
            mask = df_out[col].notna()
            if not mask.any():
                continue
                
            temp_col = df_out.loc[mask, col].astype(str)
            # 1. 送料185円 を空白（空文字列）に置換
            temp_col = temp_col.str.replace('送料185円', '', regex=False)
            
            # URL置換
            if store_id == 'YS2':
                temp_col = temp_col.str.replace('yahoo.co.jp/solltd', 'yahoo.co.jp/solltd2', regex=False)
            elif store_id == 'YS3':
                temp_col = temp_col.str.replace('yahoo.co.jp/solltd', 'yahoo.co.jp/solltd3', regex=False)
            elif store_id == 'YS4':
                temp_col = temp_col.str.replace('yahoo.co.jp/solltd', 'yahoo.co.jp/solltd4', regex=False)
            
            # 画像ファイル名置換
            temp_col = temp_col.str.replace('parts03.gif', 'parts23.gif', regex=False)
            temp_col = temp_col.str.replace('supplies03.gif', 'supplies23.gif', regex=False)

            if store_id == 'YS4':
                temp_col = temp_col.str.replace('lib/solltd/parts', 'lib/solltd4/parts', regex=False)
                temp_col = temp_col.str.replace('lib/solltd/supplies', 'lib/solltd4/supplies', regex=False)
                temp_col = temp_col.str.replace('lib/solltd/hapadparts', 'lib/solltd4/hapadparts', regex=False)

            df_out.loc[mask, col] = temp_col

        # codeの末尾にship-weightの値に応じた文字列を付与
        if 'code' in df_out.columns and 'ship-weight' in df_out.columns:
            def suffix_code(row):
                c = row['code']
                if pd.isna(c):
                    return c
                sw = str(row['ship-weight']).strip() if pd.notna(row['ship-weight']) else ""
                if sw.endswith('.0'):
                    sw = sw[:-2]
                
                c_str = str(c)
                if sw == '0':
                    return c_str + 'VVV'
                elif sw == '100':
                    return c_str + 'WWW'
                elif sw == '1':
                    return c_str + 'XXX'
                elif sw == '1000':
                    return c_str + 'YYY'
                return c_str
            
            df_out['code'] = df_out.apply(suffix_code, axis=1)
            
    return df_out

def to_csv_bytes(df):
    output = io.BytesIO()
    df.to_csv(output, index=False, encoding='cp932', errors='replace')
    return output.getvalue()

def filter_download_columns(df, processed_cols):
    """ダウンロード用に必須列(code, name)と処理済み列のみを抽出する"""
    cols_to_keep = []
    if 'code' in df.columns:
        cols_to_keep.append('code')
    if 'name' in df.columns:
        cols_to_keep.append('name')
        
    for col in processed_cols:
        if col in df.columns and col not in cols_to_keep:
            cols_to_keep.append(col)
            
    return df[cols_to_keep]

def main():
    st.set_page_config(page_title="AutoPage Updater", layout="wide")
    st.title("AutoPage Updater")
    st.write("商品データのテキスト置換および多店舗(YS1〜YS5)向け一括ファイル生成アプリケーション")

    # 1: ファイルアップロード
    uploaded_file = st.file_uploader("1: ベースデータのアップロード (CSV, xlsx対応)", type=["csv", "xlsx"])

    if uploaded_file is not None:
        if 'uploaded_file_id' not in st.session_state or st.session_state['uploaded_file_id'] != uploaded_file.file_id:
            with st.spinner("ファイルを読み込み中..."):
                try:
                    if uploaded_file.name.endswith('.csv'):
                        try:
                            df = pd.read_csv(uploaded_file, encoding='cp932', dtype=str)
                        except UnicodeDecodeError:
                            uploaded_file.seek(0)
                            df = pd.read_csv(uploaded_file, encoding='utf-8', dtype=str)
                    else:
                        df = pd.read_excel(uploaded_file, dtype=str)
                    
                    st.session_state['current_df'] = df
                    st.session_state['uploaded_file_id'] = uploaded_file.file_id
                    st.session_state['original_filename'] = uploaded_file.name
                    st.session_state['processed_columns'] = []
                    if 'warnings' in st.session_state:
                        del st.session_state['warnings']
                    st.success("ファイルを読み込み、ベースデータを保持しました。")
                except Exception as e:
                    st.error(f"ファイルの読み込みに失敗しました: {e}")

    if 'current_df' not in st.session_state:
        st.info("ファイルがアップロードされていません。")
        return

    df = st.session_state['current_df']

    st.markdown("---")
    st.header("テキスト置換処理")
    # 2: 処理対象の列名を選択するドロップダウン
    columns = list(df.columns)
    default_cols = [c for c in ["explanation", "additional1", "sp-additional"] if c in columns]
    if not default_cols and columns:
        default_cols = [columns[0]]
    
    target_column = st.selectbox("処理対象の列を選択", columns, index=columns.index(default_cols[0]) if default_cols else 0)

    # 3: テキスト入力エリア
    st.write("置換内容の指定")
    st.info("""💡 **Tips：新しい文言を追加（追記）したい場合**
新しい文言を追加したい場合は、置き換え機能を使って簡単に追加できます。

1: 「置き換え元」に、目印となる既存の文章をコピーして貼り付けます。
2: 「置き換え後」に、貼り付けた既存の文章と「追加したい文言」を一緒に入力します。

「置き換え後」の入力欄で作った改行や空白がそのまま実際のページに反映されるため、直感的な位置調整が可能です。""")
    col1, col2 = st.columns(2)
    with col1:
        original_text = st.text_area("置き換え元テキスト (プレーンテキストを入力)", height=200)
    with col2:
        replacement_text = st.text_area("置き換え後テキスト (プレーンテキストを入力)", height=200)

    # 4: 実行ボタン
    if st.button("テキスト置換を実行し、ベースデータを更新", type="primary"):
        if not original_text.strip():
            st.error("置き換え元テキストを入力してください。")
        else:
            warnings_list = []
            processed_df = df.copy()

            with st.spinner("置換処理中..."):
                for idx, row in processed_df.iterrows():
                    val = str(row[target_column]) if pd.notna(row[target_column]) else ""
                    
                    try:
                        orig_lines = original_text.splitlines()
                        repl_lines = replacement_text.splitlines()
                        matcher = difflib.SequenceMatcher(None, orig_lines, repl_lines)
                        
                        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
                            if tag == 'equal':
                                continue
                            
                            orig_chunk = "\n".join(orig_lines[i1:i2])
                            repl_chunk = "\n".join(repl_lines[j1:j2])
                            
                            if tag in ('replace', 'delete'):
                                if orig_chunk.strip():
                                    chars = [re.escape(c) for c in orig_chunk if not c.isspace()]
                                    if chars:
                                        pattern_str = r"(?:<[^>]+>|\s)*".join(chars)
                                        pattern = re.compile(pattern_str, re.IGNORECASE)
                                        repl = repl_chunk.replace("\n", "<br>") if repl_chunk else ""
                                        val = pattern.sub(lambda m: repl, val, count=1)
                                        
                            elif tag == 'insert':
                                if repl_chunk.strip():
                                    if i1 > 0:
                                        prev_line = orig_lines[i1 - 1]
                                        chars = [re.escape(c) for c in prev_line if not c.isspace()]
                                        if chars:
                                            pattern_str = r"(?:<[^>]+>|\s)*".join(chars)
                                            pattern = re.compile(pattern_str, re.IGNORECASE)
                                            repl = repl_chunk.replace("\n", "<br>")
                                            val = pattern.sub(lambda m: m.group(0) + "<br>" + repl, val, count=1)
                                    elif i2 < len(orig_lines):
                                        next_line = orig_lines[i2]
                                        chars = [re.escape(c) for c in next_line if not c.isspace()]
                                        if chars:
                                            pattern_str = r"(?:<[^>]+>|\s)*".join(chars)
                                            pattern = re.compile(pattern_str, re.IGNORECASE)
                                            repl = repl_chunk.replace("\n", "<br>")
                                            val = pattern.sub(lambda m: repl + "<br>" + m.group(0), val, count=1)
                    except Exception as e:
                        item_id = row.get("code", row.get("商品コード", f"行番号 {idx+1}"))
                        warnings_list.append(f"スキップ: {item_id} の置換処理でエラーが発生しました。詳細: {e}")
                        
                    processed_df.at[idx, target_column] = val

                    # explanation 文字数チェック
                    if target_column == "explanation":
                        char_count = count_fullwidth_chars(val)
                        if char_count > 500:
                            item_id = row.get("code", row.get("商品コード", f"行番号 {idx+1}"))
                            warnings_list.append(f"警告: {item_id} の「explanation」が全角500文字(1000バイト)を超過しています。(推定 {char_count} 文字)")

            st.session_state['current_df'] = processed_df
            st.session_state['warnings'] = warnings_list
            
            if 'processed_columns' not in st.session_state:
                st.session_state['processed_columns'] = []
            if target_column not in st.session_state['processed_columns']:
                st.session_state['processed_columns'].append(target_column)

            st.success("テキスト置換が完了し、ベースデータが更新されました！")

    if st.session_state.get('warnings'):
        st.warning("最新の処理で一部のデータに警告やスキップが発生しています。詳細は以下をご確認ください。")
        for w in st.session_state['warnings']:
            st.write(f"- {w}")

    st.markdown("---")
    st.header("店舗別ダウンロード")
    st.write("最新のベースデータをもとに、各店舗(YS1〜YS5)向けのファイルを動的生成してダウンロードします。")

    # ダウンロードセクション
    col_ys1, col_ys2, col_ys3, col_ys4, col_ys5 = st.columns(5)
    
    current_df = st.session_state['current_df']
    orig_name = st.session_state.get('original_filename', 'data.csv')
    base_name = orig_name.rsplit('.', 1)[0]
    processed_cols = st.session_state.get('processed_columns', [])

    with col_ys1:
        st.subheader("YS1")
        st.write("ベースデータそのまま")
        df_ys1 = generate_ys_data(current_df, 'YS1')
        df_ys1_dl = filter_download_columns(df_ys1, processed_cols)
        st.download_button(
            label="📥 YS1 ダウンロード",
            data=to_csv_bytes(df_ys1_dl),
            file_name=f"{base_name}_YS1.csv",
            mime="text/csv",
            key="btn_ys1"
        )
        
    with col_ys2:
        st.subheader("YS2")
        st.write("URL・画像置換、code付与")
        df_ys2 = generate_ys_data(current_df, 'YS2')
        df_ys2_dl = filter_download_columns(df_ys2, processed_cols)
        st.download_button(
            label="📥 YS2 ダウンロード",
            data=to_csv_bytes(df_ys2_dl),
            file_name=f"{base_name}_YS2.csv",
            mime="text/csv",
            key="btn_ys2"
        )
        
    with col_ys3:
        st.subheader("YS3")
        st.write("YS2と同等 + URL(YS3)")
        df_ys3 = generate_ys_data(current_df, 'YS3')
        df_ys3_dl = filter_download_columns(df_ys3, processed_cols)
        st.download_button(
            label="📥 YS3 ダウンロード",
            data=to_csv_bytes(df_ys3_dl),
            file_name=f"{base_name}_YS3.csv",
            mime="text/csv",
            key="btn_ys3"
        )
        
    with col_ys4:
        st.subheader("YS4")
        st.write("YS2と同等 + 特別URL置換")
        df_ys4 = generate_ys_data(current_df, 'YS4')
        df_ys4_dl = filter_download_columns(df_ys4, processed_cols)
        st.download_button(
            label="📥 YS4 ダウンロード",
            data=to_csv_bytes(df_ys4_dl),
            file_name=f"{base_name}_YS4.csv",
            mime="text/csv",
            key="btn_ys4"
        )
        
    with col_ys5:
        st.subheader("YS5")
        st.write("全体URL置換(solltd5)")
        df_ys5 = generate_ys_data(current_df, 'YS5')
        df_ys5_dl = filter_download_columns(df_ys5, processed_cols)
        st.download_button(
            label="📥 YS5 ダウンロード",
            data=to_csv_bytes(df_ys5_dl),
            file_name=f"{base_name}_YS5.csv",
            mime="text/csv",
            key="btn_ys5"
        )

if __name__ == "__main__":
    main()
