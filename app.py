import pandas as pd
import requests
import streamlit as st
from io import BytesIO
import re

LANGUAGE_OPTIONS_UI = {
    "Shqipe": "sq",
    "Angleze": "en",
    "Serbe": "sr",
    "Maqedonase": "mk"
}

def adjust_question_code(text, from_lang, to_lang):
    match = re.match(r'^(Q\d+[a-zA-Z]?|P\d+[a-zA-Z]?)(.*)', str(text))
    if match:
        code = match.group(1)
        rest = match.group(2)
        if from_lang == "en" and to_lang in ["sq", "sr", "mk"]:
            code = code.replace("Q", "P")
        elif from_lang == "sq" and to_lang == "en":
            code = code.replace("P", "Q")
        elif from_lang == "en" and to_lang == "sr":
            code = code.replace("Q", "P")
        return code, rest
    else:
        return '', text

def translate_text(text, from_lang, to_lang):
    if pd.isna(text) or not str(text).strip():
        return text

    code, remaining_text = adjust_question_code(text, from_lang, to_lang)

    path = "/translate?api-version=3.0"
    params = f"&from={from_lang}&to={to_lang}"
    url = AZURE_TRANSLATOR_ENDPOINT + path + params

    headers = {
        'Ocp-Apim-Subscription-Key': AZURE_TRANSLATOR_KEY,
        'Ocp-Apim-Subscription-Region': AZURE_REGION,
        'Content-type': 'application/json'
    }

    body = [{"text": str(remaining_text)}]
    response = requests.post(url, headers=headers, json=body)

    if response.status_code != 200:
        print(f"Error: {response.status_code}, {response.text}")
        return text

    result = response.json()

    try:
        translated_text = result[0]["translations"][0]["text"]
        return code + translated_text
    except (KeyError, IndexError, TypeError) as e:
        print(f"Translation failed for: {text} → Response: {result}")
        return text

def translate_dataframe(df, source_col, target_col, from_lang, to_lang):
    df[target_col] = df[source_col].apply(lambda x: translate_text(x, from_lang, to_lang))
    return df

st.title("Aplikacion për Përkthimin e Pyetësorëve")

uploaded_file = st.file_uploader("Ngarko skedarin Excel", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    all_sheets = {sheet: pd.read_excel(uploaded_file, sheet_name=sheet) for sheet in sheet_names}

    if "translated_sheets" not in st.session_state:
        st.session_state.translated_sheets = {sheet: df.copy() for sheet, df in all_sheets.items()}
    if "used_sheets" not in st.session_state:
        st.session_state.used_sheets = []
    if "selected_sheet" not in st.session_state:
        st.session_state.selected_sheet = sheet_names[0]

    available_sheets = [sheet for sheet in sheet_names if sheet not in st.session_state.used_sheets]
    if not available_sheets:
        st.warning("Nuk ka faqe (sheet) të tjera për përkthim.")
    else:
        selected_sheet = st.selectbox("Zgjidh një faqe (sheet) për përkthim", available_sheets, index=available_sheets.index(st.session_state.selected_sheet) if st.session_state.selected_sheet in available_sheets else 0, key=f"sheet_select_{len(st.session_state.used_sheets)}")
        st.session_state.selected_sheet = selected_sheet
        df = st.session_state.translated_sheets[selected_sheet]
        st.write("Pamje paraprake e faqes së zgjedhur:", df.head())

        columns = df.columns.tolist()
        source_col = st.selectbox("Zgjidh kolonën që përmban tekstin për përkthim", columns, key=f"{selected_sheet}_source")
        from_lang_label = st.selectbox("Zgjidh gjuhën origjinale", list(LANGUAGE_OPTIONS_UI.keys()), key=f"{selected_sheet}_from")
        from_lang = LANGUAGE_OPTIONS_UI[from_lang_label]
        multiple_targets = st.multiselect("Zgjidh kolonat ku dëshiron të vendoset përkthimin (mund të zgjedhësh më shumë se një)", columns, key=f"{selected_sheet}_multitarget")
        target_languages = []
        for target_col in multiple_targets:
            lang_label = st.selectbox(f"Zgjidh gjuhën për {target_col}", list(LANGUAGE_OPTIONS_UI.keys()), key=f"{selected_sheet}_{target_col}_lang")
            target_languages.append((target_col, LANGUAGE_OPTIONS_UI[lang_label]))

        if st.button(f"Fillo Përkthimin për {selected_sheet}", key=f"start_translation_{selected_sheet}"):
            with st.spinner("Duke përkthyer... Ju lutemi prisni"):
                for target_col, to_lang in target_languages:
                    df = translate_dataframe(df, source_col, target_col, from_lang=from_lang, to_lang=to_lang)
            st.session_state.translated_sheets[selected_sheet] = df.copy()
            if selected_sheet not in st.session_state.used_sheets:
                st.session_state.used_sheets.append(selected_sheet)
            st.write("Pamje pas përkthimit:", df.head())

        if len([sheet for sheet in sheet_names if sheet not in st.session_state.used_sheets]) > 0:
            continue_translation = st.checkbox("Dëshiron të përkthesh një faqe tjetër?", key=f"continue_checkbox_{selected_sheet}")
            if not continue_translation:
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    for sheet in sheet_names:
                        st.session_state.translated_sheets.get(sheet, all_sheets[sheet]).to_excel(writer, sheet_name=sheet, index=False)

                st.download_button(
                    label="Shkarko Excel-in e Përditësuar me të gjitha përkthimet",
                    data=output.getvalue(),
                    file_name=uploaded_file.name
                )
        else:
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for sheet in sheet_names:
                    st.session_state.translated_sheets.get(sheet, all_sheets[sheet]).to_excel(writer, sheet_name=sheet, index=False)

            st.download_button(
                label="Shkarko Excel-in e Përditësuar me të gjitha përkthimet",
                data=output.getvalue(),
                file_name=uploaded_file.name
            )
