import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
import plotly.io as pio
import tempfile
import io
import os
import json
from datetime import datetime
import cohere
import xlsxwriter

st.set_page_config(page_title="ExcelVA!: Smart Excel Analytics Suite", layout="wide")

# Sidebar App Icon + Username Input
st.sidebar.markdown("<img src='https://cdn-icons-png.flaticon.com/512/888/888879.png' width='30'>", unsafe_allow_html=True)
st.sidebar.markdown("### Enter your name to personalize your assistant")
username = st.sidebar.text_input("Your Name", value="Guest")

# Load previous chat history if it exists
history_file = f"chat_history_{username}.json"
if "chat_history" not in st.session_state:
    if os.path.exists(history_file):
        with open(history_file, "r") as f:
            st.session_state.chat_history = json.load(f)
    else:
        st.session_state.chat_history = []

# Placeholder for chatbot DataFrame results
if "chatbot_filtered_df" not in st.session_state:
    st.session_state.chatbot_filtered_df = None

# Delete chat history toggle
if st.sidebar.button("🗑️ Clear Chat History"):
    if os.path.exists(history_file):
        os.remove(history_file)
    st.session_state.chat_history = []

# Check for agent mode
if "agent_mode" not in st.session_state:
    st.session_state.agent_mode = False

# Detect commands for agent mode switching
if st.session_state.chat_history:
    last_user_msg = st.session_state.chat_history[-1][1]
    if last_user_msg.strip().lower() == "/ult":
        st.session_state.agent_mode = True
        st.session_state.chat_history.append(("bot", "Agent Mode activated. Let’s break down your stats like a pro. — ExcelVA! (Agent Mode)"))
    elif last_user_msg.strip().lower() == "/calm":
        st.session_state.agent_mode = False
        st.session_state.chat_history.append(("bot", "Returning to standard assistant mode. Breathe easy. — ExcelVA!"))

# Styling
header_style = """
    <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@700&display=swap" rel="stylesheet">
    <style>
    .header-container {
        display: flex;
        align-items: center;
        justify-content: center;
        margin-bottom: 0.5rem;
        margin-top: -2rem;
    }
    .header-text {
        font-family: 'Montserrat', sans-serif;
        color: #000000;
        font-weight: bold;
        font-size: 48px;
        margin-left: 15px;
    }
    .sub-text {
        font-family: 'Montserrat', sans-serif;
        text-align: left;
        margin-left: calc(50% - 80px);
        color: #555555;
        font-size: 18px;
        margin-top: -8px;
    }
    .stChatMessageContent, .stChatMessageContent * {
        color: #222222 !important;
    }
    </style>
"""

agent_style = """
    <style>
    .header-text {
        color: #00FFFF !important;
        text-shadow: 0 0 8px #00FFFF;
    }
    .stChatMessageContent, .stChatMessageContent * {
        color: #E0FFFF !important;
    }
    .sub-text {
        color: #00CED1 !important;
    }
    </style>
""" if st.session_state.agent_mode else ""

st.markdown(header_style + agent_style, unsafe_allow_html=True)

# Header and Subheader
st.markdown("""
    <div class="header-container">
        <img src="https://cdn-icons-png.flaticon.com/512/888/888879.png" width="50"/>
        <div class="header-text">ExcelVA!</div>
    </div>
    <div class='sub-text'>Smart Excel Analytics Suite</div>
""", unsafe_allow_html=True)

# File uploader
uploaded_file = st.sidebar.file_uploader("Upload an Excel or CSV file", type=["xlsx", "csv"])

if uploaded_file:
    file_extension = os.path.splitext(uploaded_file.name)[1]
    if file_extension == ".csv":
        df = pd.read_csv(uploaded_file)
    elif file_extension == ".xlsx":
        df = pd.read_excel(uploaded_file)
    else:
        st.error("Unsupported file format!")
        st.stop()

    st.success(f"File uploaded: **{uploaded_file.name}**")

    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    categorical_cols = df.select_dtypes(exclude='number').columns.tolist()

    date_cols = [col for col in df.columns if pd.api.types.is_datetime64_any_dtype(df[col]) or 'date' in col.lower()]
    if date_cols:
        date_col = date_cols[0]
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        df.dropna(subset=[date_col], inplace=True)
        min_date = df[date_col].min()
        max_date = df[date_col].max()
        start_date, end_date = st.sidebar.date_input("Filter by Date Range", [min_date, max_date])
        df = df[(df[date_col] >= pd.to_datetime(start_date)) & (df[date_col] <= pd.to_datetime(end_date))]

    st.sidebar.markdown("---")
    st.sidebar.markdown("**General Filters** (Optional)")

    if numeric_cols:
        for col in numeric_cols:
            if st.sidebar.checkbox(f"Enable filter for: {col}"):
                min_val, max_val = int(df[col].min()), int(df[col].max())
                selected_range = st.sidebar.slider(f"{col} Range", min_val, max_val, (min_val, max_val))
                df = df[(df[col] >= selected_range[0]) & (df[col] <= selected_range[1])]

    if categorical_cols:
        for col in categorical_cols:
            unique_vals = df[col].dropna().unique()
            if len(unique_vals) < 50:
                if st.sidebar.checkbox(f"Enable filter for: {col}"):
                    selected_vals = st.sidebar.multiselect(f"Select {col}", unique_vals.tolist(), default=unique_vals.tolist())
                    df = df[df[col].isin(selected_vals)]

    st.subheader("Data Preview")
    st.dataframe(df)

    # --- Excel Formula Assistant ---
    st.markdown("### 🔍 Excel Formula Assistant")
    mode = st.radio("What do you want help with?", ["Explain a formula", "Generate a formula"])
    user_question = st.text_input("Type your question or formula")

    if st.button("Ask ExcelVA!"):
        with st.spinner("Thinking..."):
            try:
                co = cohere.Client(st.secrets["COHERE_API_KEY"])
                prompt = f"Explain what this Excel formula does in simple English: {user_question}" if mode == "Explain a formula" \
                         else f"Generate an Excel formula for the following request: {user_question}"
                response = co.generate(
                    model="command",
                    prompt=prompt,
                    max_tokens=250,
                    temperature=0.6
                )
                st.success("Here's what ExcelVA! says:")
                st.write(response.generations[0].text.strip())
            except Exception as e:
                st.error(f"Something went wrong: {e}")

    # --- Charts and Stats ---
    if numeric_cols:
        selected_numeric = st.selectbox("Select numeric column for analysis", numeric_cols)
        st.subheader(f"Summary Statistics for {selected_numeric}")
        st.write(df[selected_numeric].describe())

        st.subheader(f"Histogram of {selected_numeric}")
        st.plotly_chart(px.histogram(df, x=selected_numeric, nbins=30, title=f"Histogram of {selected_numeric}"))

        if categorical_cols:
            selected_cat = st.selectbox("Group by categorical column", categorical_cols)
            st.subheader(f"Bar Chart: Average {selected_numeric} by {selected_cat}")
            st.plotly_chart(px.bar(df.groupby(selected_cat)[selected_numeric].mean().reset_index(), x=selected_cat, y=selected_numeric))

            st.subheader(f"Pie Chart: Distribution by {selected_cat}")
            st.plotly_chart(px.pie(df, names=selected_cat, title=f"Distribution by {selected_cat}"))

        if date_cols:
            st.subheader(f"{selected_numeric} Trend Over Time")
            trend = df.groupby(date_col)[selected_numeric].sum().reset_index()
            st.plotly_chart(px.line(trend, x=date_col, y=selected_numeric))

    # --- Bot Greeting ---
    st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px; margin-top: 30px;'>
        <img src='https://cdn-icons-png.flaticon.com/512/888/888879.png' width='32'>
        <span style='font-family: Montserrat, sans-serif; font-size: 20px; color: #333;'>Hey there! I am <b>ExcelVA!</b>, here to help you with your data 🔍</span>
    </div>
    """, unsafe_allow_html=True)

    # --- AI Chat Assistant ---
    user_query = st.chat_input("Ask me anything about your data...")

    if user_query:
        st.session_state.chat_history.append(("user", user_query))
        try:
            co = cohere.Client(st.secrets["COHERE_API_KEY"])
            csv_sample = df.head(50).to_csv(index=False)
            prompt = f"""
You are ExcelVA!, a smart Excel analytics AI. Based on the following dataset, answer user queries clearly and end each response with '— ExcelVA!'.

Data:
{csv_sample}

User question: {user_query}
"""
            response = co.generate(
                model='command',
                prompt=prompt,
                max_tokens=300,
                temperature=0.6
            )
            bot_reply = response.generations[0].text.strip()

            # If data requested, simulate filtered data example
            if "data" in user_query.lower():
                st.session_state.chatbot_filtered_df = df[df.select_dtypes(include='number').columns[0] > 2].head(10)
                bot_reply += "\n\nYou can download the filtered data below — ExcelVA!"

            st.session_state.chat_history.append(("bot", bot_reply))
        except Exception as e:
            st.session_state.chat_history.append(("bot", f"Failed to generate insight: {e}"))

        with open(history_file, "w") as f:
            json.dump(st.session_state.chat_history, f)

    for role, msg in st.session_state.chat_history:
        with st.chat_message(role):
            st.markdown(msg)

    if st.session_state.chatbot_filtered_df is not None:
        download_buffer = io.BytesIO()
        with pd.ExcelWriter(download_buffer, engine='xlsxwriter') as writer:
            st.session_state.chatbot_filtered_df.to_excel(writer, index=False)
            writer.close()
            data_excel = download_buffer.getvalue()

        st.download_button("📥 Download Chatbot Filtered Data", data=data_excel, file_name="chatbot_filtered_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.warning("Upload a file to begin.")
