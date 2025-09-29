import streamlit as st
import pandas as pd
import re
import string
import nltk
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer
from nltk.sentiment.vader import SentimentIntensityAnalyzer
import os
import torch
import torch.nn.functional as F
from transformers import AutoTokenizer, AutoModelForSequenceClassification
import time
import sys # You'll need to import the sys module
    
# --- Streamlit App Configuration ---
# This must be the first Streamlit command in the entire script.
st.set_page_config(page_title="Multi-Language Sentiment Analyzer", layout="wide", page_icon="📈")

# --- Setup and NLTK Downloads ---
@st.cache_resource
def download_nltk_data():
    """Downloads necessary NLTK data and caches it."""
    nltk.download('punkt')
    nltk.download('stopwords')
    nltk.download('wordnet')
    nltk.download('omw-1.4')
    nltk.download('vader_lexicon')
    return True

if not download_nltk_data():
    st.error("Failed to download NLTK data. Please check your internet connection.")
    st.stop()



# --- Model Loading ---
@st.cache_resource
def load_sinhala_model():
    """Loads the Sinhala sentiment analysis model and tokenizer."""
    model_name = "sinhala-nlp/sinhala-sentiment-analysis-sinbert-small"
    tokenizer = AutoTokenizer.from_pretrained(model_name)
    model = AutoModelForSequenceClassification.from_pretrained(model_name)
    return tokenizer, model

tokenizer, model = load_sinhala_model()

# --- Constants ---
stop_words = set(stopwords.words("english"))
lemmatizer = WordNetLemmatizer()
sia = SentimentIntensityAnalyzer()
label_map = {0: "Neutral", 1: "Positive", 2: "Negative"}

# --- Helper Functions ---
def detect_language(text):
    """Detects if a string contains Sinhala characters."""
    sinhala_unicode_range = any('඀' <= c <= '෿' for c in str(text))
    return 'si' if sinhala_unicode_range else 'en'

def clean_text(text):
    """Cleans English text for sentiment analysis."""
    try:
        text = str(text).lower()
        text = re.sub(r'http\S+|www\S+', '', text)
        text = re.sub(r'<.*?>', '', text)
        text = text.translate(str.maketrans('', '', string.punctuation))
        tokens = nltk.word_tokenize(text)
        tokens = [lemmatizer.lemmatize(w) for w in tokens if w.isalpha() and w not in stop_words]
        cleaned = " ".join(tokens)
        return cleaned if cleaned.strip() else text
    except Exception:
        return str(text)

def get_english_sentiment(text):
    """Analyzes sentiment of English text using VADER."""
    cleaned_text = clean_text(text)
    score = sia.polarity_scores(cleaned_text)['compound']
    if score >= 0.05:
        return 'Positive'
    elif score <= -0.05:
        return 'Negative'
    else:
        return 'Neutral'

def predict_sinhala_sentiment(text):
    """Analyzes sentiment of Sinhala text using a pre-trained model."""
    try:
        inputs = tokenizer(text, return_tensors="pt", truncation=True, max_length=512)
        with torch.no_grad():
            logits = model(**inputs).logits
        probs = F.softmax(logits, dim=1)[0]
        idx = torch.argmax(probs).item()
        return label_map[idx]
    except Exception:
        return "Neutral"

def get_final_sentiment(text):
    """Combines English and Sinhala sentiment analysis."""
    lang = detect_language(text)
    if lang == 'en':
        return get_english_sentiment(text)
    elif lang == 'si':
        return predict_sinhala_sentiment(text)
    else:
        return 'Neutral'

# Sidebar
with st.sidebar:
    st.header("⚙️ Settings")
    st.write("Upload a CSV file and analyze sentiment.")
    st.info("💡 **Tip:** The app analyzes both English and Sinhala text.")
    st.markdown("---")
    st.header("About")
    st.write("This application uses a combination of **VADER** for English sentiment and a **Hugging Face model** for Sinhala sentiment analysis.")
    st.write("Created by @djslash9.")
    st.markdown("---")
    # --- New Exit Button ---
    if st.button("🚪 Exit App"):
        st.stop()   
    
st.title("📊 Sprout Social Sentiment Analyzer")
st.markdown("Upload a CSV file to perform sentiment analysis on its contents. The app supports both **English** and **Sinhala** languages.")

# File Uploader
uploaded_file = st.file_uploader("📂 Choose a CSV file", type="csv")

if uploaded_file:
    # Store the original file name as soon as it's uploaded
    original_file_name = uploaded_file.name
    
    try:
        df = pd.read_csv(uploaded_file)
        st.success("🎉 File uploaded successfully!")
        
        # Display file contents
        st.subheader("📝 File Contents")
        st.dataframe(df.head())
        
        # Select columns to remove
        columns_to_remove = st.multiselect(
            "🗑️ Select columns to remove",
            options=df.columns
        )
        
        if columns_to_remove:
            df = df.drop(columns=columns_to_remove)
            st.info(f"Columns {', '.join(columns_to_remove)} have been removed.")
            st.dataframe(df.head())

        # Select column to analyze
        columns_to_analyze = [col for col in df.columns if df[col].dtype in ['object', 'string']]
        if not columns_to_analyze:
            st.warning("No text-based columns found to analyze. Please upload a file with string data.")
            st.stop()
            
        selected_column = st.selectbox(
            "✍️ Select the column to analyze for sentiment",
            options=columns_to_analyze
        )
        
        # Analysis button
        if st.button("🚀 Analyze Sentiment"):
            if selected_column:
                with st.spinner("Analyzing sentiments... This may take a moment."):
                    progress_bar = st.progress(0)
                    total_rows = len(df)
                    
                    df["Sentiment"] = None
                    
                    start_time = time.time()
                    
                    for i, row in df.iterrows():
                        text = row[selected_column]
                        sentiment = get_final_sentiment(text)
                        df.at[i, "Sentiment"] = sentiment
                        
                        # Update progress bar
                        progress = (i + 1) / total_rows
                        progress_bar.progress(progress)
                        
                    progress_bar.empty()
                    end_time = time.time()
                    duration = end_time - start_time
                    st.success(f"✅ Analysis complete! (Took {duration:.2f} seconds)")

                # Show analyzed file
                st.subheader("✅ Analyzed Data")
                st.dataframe(df)

                # Show sentiment distribution
                st.subheader("📈 Sentiment Distribution")
                sentiment_counts = df['Sentiment'].value_counts()
                st.bar_chart(sentiment_counts)
                
                # Download button
                @st.cache_data
                def convert_df_to_csv(df):
                    return df.to_csv(index=False).encode('utf-8')

                csv_data = convert_df_to_csv(df)
                
                st.download_button(
                    label="⬇️ Download Analyzed CSV",
                    data=csv_data,
                    # Use the stored original_file_name here
                    file_name=f"analyzed_{original_file_name}", # Optionally, prepend "analyzed_"
                    mime="text/csv",
                )
    
    except Exception as e:
        st.error(f"An error occurred: {e}")
        st.warning("Please ensure the uploaded file is a valid CSV and contains text data.")
else:
    st.info("Please upload a CSV file to begin sentiment analysis.")