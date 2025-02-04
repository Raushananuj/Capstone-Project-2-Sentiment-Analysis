import os
import pandas as pd
import requests
from bs4 import BeautifulSoup
from textblob import TextBlob
from collections import Counter
import re
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize

# Download necessary NLTK resources
nltk.download('punkt')
nltk.download('stopwords')

# Define stopwords for analysis
stop_words = set(stopwords.words('english'))

def fetch_article_content(url):
    """
    Extract the title and main text from the article URL.
    """
    try:
        response = requests.get(url)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')

        # Extract title and article body
        title = soup.find('h1').get_text(strip=True) if soup.find('h1') else "No Title"
        paragraphs = soup.find_all('p')
        article_text = ' '.join([p.get_text(strip=True) for p in paragraphs])
        return title, article_text
    except Exception as e:
        print(f"Error fetching {url}: {e}")
        return None, None

def save_text_file(url_id, title, text):
    """
    Save the extracted article to a text file.
    """
    try:
        with open(f"{url_id}.txt", 'w', encoding='utf-8') as file:
            file.write(f"Title: {title}\n\n{text}")
    except Exception as e:
        print(f"Error saving file for URL_ID {url_id}: {e}")

def analyze_text(text):
    """
    Perform text analysis and compute metrics.
    """
    sentences = sent_tokenize(text)
    words = word_tokenize(text)
    words_cleaned = [word.lower() for word in words if word.isalpha()]
    word_count = len(words_cleaned)
    
    # Positive and Negative word lists
    positive_words = ['good', 'great', 'positive', 'excellent', 'fortunate', 'correct', 'superior']
    negative_words = ['bad', 'poor', 'negative', 'wrong', 'inferior', 'unfortunate']
    
    positive_score = sum(word in positive_words for word in words_cleaned)
    negative_score = sum(word in negative_words for word in words_cleaned)
    
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 1e-5)
    subjectivity_score = (positive_score + negative_score) / (word_count + 1e-5)
    
    avg_sentence_length = word_count / len(sentences) if sentences else 0
    syllables = sum(len(re.findall(r'[aeiouy]', word)) for word in words_cleaned)
    syllable_per_word = syllables / word_count if word_count else 0
    
    # Fog Index and Complex Words
    complex_word_count = sum(len(re.findall(r'[aeiouy]', word)) > 2 for word in words_cleaned)
    percentage_complex_words = (complex_word_count / word_count) * 100 if word_count else 0
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
    
    avg_word_length = sum(len(word) for word in words_cleaned) / word_count if word_count else 0
    personal_pronouns = sum(word.lower() in ['i', 'we', 'me', 'us', 'my', 'our'] for word in words_cleaned)
    
    return {
        "Positive_score": positive_score,
        "Negative_score": negative_score,
        "Polarity_score": polarity_score,
        "Subjectivity_score": subjectivity_score,
        "Avg_sentence_length": avg_sentence_length,
        "Percentage_complex_words": percentage_complex_words,
        "Fog_index": fog_index,
        "complex_word_count": complex_word_count,
        "word_count": word_count,
        "Syllable_per_word": syllable_per_word,
        "Personal_pronouns": personal_pronouns,
        "Avg_word_length": avg_word_length,
    }

def main():
    # Input file and output directory
    input_file = "C:\\Users\\Dell\\Desktop\\Sentiment Analysis\\Input.xlsx"
    output_file = "Output.xlsx"
    output_directory = "articles"
    
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    
    os.chdir(output_directory)

    try:
        df = pd.read_excel(input_file)
        results = []
        
        for _, row in df.iterrows():
            url_id = row['URL_ID']
            url = row['URL']
            print(f"Processing {url_id}: {url}")
            
            title, article_text = fetch_article_content(url)
            if title and article_text:
                save_text_file(url_id, title, article_text)
                analysis = analyze_text(article_text)
                analysis.update({"URL_ID": url_id, "Title": title})
                results.append(analysis)
            else:
                print(f"Failed to process {url_id}")
        
        # Save results to Excel
        output_df = pd.DataFrame(results)
        output_df.to_excel(output_file, index=False)
        print(f"Analysis saved to {output_file}")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
