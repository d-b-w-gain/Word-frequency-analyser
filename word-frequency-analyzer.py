import docx
import matplotlib.pyplot as plt
from collections import Counter
import string
import sys
from pathlib import Path

def analyze_word_doc(file_path):
    """
    Analyze word frequencies in a Word document and create a bar chart.
    
    Args:
        file_path (str): Path to the Word document
    """
    # Read the document
    doc = docx.Document(file_path)
    
    # Extract text and combine all paragraphs
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text.lower())
    
    # Join all text and split into words
    words = ' '.join(full_text).split()
    
    # Clean words (remove punctuation and numbers)
    cleaned_words = []
    for word in words:
        word = word.strip(string.punctuation)
        if word and not word.isdigit() and len(word) > 2:  # Skip small words and numbers
            cleaned_words.append(word)
    
    # Count word frequencies
    word_freq = Counter(cleaned_words)
    
    # Get top 20 words
    most_common = word_freq.most_common(20)
    words, frequencies = zip(*most_common)
    
    # Create bar chart
    plt.figure(figsize=(15, 8))
    plt.bar(words, frequencies)
    plt.xticks(rotation=45, ha='right')
    plt.title('Most Common Words in Your Document')
    plt.xlabel('Words')
    plt.ylabel('Frequency')
    
    # Adjust layout to prevent label cutoff
    plt.tight_layout()
    
    # Save the plot
    output_path = Path(file_path).parent / 'word_frequency_chart.png'
    plt.savefig(output_path)
    print(f"Chart saved as: {output_path}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script.py path/to/document.docx")
    else:
        analyze_word_doc(sys.argv[1])