# Import the necessary libraries and modules
import os
import spacy
from docx import Document
import re

# Define a function named "read_paragraphs_from_docx"
def read_paragraphs_from_docx(docx_file_path):
    # Load the Word document using the "docx" library
    doc = Document(docx_file_path)
    paragraphs = []

    # Loop over all paragraphs in the document
    for paragraph in doc.paragraphs:
        # Replace newline characters with space
        cleaned_text = paragraph.text.replace('\n', ' ').strip()
        # If the paragraph has text, add it to the list of paragraphs
        if cleaned_text:
            paragraphs.append(cleaned_text)

    return paragraphs

# Load the medium model of the "spacy" library to access word vectors for similarity calculation
nlp = spacy.load("en_core_web_md")

# Define a function named "find_synonyms_in_paragraphs"
def find_synonyms_in_paragraphs(paragraphs, search_word, SIMILARITY_THRESHOLD=0.5):
    # Process the search word using the "spacy" model
    search_tokens = nlp(search_word)
    synonyms = []

    # Loop over all paragraphs and lines in the list of paragraphs
    for paragraph_number, paragraph in enumerate(paragraphs, start=1):
        # Split the paragraph into sentences using regex
        lines = re.split(r'[.!?] *', paragraph)
        for line_number, line in enumerate(lines, start=1):
            # Process the line using the "spacy" model
            tokens = nlp(line)
            # Loop over all tokens in the line
            for token in tokens:
                # Check if the token is a word and is different from the search word
                if token.is_alpha and token.lower_ != search_word.lower():
                    # Calculate the similarity score between the search word and the token
                    similarity = search_tokens.similarity(token)
                    # If the similarity score is above the threshold, add the token to the list of synonyms
                    if similarity > SIMILARITY_THRESHOLD:
                        synonyms.append({
                            'word': token.text,
                            'line_number': line_number,
                            'paragraph_number': paragraph_number,
                            'similarity_score': similarity,
                        })

    return synonyms
