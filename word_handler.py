import os
import spacy
from docx import Document
import re


def read_paragraphs_from_docx(docx_file_path):
    doc = Document(docx_file_path)
    paragraphs = []

    for paragraph in doc.paragraphs:
        # Replace newline characters with space
        cleaned_text = paragraph.text.replace('\n', ' ').strip()
        if cleaned_text:
            paragraphs.append(cleaned_text)

    return paragraphs


nlp = spacy.load("en_core_web_md")  # Use the medium model to have access to word vectors for similarity calculation



def find_synonyms_in_paragraphs(paragraphs, search_word,SIMILARITY_THRESHOLD=0.5):
    search_tokens = nlp(search_word)
    synonyms = []

    for paragraph_number, paragraph in enumerate(paragraphs, start=1):
        # Split paragraph into sentences using regex
        lines = re.split(r'[.!?] *', paragraph)
        for line_number, line in enumerate(lines, start=1):
            tokens = nlp(line)
            for token in tokens:
                if token.is_alpha and token.lower_ != search_word.lower():
                    similarity = search_tokens.similarity(token)
                    if similarity > SIMILARITY_THRESHOLD:
                        synonyms.append({
                            'word': token.text,
                            'line_number': line_number,
                            'paragraph_number': paragraph_number,
                            'similarity_score': similarity,
                        })

    return synonyms

