import os
from pathlib import Path
from docx import Document
from word_handler import find_synonyms_in_paragraphs,read_paragraphs_from_docx

def main():
    target_directory = input("Enter the target directory: ")
    search_word = input("Enter the word you want to find synonyms for: ")

    for root, _, files in os.walk(target_directory):
        for file in files:
            if file.endswith(".docx"):
                file_path = Path(root) / file
                paragraphs = read_paragraphs_from_docx(file_path)

                synonyms = find_synonyms_in_paragraphs(paragraphs, search_word)
                if synonyms:
                    print(f"\nSynonyms found in file: {file_path.absolute()}")
                    for synonym in synonyms:
                        print(f"  - Word: {synonym['word']}, "
                              f"Paragraph number: {synonym['paragraph_number']}, "
                              f"Line number: {synonym['line_number']}",
                              f"Similarity score: {synonym['similarity_score']}")

if __name__ == "__main__":
    main()
