# Import the necessary libraries and modules
import os
from pathlib import Path
from docx import Document
from word_handler import find_synonyms_in_paragraphs, read_paragraphs_from_docx

# Define a function named "main"
def main():
    # Prompt the user to enter a target directory and a search word
    target_directory = input("Enter the target directory: ")
    search_word = input("Enter the word you want to find synonyms for: ")

    # Use os.walk to iterate over all files in the directory tree
    for root, _, files in os.walk(target_directory):
        # Loop over all files in the current directory
        for file in files:
            # Check if the file is a Word document
            if file.endswith(".docx"):
                # Construct the full path to the file
                file_path = Path(root) / file
                # Read the paragraphs from the file
                paragraphs = read_paragraphs_from_docx(file_path)

                # Find synonyms of the search word in the paragraphs
                synonyms = find_synonyms_in_paragraphs(paragraphs, search_word)
                # If synonyms are found, print them to the console
                if synonyms:
                    print(f"\nSynonyms found in file: {file_path.absolute()}")
                    for synonym in synonyms:
                        print(f"  - Word: {synonym['word']}, "
                              f"Paragraph number: {synonym['paragraph_number']}, "
                              f"Line number: {synonym['line_number']}",
                              f"Similarity score: {synonym['similarity_score']}")

# Check if the script is being run directly and call the "main" function if it is
if __name__ == "__main__":
    main()
