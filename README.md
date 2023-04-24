# Synonym Finder

This Python-based project allows you to search for synonyms of a given word in Microsoft Word documents within a specified directory. The project uses [spaCy](https://spacy.io/) for natural language processing to find synonyms based on similarity scores.

## Features

- Crawl a given directory and its subdirectories
- Open and read Word files (.docx)
- Detect synonyms of a given input word based on similarity scores
- Output the found synonym, line number, paragraph number, and full path of the Word file

## Getting Started

### Prerequisites

- Python 3.x
- pip
- virtualenv (optional but recommended)

### Installation

1. Clone the repository:

git clone https://github.com/fajim1/synonym_finder.git
cd synonym-finder


2. (Optional) Create and activate a virtual environment:

python3 -m venv venv
source venv/bin/activate # For Unix-based systems
venv\Scripts\activate # For Windows systems


3. Install the required packages:

pip install -r requirements.txt
python -m spacy download en_core_web_sm


### Usage

1. Run the script:
python main.py


2. Enter the target directory, which can be any directory that contains Word files.
3. Enter the word you want to find synonyms for.
4. The script will display the found synonyms, line number, paragraph number, and the absolute path of the Word file containing the synonym.

## License

This project is licensed under the [MIT License](LICENSE).


