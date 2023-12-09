import docx
import os
from underthesea import sent_tokenize
import shutil
import re


def read_doc(file_path):
    """
    Read a Word document and return a list of tokenized sentences.
    """
    document = docx.Document(file_path)
    sentences = [sentence for paragraph in document.paragraphs for sentence in sent_tokenize(paragraph.text)]
    sentences = [s.lower() for s in sentences if s != ""]
    return sentences

def extract(input_files, output_dir, terms_group):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    for input_file in input_files:
        sentences = read_doc(input_file)

        for words in terms_group:
            filtered_texts = []
            for sentence in sentences:
                if len(words) > 1:
                    # Check if all words are in the sentence in the correct order
                    matched = True
                    last_found = -1
                    for word in words:
                        match = re.search(r'\b{}\b'.format(re.escape(word)), sentence[last_found+1:])
                        if match:
                            last_found += match.start() + 1
                        else:
                            matched = False
                            break
                    if matched:
                        filtered_texts.append(sentence)
                elif len(words) == 1:
                    # Check if the single word is in the sentence
                    if re.search(r'\b{}\b'.format(re.escape(words[0])), sentence):
                        filtered_texts.append(sentence)

            if filtered_texts:
                filename = "-".join(words)
                subfolder = os.path.join(output_dir, filename)
                if not os.path.exists(subfolder):
                    os.makedirs(subfolder)

                file_name = os.path.basename(input_file)
                output_file = os.path.join(subfolder, file_name + '.txt')

                with open(output_file, 'w') as f:
                    for text in filtered_texts:
                        f.write(text + '\n\n')

input_dir = "inputs"
output_dir = "outputs"

# Remove the outputs folder if it exists, create a new empty outputs folder
if os.path.exists(output_dir) and os.path.isdir(output_dir):
    shutil.rmtree(output_dir)
os.makedirs(output_dir)

input_files = [os.path.join(input_dir, file) for file in os.listdir(input_dir) if file.endswith(".docx")]

with open("query.txt", "r") as f:
    terms_group = [line.split(",") for line in f.read().splitlines() if line != ""]

extract(input_files, output_dir, terms_group)