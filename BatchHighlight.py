import os
import csv
import json
import win32com.client as win32
from tkinter import Tk, filedialog, Button, Label, StringVar, Entry, Frame, Checkbutton, BooleanVar
import fitz  # PyMuPDF
from collections import defaultdict

class Application(Tk):
    def __init__(self):
        Tk.__init__(self)
        self.title("Find bad words")
        self.word = win32.gencache.EnsureDispatch('Word.Application')

        self.input_dir = StringVar()
        self.output_dir = StringVar()
        self.bad_words_entries = []
        self.match_all_word_forms = BooleanVar()  # Checkbox variable for "Match all word forms"
        self.trim_spaces = BooleanVar()

        default_bad_words = [
            "ensure", "assure", "insure",
            "warrant", "warranty", "warrantee",
            "guarant", "guarantee",
            "certify",
            "expert", "expertise",
            "best",
            "highest"
        ]

        # Create a main frame with padding
        self.main_frame = Frame(self, padx=40, pady=40)
        self.main_frame.pack(fill='both', expand=True)

        try:
            with open("config.json", "r") as f:
                data = json.load(f)
            self.input_dir.set(data.get("input_dir", ""))
            self.output_dir.set(data.get("output_dir", ""))
            loaded_bad_words = data.get("bad_words", default_bad_words)
            self.match_all_word_forms.set(data.get("match_all_word_forms", True))  # Load checkbox state
            self.trim_spaces.set(data.get("trim_spaces", True))  # Load checkbox state
        except FileNotFoundError:
            self.match_all_word_forms.set(True)  # Enable checkbox by default
            self.trim_spaces.set(True)  # Enable checkbox by default
            loaded_bad_words = default_bad_words
            self.save_json()

        Button(self.main_frame, text="Select Input Folder", command=self.select_input_folder).pack()
        Label(self.main_frame, textvariable=self.input_dir).pack()

        Button(self.main_frame, text="Select Output Folder", command=self.select_output_folder).pack()
        Label(self.main_frame, textvariable=self.output_dir).pack()

        Label(self.main_frame, text="Bad Words:").pack()

        Checkbutton(self.main_frame, text="Match all word forms for Word documents", variable=self.match_all_word_forms).pack()

        Checkbutton(self.main_frame, text="Trim spaces before and after words below for PDF search", variable=self.trim_spaces).pack()

        self.words_frame = Frame(self.main_frame)
        self.words_frame.pack()

        for bad_word in loaded_bad_words:
            self.add_word_entry(bad_word)

        Button(self.main_frame, text="Add Word", command=self.add_word_entry).pack(padx=10, pady=10)
        Button(self.main_frame, text="Remove Word", command=self.remove_word_entry).pack()

        Label(self.main_frame, text="The search is case-insensitive and PDF files will be searched just on basic match \nso any word containing one of the above entries will be highlighted.").pack()

        Button(self.main_frame, text="Run", command=self.run).pack(padx=10, pady=10)

    def select_input_folder(self):
        self.input_dir.set(filedialog.askdirectory())
        self.save_json()

    def select_output_folder(self):
        self.output_dir.set(filedialog.askdirectory())
        self.save_json()

    def add_word_entry(self, word=None) -> None:
        new_bad_word = StringVar()
        new_bad_word.set(word or "")
        new_entry = Entry(self.words_frame, textvariable=new_bad_word)
        new_entry.pack()
        self.bad_words_entries.append(new_entry)
        self.save_json()

    def remove_word_entry(self):
        if self.bad_words_entries:
            entry_to_remove = self.bad_words_entries.pop()
            entry_to_remove.destroy()
        self.save_json()

    def run(self):
        self.word = win32.gencache.EnsureDispatch('Word.Application')
        result_file_path = os.path.normpath(os.path.join(self.output_dir.get(), "counts.csv"))

        overall_result = {}

        for file_name in os.listdir(self.input_dir.get()):
            full_path = os.path.normpath(os.path.join(self.input_dir.get(), file_name))
            word_counts_dict = defaultdict(int)  # Keep track of counts for each bad word

            if file_name.endswith(('.doc', '.docx', '.docm')):
                try:
                    doc = self.word.Documents.Open(full_path)
                    self.word.ActiveDocument.ActiveWindow.View.Type = 3
                    for word in self.bad_words:
                        word_count = self.highlight_word(doc, word)  # Get both count and details
                        word_counts_dict[word] += word_count
                    if sum(word_counts_dict.values()) > 0:  # If any bad words found
                        output_path = os.path.normpath(os.path.join(self.output_dir.get(), file_name))
                        doc.SaveAs(output_path)
                        overall_result[file_name] = {'Bad words count': sum(word_counts_dict.values()), 
                                                    'Details': dict(word_counts_dict)}
                    doc.Close()
                except Exception as e:
                    print(f"Error processing file {full_path}: {str(e)}")
            elif file_name.endswith('.pdf'):
                try:
                    doc = fitz.open(full_path)
                    word_count, word_detail = self.highlight_word_pdf(doc, self.bad_words) # Update here
                    if word_count > 0:
                        output_path = os.path.normpath(os.path.join(self.output_dir.get(), file_name))
                        doc.save(output_path)
                        overall_result[file_name] = {'Bad words count': word_count, 'Details': word_detail} # Update here
                except Exception as e:
                    print(f"Error processing file {full_path}: {str(e)}")

        # Save CSV file
        with open(result_file_path, "w", newline='') as f:
            writer = csv.writer(f)
            # Write the header row
            header_row = ['File Name'] + self.bad_words + ['Total Count']
            writer.writerow(header_row)
            # Write the data rows
            for file_name, data in overall_result.items():
                row = [file_name] + [data['Details'].get(word, 0) for word in self.bad_words] + [data['Bad words count']]
                writer.writerow(row)
        
        self.word.Quit()
        os.startfile(result_file_path)
        os.startfile(self.output_dir.get())  # open the output directory

        self.save_json()


    def highlight_word_pdf(self, doc, words):
        word_counts_dict = defaultdict(int)  # Keep track of counts for each bad word
        for word in words: # Loop over each word
            for page in doc:
                text_instances = page.search_for(word)
                for inst in text_instances:
                    page.add_highlight_annot(inst)
                    word_counts_dict[word] += 1
        return sum(word_counts_dict.values()), dict(word_counts_dict)

    def save_json(self):

        if self.trim_spaces.get():
            self.bad_words = [entry.get().strip() for entry in self.bad_words_entries]
        else:
            self.bad_words = [entry.get() for entry in self.bad_words_entries]

        data = {
            "input_dir": self.input_dir.get(),
            "output_dir": self.output_dir.get(),
            "bad_words": self.bad_words,
            "match_all_word_forms": self.match_all_word_forms.get(),  # Save checkbox state
            "trim_spaces": self.trim_spaces.get()
        }
        with open("config.json", "w") as f:
            json.dump(data, f)

    def highlight_word(self, doc, word: str):
        word = word.strip()
        word_count = 0
        rng = doc.Range()
        rng.Find.ClearFormatting()
        rng.Find.Replacement.ClearFormatting()
        rng.Find.Replacement.Highlight = True
        rng.Find.MatchCase = False  # Make it case-insensitive
        rng.Find.MatchAllWordForms = self.match_all_word_forms.get()
        
        found = rng.Find.Execute(
            FindText=word,
            MatchWholeWord=True,
            Forward=True,
            Wrap=win32.constants.wdFindStop,  # Stop at the end of the document
            Format=True,
            Replace=win32.constants.wdReplaceOne,  # Replace one instance at a time
        )
        
        while found:
            word_count += 1
            rng.Collapse(Direction=win32.constants.wdCollapseEnd)  # Collapse the range past the found word
            found = rng.Find.Execute(
                FindText=word,
                MatchWholeWord=True,
                Forward=True,
                Wrap=win32.constants.wdFindStop,
                Format=True,
                Replace=win32.constants.wdReplaceOne,
            )

        return word_count




if __name__ == "__main__":
    app = Application()
    app.mainloop()