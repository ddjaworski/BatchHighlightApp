import os
import csv
import json
import win32com.client as win32
from tkinter import Tk, filedialog, Button, Label, StringVar, Entry, Frame, Checkbutton, BooleanVar

class Application(Tk):
    def __init__(self):
        Tk.__init__(self)
        self.title("Find bad words")
        self.word = win32.gencache.EnsureDispatch('Word.Application')

        self.input_dir = StringVar()
        self.output_dir = StringVar()
        self.bad_words_entries = []
        self.match_all_word_forms = BooleanVar()  # Checkbox variable for "Match all word forms"

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
        except FileNotFoundError:
            with open("config.json", "w") as f:
                json.dump({"bad_words": default_bad_words, "match_all_word_forms": True}, f)  # Save default checkbox state
            loaded_bad_words = default_bad_words

        

        Button(self.main_frame, text="Select Input Folder", command=self.select_input_folder).pack()
        Label(self.main_frame, textvariable=self.input_dir).pack()

        Button(self.main_frame, text="Select Output Folder", command=self.select_output_folder).pack()
        Label(self.main_frame, textvariable=self.output_dir).pack()

        Label(self.main_frame, text="Bad Words (search is case-insensitive):").pack()

        Checkbutton(self.main_frame, text="Match all word forms", variable=self.match_all_word_forms).pack()

        self.words_frame = Frame(self.main_frame)
        self.words_frame.pack()

        for bad_word in loaded_bad_words:
            self.add_word_entry(bad_word)

        Button(self.main_frame, text="Add Word", command=self.add_word_entry).pack(padx=10, pady=10)
        Button(self.main_frame, text="Remove Word", command=self.remove_word_entry).pack()

        Button(self.main_frame, text="Run", command=self.run).pack(padx=10, pady=10)

    def select_input_folder(self):
        self.input_dir.set(filedialog.askdirectory())
        self.save_json()

    def select_output_folder(self):
        self.output_dir.set(filedialog.askdirectory())
        self.save_json()
        

    def add_word_entry(self, word=None):
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
        result_file_path = os.path.normpath(os.path.join(self.output_dir.get(), "result.csv"))
        result_file = open(result_file_path, 'w', newline='')
        csv_writer = csv.writer(result_file)
        csv_writer.writerow(["File", "Bad words count"])

        for file_name in os.listdir(self.input_dir.get()):
            if file_name.endswith(('.doc', '.docx', '.docm')):
                full_path = os.path.normpath(os.path.join(self.input_dir.get(), file_name))
                try:
                    doc = self.word.Documents.Open(full_path)
                    word_count = 0
                    for word in self.bad_words:
                        count = self.highlight_word(doc, word)
                        word_count += count
                    if word_count > 0:
                        output_path = os.path.normpath(os.path.join(self.output_dir.get(), file_name))
                        doc.SaveAs(output_path)
                        csv_writer.writerow([file_name, word_count])
                    doc.Close()
                except Exception as e:
                    print(f"Error processing file {full_path}: {str(e)}")

        result_file.close()
        self.word.Quit()
        os.startfile(result_file_path)
        os.startfile(self.output_dir.get())  # open the output directory

        self.save_json()

    def save_json(self):
        self.bad_words = [entry.get().strip() for entry in self.bad_words_entries]

        data = {
            "input_dir": self.input_dir.get(),
            "output_dir": self.output_dir.get(),
            "bad_words": self.bad_words,
            "match_all_word_forms": self.match_all_word_forms.get(),  # Save checkbox state
        }
        with open("config.json", "w") as f:
            json.dump(data, f)

    def highlight_word(self, doc, word):
        word_count = 0
        rng = doc.Range()
        rng.Find.ClearFormatting()
        rng.Find.Replacement.ClearFormatting()
        rng.Find.Replacement.Highlight = True
        rng.Find.MatchCase = False  # Make it case-insensitive
        rng.Find.MatchAllWordForms = self.match_all_word_forms.get()

        rng.Find.Execute(
            FindText=word,
            MatchWholeWord=True,
            Forward=True,
            Wrap=win32.constants.wdFindContinue,
            Format=True,
            Replace=win32.constants.wdReplaceAll,
        )
        word_count += rng.Find.Found
        
        return word_count

if __name__ == "__main__":
    app = Application()
    app.mainloop()
