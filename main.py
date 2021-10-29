import os
from tkinter.filedialog import askdirectory
from win32com import client


class TDC:
    def __init__(self):
        self.path = os.path.normpath(askdirectory(title="Select Folder"))
        self.emails = [file for file in os.listdir(self.path) if file.endswith(".msg")]
        self.outlook = client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.numbers = set()

    def extract(self):
        for index in range(len(self.emails)):
            path = os.path.join(self.path, self.emails[index])
            msg = self.outlook.OpenSharedItem(path).HTMLBody
            number = self.extract_number(msg)
            invoice = self.extract_invoice(msg)
            if invoice[0].lower() != "m":
                continue
            self.numbers.add(", ".join([number, invoice]) + "\n")
        return self.numbers

    def extract_to_file(self, name="output.txt"):
        with open(name, "w") as out_file:
            for entry in self.extract():
                out_file.write(entry)

    def extract_number(self, msg):
        index = msg.find("Mobilnummer")
        return msg[index + 644 : index + 652]

    def extract_invoice(self, msg):
        index = msg.find("Fakturabe")
        note = msg[index + 206 : index + 240]
        return note[: note.index("\r")]


bot = TDC()
bot.extract_to_file()
