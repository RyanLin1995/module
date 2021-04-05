import sqlite3
import os
# source_file = "C:\\Users\\ryan.lin\\Desktop\\New Text Document.txt"
# end_file = "C:\\Users\\ryan.lin\\Desktop\\1.txt"
#
# sourcefile = open(source_file, encoding='UTF-8')
# file_read = sourcefile.read().split("\n")
# sourcefile.close()
# endfile = open(end_file, 'w', encoding='UTF-8')
# for value in file_read:
#     endfile.write("{}\n".format(value[41:]))
# endfile.close()

file = "C:\\temp\\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\\LocalState\\plum.sqlite"
source_file = "C:\\temp\\sticky_notes.txt"
notes_list = []
notes_list_source = []

with sqlite3.connect(file) as con:
    c = con.cursor()
    c.execute("SELECT name FROM sqlite_master WHERE type='table';")
    c.execute("SELECT * FROM 'Note'")
    results = c.fetchall()
    for i in results:
        note = list(i)
        for j in note:
            if "\\id" in str(j):
                notes_list_source.append(j)

for note_data_source in notes_list_source:
    notes_list.extend(note_data_source.split('\n'))

with open(source_file, 'w', encoding="utf-8") as f:
    for note_data in notes_list:
        f.write("{}\n".format(note_data[40:]))





