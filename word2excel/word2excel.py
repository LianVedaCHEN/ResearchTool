import os
from docx import Document
import re
import pandas as pd
class Word2Excel(object):


    def file_name(self,file_dir):
        for root, dirs, files in os.walk(file_dir):
            print(root)  # current path
            print(dirs)  # subfolder under current path
            print(files)  # non-folder files under current path
        return files

    def __init__(self,file_dir):
        self.file_list = [('files/' + v) for v in self.file_name(file_dir)  if v != '.DS_Store' ]
        print(self.file_list)

    def __read_word__(self,file):
        document = Document(file)
        return document

    def __one_file_process__(self,file):
        document = self.__read_word__(file)
        paragraphs = [p.text for p in document.paragraphs]
        court = paragraphs[0]
        document_type = paragraphs[1]
        seriesID = ''
        prog = re.compile("\（(20[0-1][0-9])\）")
        result = prog.search(paragraphs[2])
        if result:
            seriesID = paragraphs[2]
        end_part = paragraphs[-6:]
        judge = []
        secretary = []
        judge_number = 0
        secretary_number = 0
        for i,paragraph in enumerate(end_part):
            prog2 = re.compile("(审((.)|)判((.)|)长)|(审((.)|)判((.)|)员)")
            result = prog2.search(paragraph)
            if result:
                judge.append(paragraph)
                judge_number += 1
                print(judge_number)
                continue
            prog2 = re.compile("(书((.)|)记((.)|员))")
            result2 = prog2.search(paragraph)
            if result2:
                secretary.append(paragraph)
                secretary_number = len(end_part) - i
                print(secretary_number)
                break
        secretary = end_part[-secretary_number]
        date = paragraphs[len(paragraphs) - 1 - secretary_number]
        if(seriesID == ''):
            start = 2
            end = len(paragraphs)-2-judge_number-secretary_number
            content = paragraphs[start:end]
        else:
            start = 3
            end = len(paragraphs) - 2 - judge_number - secretary_number
            content = paragraphs[start:end]
        doc ={'court':court,
        'document type':document_type, 'series ID':seriesID, 'Judge Number': judge_number, 'Secretary Number': secretary_number,
        'date': date,'content':content, 'judges':judge, 'secretary':secretary}
        return doc
    def __process__(self):
        doc_list = [ self.__one_file_process__(file) for file in self.file_list ]
        return doc_list

    def word_2_excel(self,excel_name):
        doc_list = self.__process__()
        df = pd.DataFrame(data = doc_list)
        writer = pd.ExcelWriter(excel_name)

        df.to_excel(writer, 'Sheet1')
        writer.save()



word2excel = Word2Excel('files')
word2excel.word_2_excel("output.xlsx")
