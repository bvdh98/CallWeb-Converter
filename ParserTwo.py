from docx2python import docx2python
import re
from docx import Document
import pandas as pd
import numpy as np
from CWConverter import CWConverter
# from bs4 import BeautifulSoup

# REQUIREMENTS:
# dont format regular questions as tables like in saanich citz survey
# section header must be all caps not bold which makes text look all caps
# section description must be between header and question
# questions must start with Q followed by any single or double digit followed by .
# links in tables will not be read

# TODO REMOVE unused table lines
# TODO: important test with word apostrophes and other word symbols
# TODO: important make ui prompt
# TODO update note property of question
# TODO add callweb code for comment boxes underneath please specify ()
# TODO: possibly add skips and conditions
# TODO handle intro and closing sections
# TODO: IMPORTANT organize methods into different classes instead of all in Parser class
# TODO: important notify user of missing questions and properties
# TODO: important error handling
# TODO: maybe change table questions into one dataframe instead of dictionary
# TODO: important print out callweb scw template
# TODO: handle mutiple table sections eg: q17 in saanich citz survey
# TODO: handle images
# TODO: handle q-ineligble section
# TODO: important test print missing data with oop
# TODO: important test with surveys in diff format (qtext ,but not section headers, and description)
# TODO: add logging
# TODO: add codes to docx template
# TODO: make table question child of question class
# TODO: read in dk and other responses from new template
# TODO: test links in regular text
# TODO: make templates for repeating surveys
# TODO: important handle error when user doesn't pass in word doc
# TODO: important handle error when user currently in scw or tables


class Parser:
    def __init__(self):
        self.link = "./surveys/23-985 Saanich 2022 Citizen Satisfaction Survey DRAFT.docx"
        self.content = None
        self.questions = {}
        self.tbl_qs = {}
        # TODO: important change to pandas df
        self.word_tables = {}
        self.flags = {'q_num': r'^[Q][0-9][.]|^[Q][1-9][0-9][.]',
                      'sec_header_start': '#qhs',
                      'sec_header_end': '#qhe#',
                      'sect_q_start': '#sq_s',
                      'sect_q_end': '#sq_e#',
                      'q_start': '#qs',
                      'q_end': '#qe#',
                      'q_text': '#qtext',
                      'option_start': '#ops',
                      'option_end': '#ope#',
                      'code': r'^[0-9][)]|^[1-9][0-9][)]',
                      'tbl_ref': 'tbl_q:',
                      'tbl_cell': 'table cell',
                      'survey_end': '#end',
                      'q_ineligible': 'Q-INELIGIBLE'
                      }
        self.main()

    def print_word_tbls(self):
        writer = pd.ExcelWriter('word_tables.xlsx', engine='xlsxwriter')
        for k, t in self.word_tables.items():
            t.to_excel(writer, sheet_name=f'table-{k}', index=False)
        writer.close()

    def main(self):
        document = docx2python(self.link)
        # convert document to list with clean data
        self.content = self.get_clean_data(document.text)
        self.create_word_tables()
        self.create_table_questions()
        self.add_tbl_qs_ref_to_content()
        # df = pd.DataFrame(self.content, columns=['text'])
        # df.to_csv('res.csv')
        self.parse()
        # convert questions to callweb scw
        CWConverter(self.questions)
    # read over tables in document and store as dictionary of dataframes

    def create_word_tables(self):
        doc = Document(self.link)
        for i, table in enumerate(doc.tables):
            # store cells of table as 2d list
            cells = [[cell.text for cell in row.cells] for row in table.rows]
            word_tble = pd.DataFrame(cells)
            # rename columns with table question first row: strongly agree, 4, ....
            word_tble = word_tble.rename(columns=word_tble.iloc[0]).drop(
                word_tble.index[0]).reset_index(drop=True)
            # append table to list
            self.word_tables[i] = word_tble

    def create_table_questions(self):
        # TODO replace double loop
        # for each table create a table question and add to tbl qs dictioanry
        for tbl in self.word_tables.values():
            headers = list(tbl.columns)
            headers = self.clean_headers(headers)
            # headers may be empty if table isn't structured properly
            if headers == None:
                # notify user about issue and move to next table
                print(
                    f"Could not read the table with the question \"{tbl.iloc[0, 0]}\"."
                    "Please refer to the README on how to structure table questions.\n")
                continue
            try:
                # values of 5 point scale found in first row starting at second column
                scale = tbl.iloc[0, 1:].values.tolist()
                # scale may be empty if table isn't structured properly
            except:
                # notify user about issue and move to next table
                print(
                    f"Could not read the table with the question \"{tbl.iloc[0, 0]}\"."
                    "Please refer to the README on how to structure table questions.\n")
                continue
            # remove duplicates from scale
            scale = list(dict.fromkeys(scale))
            # TODO: replace with pandas vectorization
            # iterate over questions found in first column and create table questions
            for i, q_text in enumerate(tbl.iloc[0:, 0].values):
                # create letter for question e.g.: A,B,C...
                q_letter = chr(i+65)
                # remove trailing and ending white space
                q_text = q_text.strip()
                # create table question and add it to dictionary
                tbl_q = TableQuestion(
                    q_text=q_text, headers=headers, letter=q_letter, scale=scale)
                self.tbl_qs[q_text] = tbl_q

    def clean_headers(self, headers):
        # remove duplicates from headers
        headers = list(dict.fromkeys(headers))
        # remove blank column if it exsists
        if '' in headers:
            headers.remove('')
        # remove newline character in header
        self.remove_newline_from_headers(headers)
        return headers

    # TODO remove extra spaces in DK/ NA/NR
    # TODO add <br> to text eg: Strongly<br />Disagree <br> 1
    # TODO fix: Very Satisfied5

    def remove_newline_from_headers(self, headers):
        for header in headers:
            # get index of header
            indx = headers.index(header)
            # remove newline character
            header = header.replace('\n', '')
            # update header in list
            headers[indx] = header
    # for each table question add a reference to it in the content

    def add_tbl_qs_ref_to_content(self):
        # iterate over each table question in dictionary
        for q in self.tbl_qs.keys():
            # find index of text that matches table question
            # then subtract 1 from index to place the referenece before the text
            try:
                ref_indx = self.content.index(q)-1
            except Exception:
                # if table question can't be found, notify the user
                print(f'Could not convert the table question: \"{q}\" to CallWeb.'
                      'Please refer to the README on how to structure table questions\n')
                continue
            # avoid index out of bounds error
            if ref_indx >= 0:
                self.content[ref_indx] = f'tbl_q:{q}'

    def get_clean_data(self, data):
        # characters to replace with CallWeb recognized equivalents
        chars_to_replace = {'“': '\"', '”': '\"',
                            '’': '\'', '–': '-', '…': '...', 'é': 'e', '\t': ''}
        # replace each character in dictionary above
        for k, v in chars_to_replace.items():
            data = [str.replace(k, v) for str in data]
        # convert back to string and then split the lines into a list
        data = ''.join(data).split('\n')
        # remove spaces from each line if its not an empty string
        data = [str.strip() for str in data if str != '']
        return data

    def parse(self):
        #keep track of last line iterated over
        last_line = None
        # keep track of current question and current section header and description iterated over
        cur_sec_header = None
        cur_sec_desc = None
        cur_q = None
        # iterate over each line
        for line in self.content:
            #check if line is section header
            if(line.isupper()):
                #update section header
                cur_sec_header = line
                #update last line flag
                last_line = 'sec_header'
                #skip to next line
                continue
            # check if line is question text, eg: 1) the......
            elif (self.is_q_text(line)):
                # get number from start of question
                q_num = self.get_num(line)
                # remove number from start of question
                # callweb questions dont start with a number
                line = self.remove_flag(
                    line=line, flag=self.flags['q_num'], regex=True)
                # create new question and add it to questions dictionary
                self.questions[q_num] = Question(
                    num=q_num, sec_header=cur_sec_header, sec_desc=cur_sec_desc, q_text=line, codes={}, tbl_qs=[])
                # set the current question to this question
                cur_q = self.questions[q_num]
                #reset the section description for next question
                cur_sec_desc = None
                last_line = 'q_text'
                # skip to the next line
                continue
            #if the last line is a section header and this current line is not question text, update section header to this line
            elif(last_line == 'sec_header'):
                cur_sec_desc = line
                last_line = 'sec_desc'
                continue
            # ensure that there is a current question to avoid none type error
            # checking for reference to table question
            elif (cur_q and self.is_flag(self.flags['tbl_ref'], line)):
                # extract table id from table reference and strip trailing and starting white space
                # table refrences are in the form: 'tbl_ref: Q1. question description/table id'
                tbl_id = line.replace(self.flags['tbl_ref'], '').strip()
                # get table questions
                tbl_q = self.tbl_qs[tbl_id]
                # add table question to current questions list of tbl qs
                cur_q.tbl_qs = tbl_q
                # update current question codes with table q codes
                cur_q.update_codes_from_tbl_q(tbl_q.codes)
                last_line = 'tbl_ref'
                continue
            # ensure that there is a current question to avoid none type error
            # checking for question code
            elif (cur_q and self.is_flag(line=line, regex=True, flag=self.flags['code'])):
                # remove code flag from line
                line = self.remove_flag(
                    line=line, flag=self.flags['code'], regex=True)
                # update codes of question
                cur_q.codes = line
                last_line = 'code'
                continue
            else: 
                last_line = None
        # [print(q) for q in self.questions.values()]

    def is_flag(self, flag, line, regex=False):
        # if regex is true, use regex library to look for pattern in line
        if regex:
            return re.search(flag, line)
        # if regex is false check if line contains flag
        return flag in line

    # get number from string
    def get_num(self, line):
        return int(re.findall(r'\d+', line)[0])

    def remove_flag(self, line, flag, regex=False):
        # if regex is true, use regex library to remove flag from text
        if regex:
            return re.sub(flag, '', line).strip()
        # if regex is false replace flag with white space and then remove white space
        return line.replace(flag, '').strip()

    def is_sec_header(self, line):
        # section headers are in all caps
        return line.isupper()

    # TODO: possibly add conditions (question is immediately after q flag or after sec header or after sec desc)
    def is_q_text(self, line):
        # true if paragraph starts with number immediately followed by ')'
        return self.is_flag(flag=self.flags['q_num'], line=line, regex=True)


class Question:
    def __init__(self, num=None, sec_header=None, sec_desc=None, q_text=None, codes={}, q_note=None, tbl_qs=[]):
        self._num = num
        self._sec_header = sec_header
        self._sec_desc = sec_desc
        self._q_text = q_text
        # TODO: order codes numerically for each question
        self._codes = codes
        self._q_note = q_note
        self._tbl_qs = tbl_qs
        self._has_oe_opt = False
        # TODO: move flags out of Question class
        self._99_flags = ['don\'t know', 'no response',
                          'not applicable', 'prefer not to answer']
        self._66_flags = ['other', 'please specify']

    # print out object in nicer format
    def __str__(self):
        return str(self.__class__) + '\n' + '\n'.join(('{} = {}'.format(item, self.__dict__[item]) for item in self.__dict__))

    @ property
    def sec_header(self):
        return self._sec_header

    @ sec_header.setter
    def sec_header(self, val):
        self._sec_header = val

    @ property
    def num(self):
        return self._num

    @ num.setter
    def num(self, val):
        self._num = val

    @ property
    def q_text(self):
        return self._q_text

    @ q_text.setter
    def q_text(self, val):
        self._q_text = val

    @ property
    def sec_desc(self):
        return self._sec_desc

    @ sec_desc.setter
    def sec_desc(self, val):
        self._sec_desc = val

    @ property
    def codes(self):
        return self._codes

    @ codes.setter
    def codes(self, val):
        # get key for codes dictionary (can't start at 0)
        key = len(self._codes) + 1
        # check if value is 99 code: dk/na
        if self.is_special_code(val, self._99_flags):
            self._codes[99] = val
        # check if value is 66 code: other/please specify
        elif self.is_special_code(val, self._66_flags):
            # set it true that question has open ended response
            self._has_oe_opt = True
            self._codes[66] = val
        else:
            self._codes[key] = val

    def update_codes_from_tbl_q(self, tbl_codes):
        self._codes = tbl_codes

    @ property
    def q_note(self):
        return self._q_note

    @ q_note.setter
    def q_note(self, val):
        self._q_note = val

    @ property
    def tbl_qs(self):
        return self._tbl_qs

    @ tbl_qs.setter
    def tbl_qs(self, val):
        # print(val)
        self._tbl_qs.append(val)
    # TODO move function out of Question class

    def is_special_code(self, option, flags):
        for flag in flags:
            if flag in option.lower():
                return True
        return False


class TableQuestion(Question):
    def __init__(self, num=None, letter=None, sec_header=None, sec_desc=None, q_text=None, codes={}, q_note=None, tbl_qs=[], headers=[], scale=[]):
        self._headers = headers
        self._scale = scale
        self._letter = letter
        Question.__init__(self, num, sec_header, sec_desc,
                          q_text, codes, q_note, tbl_qs)

    @property
    def codes(self):
        # check to make sure headers and scale are same length to avoid index out of bounds error
        if (len(self._headers) != len(self._scale)):
            return None
        else:
            # return codes as a dictionary of scale followed by header eg: 5: Very satisfied
            return {self._scale[i]: self._headers[i]
                    for i in range(len(self._scale))}


if __name__ == '__main__':
    parser = Parser()
