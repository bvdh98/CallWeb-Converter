from docx2python import docx2python
import re
from docx import Document
import pandas as pd
import CWConverter
import numpy as np
import unicodedata
import os
from termcolor import colored

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
# TODO: important test with surveys in diff format (qtext ,but not section col_names, and description)
# TODO: add logging
# TODO: test links in regular text
# TODO: make templates for repeating surveys
# TODO: important handle error when user doesn't pass in word doc
# TODO: important replace table method with beautiful soup


class Parser:
    def __init__(self):
        self.link = ''
        self.test_mode = False
        self.content = None
        self.questions = {}
        self.tbl_qs = {}
        # move cur vars to parse function
        self.cur_q = None
        self.cur_sec_desc = None
        self.cur_sec_header = None
        # TODO: important change to pandas df
        self.word_tables = {}
        self.flags = {'q_num': r'^[Q][0-9][.]|^[Q][1-9][0-9][.]',
                      'code': r'^[0-9][)]|^[1-9][0-9][)]',
                      'tbl_ref': 'tbl_q:'
                      }
        self.main()

    def get_survey_doc(self):
        while True:
            # get link to survey questionnaire word doc
            self.link = input(
                "enter the path to the survey questionnaire word document: \n")
            # remove quotation marks from link
            self.link = self.link.strip('"').strip('\'')
            # if not a link to a word doc, notify user
            if os.path.splitext(self.link)[1] != '.docx':
                print(colored('Only (.docx) word documents are accepted\n', color='red'))
                continue
            try:
                with open(self.link):
                    return docx2python(self.link)
            # if the file does not exist, notify user
            except FileNotFoundError:
                print(
                    colored("The questionnaire could not be found. Please enter the correct path\n", color='red'))
            # if the file can't be opened for whatever reason, notify user
            except e as e:
                print(colored(e, color='red'))

    def word_tbls_to_xlsx(self):
        writer = pd.ExcelWriter(
            'devtest/word_tables.xlsx', engine='xlsxwriter')
        for k, t in self.word_tables.items():
            t.to_excel(writer, sheet_name=f'table-{k}', index=False)
        writer.close()

    def main(self):
        document = self.get_survey_doc()
        # convert document to list with clean data
        self.content = self.get_clean_data(document.text)
        self.create_word_tables()
        self.create_table_questions()
        self.add_tbl_qs_ref_to_content()
        self.parse()
        if self.test_mode:
            # make workbook of word tables
            self.word_tbls_to_xlsx()
            df = pd.DataFrame(self.content, columns=['text'])
            # make csv of content
            df.to_csv('devtest/content.csv')
            # print out each question to a text file
            with open('devtest/qs.txt', 'w+') as f:
                [f.write(f'{q}\n') for q in self.questions.values()]
        # convert questions to callweb scw
        CWConverter.CWConverter(self.questions)
    # read over tables in document and store as dictionary of dataframes

    def create_word_tables(self):
        doc = Document(self.link)
        for i, table in enumerate(doc.tables):
            # if table is empty notify user and skip to next table
            if self.is_empty_tbl(table):
                print(
                    colored(
                        f"An empty table was found in this document. "
                        "If you are trying to reformat the tables in this document, please do so in a new blank document."
                        "Refer to the README on how to structure table questions.\n", color='yellow'))
                continue
            # store cells of table as 2d list
            cells = [[self.clean_str(cell.text)
                      for cell in row.cells] for row in table.rows]
            word_tble = pd.DataFrame(cells)
            # rename columns with table question first row: strongly agree, 4, ....
            word_tble = word_tble.rename(columns=word_tble.iloc[0]).drop(
                word_tble.index[0]).reset_index(drop=True)
            # append table to list
            self.word_tables[i] = word_tble

    def is_empty_tbl(self, tbl):
        # iterate over each row in table
        for row in tbl.rows:
            # iterate over each cell in row
            for cell in row.cells:
                # if cell contains characters other than whitespace return false
                if not re.search('^[ ]{0,}$', cell.text):
                    return False
        # if no non whitespace characters were found in the table return true
        return True

    def clean_str(self, str):
        # convert all unicode character to ASCII then convert to string
        str = unicodedata.normalize('NFKD', str).encode(
            'ascii', 'ignore').decode('utf-8')
        # replace multiple spaces with single space
        return re.sub('\s+', ' ', str)

    def create_table_questions(self):
        # TODO: important catch actual error eg: except RAISEVALUEERROR:
        # TODO: important clean tables like how content was cleaned (unicodes, extra spaces and tabs)
        # TODO replace double loop
        # for each table create a table question and add to tbl qs dictionary
        for tbl in self.word_tables.values():
            col_names = list(tbl.columns)
            col_names = self.clean_col_names(col_names)
            # column names may be empty if table isn't structured properly
            if len(col_names) == 0:
                # notify user about issue and move to next table
                print(colored(
                    f"The table with the question \"{tbl.iloc[0, 0]}\" doesn't have column names."
                    "Please refer to the README on how to structure table questions.\n", color='yellow'))
                continue
            try:
                # values of 5 point scale found in first row starting at second column
                scale = tbl.iloc[0, 1:].values.tolist()
                # scale may be empty if table isn't structured properly
            except:
                # notify user about issue and move to next table
                print(colored(
                    f"Could not read the table with the question \"{tbl.iloc[0, 0]}\"."
                    "Please refer to the README on how to structure table questions.\n", color='yellow'))
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
                    q_text=q_text, col_names=col_names, letter=q_letter, scale=scale)
                self.tbl_qs[q_text] = tbl_q

    def clean_col_names(self, col_names):
        # remove duplicates from column names
        col_names = list(dict.fromkeys(col_names))
        # remove blank column if it exists
        if '' in col_names:
            col_names.remove('')
        return col_names

    def add_tbl_qs_ref_to_content(self):
        # iterate over each table question in dictionary
        for q in self.tbl_qs.keys():
            # find index of text that matches table question
            # then subtract 1 from index to place the referenece before the text
            try:
                ref_indx = self.content.index(q)-1
            except Exception:
                # if table question can't be found, notify the user
                print(colored(f'Could not convert the table question: \"{q}\" to CallWeb.'
                      'Please refer to the README on how to structure table questions\n', color="yellow"))
                continue
            # avoid index out of bounds error
            if ref_indx >= 0:
                self.content[ref_indx] = f'tbl_q:{q}'

    def get_clean_data(self, data):
        # convert all unicode characters to ASCII then convert to string
        data = unicodedata.normalize('NFKD', data).encode(
            'ascii', 'ignore').decode('utf-8')
        # split the lines into a list
        data = data.split('\n')
        # remove trailing and leading spaces from each line if its not a tab, empty string or whitespace
        data = [str.strip() for str in data if str !=
                '\t' and not re.search('^[ ]{0,}$', str)]
        return data

    def parse(self):
        # iterate over each row in data frame
        for line_num, line in enumerate(self.content):
            # TODO: possibly add conditions (question is immediately after q flag or after sec header or after sec desc)
            # check if row is question text, eg: 1) the......
            if self.is_flag(flag=self.flags['q_num'], line=line, regex=True):
                # check if previous rows are related to the survey section
                self.check_for_section(line_num)
                # get number from start of question
                q_num = self.get_num(line)
                # remove number from start of question
                # callweb questions dont start with a number
                line = self.remove_flag(
                    line=line, flag=self.flags['q_num'], regex=True)
                # create new question and add it to questions dictionary
                self.questions[q_num] = Question(
                    num=q_num, sec_header=self.cur_sec_header, sec_desc=self.cur_sec_desc, q_text=line, codes={}, tbl_qs=[])
                # set the current question to this question
                self.cur_q = self.questions[q_num]
            # ensure that there is a current question to avoid none type error
            # checking for reference to table question
            elif self.cur_q and self.is_flag(self.flags['tbl_ref'], line):
                # extract table id from table reference and strip trailing and starting white space
                # table references are in the form: 'tbl_ref: Q1. question description/table id'
                tbl_id = line.replace(self.flags['tbl_ref'], '').strip()
                # get table questions
                tbl_q = self.tbl_qs[tbl_id]
                # add table question to current questions list of tbl qs
                self.cur_q.tbl_qs = tbl_q
                # update current question codes with table q codes
                self.cur_q.update_codes_from_tbl_q(tbl_q.codes)
                # skip to the next row since we just updated the codes
                continue
            # ensure that there is a current question to avoid none type error
            # checking for question code
            elif self.cur_q and self.is_flag(line=line, regex=True, flag=self.flags['code']):
                # remove code flag from row text
                line = self.remove_flag(
                    line=line, flag=self.flags['code'], regex=True)
                # update codes of question
                self.cur_q.codes = line

    def is_flag(self, flag, line, regex=False):
        # if regex is true, use regex library to look for pattern in line
        if regex:
            return re.search(flag, line)
        # if regex is false check if line contains flag
        return flag in line

    # get number from string
    def get_num(self, line):
        return int(re.findall(r'\d+', line)[0])

    def check_for_section(self, line_num):
        # get previous row if line number is at least 1 to avoid index out of bounds error
        prev_line = self.content[line_num-1] if line_num >= 1 else None
        # get row before previous row if line number is at least 2 to avoid index out of bounds error
        sec_prev_line = self.content[line_num-2] if line_num >= 2 else None
        # check if second previous row is section header (all caps) and not None
        if sec_prev_line and sec_prev_line.isupper():
            # set current section header to second prev row
            self.cur_sec_header = sec_prev_line
            # set the section description to prev row
            # by default the section description is after the header
            self.cur_sec_desc = prev_line
        # check if previous row is section header (all caps) and not none
        elif prev_line and prev_line.isupper():
            self.cur_sec_header = prev_line
            self.cur_sec_desc = None

    def remove_flag(self, line, flag, regex=False):
        # if regex is true, use regex library to remove flag from text
        if regex:
            return re.sub(flag, '', line).strip()
        # if regex is false replace flag with white space and then remove white space
        return line.replace(flag, '').strip()


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
        self._99_flags = ['don\'t know', 'dont know', 'no response',
                          'not applicable', 'prefer not to answer', 'no opinion']
        self._66_flags = ['other', 'please specify']

    # print out object in nicer format
    def __str__(self):
        return str(self.__class__) + '\n' + '\n'.join(('{} = {}'.format(item, self.__dict__[item]) for item in self.__dict__))

    @ property
    def has_oe_opt(self):
        return self._has_oe_opt

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
        self._tbl_qs.append(val)
    # TODO move function out of Question class

    def is_special_code(self, option, flags):
        for flag in flags:
            if flag in option.lower():
                return True
        return False


class TableQuestion(Question):
    def __init__(self, num=None, letter=None, sec_header=None, sec_desc=None, q_text=None, codes={}, q_note=None, tbl_qs=[], col_names=[], scale=[]):
        self._col_names = col_names
        self._scale = scale
        self._letter = letter
        Question.__init__(self, num, sec_header, sec_desc,
                          q_text, codes, q_note, tbl_qs)

    @property
    def col_names(self):
        return self._col_names

    @col_names.setter
    def col_names(self, val):
        self._col_names = val

    @property
    def scale(self):
        return self._scale

    @scale.setter
    def scale(self, val):
        self._scale = val

    @property
    def letter(self):
        return self._letter

    @letter.setter
    def letter(self, val):
        self._letter = val

    @property
    def codes(self):
        # check to make sure column names and scale are same length to avoid index out of bounds error
        if (len(self._col_names) != len(self._scale)):
            print(colored(
                f"The table with the question \"{self.q_text}\" does not have properly formatted column names."
                "Make sure that there are no duplicate, missing values, or subsections."
                "Please refer to the README on how to structure table questions.\n", color='yellow'))
            return {}
        else:
            # return codes as a dictionary of scale numbers followed by header eg: 5: Very satisfied
            return {self._scale[i]: self._col_names[i]
                    for i in range(len(self._scale))}


if __name__ == '__main__':
    parser = Parser()
