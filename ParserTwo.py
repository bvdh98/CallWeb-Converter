from docx2python import docx2python
import re
from docx import Document
import pandas as pd
import numpy as np
from CWConverter import CWConverter

# REQUIREMENTS:
# dont format regular questions as tables like in saanich citz survey
# section header must be all caps not bold which makes text look all caps
# section description must be between header and question
# questions must start with Q followed by any single or double digit followed by .
# codes must start with --

# TODO REMOVE unused table rows
# TODO important test with word apostrophes and other word symbols
# TODO make ui prompt
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


class Parser:
    def __init__(self):
        self.link = "./surveys/23-985 Saanich 2022 Citizen Satisfaction Survey DRAFT.docx"
        self.content = None
        self.questions = {}
        # TODO important change to pandas df
        self.word_tables = {}
        # keep track of current question and current section header and description iterated over
        self.cur_q = None
        self.cur_sec_header = None
        self.cur_sec_desc = None
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
                      'code': '--',
                      'tbl_ref': 'word_tbl',
                      'tbl_cell': 'table cell',
                      'survey_end': '#end'
                      }
        self.main()

    def print_word_tbls(self):
        writer = pd.ExcelWriter('word_tables.xlsx', engine='xlsxwriter')
        for k, t in self.word_tables.items():
            t.to_excel(writer, sheet_name=f'table-{k}', index=False)
        writer.close()

    def main(self):
        document = docx2python(self.link)
        # convert document to list sepparated by end of line character
        lines = document.text.split('\n')
        # turn list into data frame
        self.content = pd.DataFrame(lines, columns=['text'])
        self.clean_data()
        self.create_word_tables()
        # self.parse()
        # convert questions to callweb scw
        # CWConverter(self.questions)

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
            # add a reference to the table in content.csv
            self.word_tables.to_csv('tables')
            self.add_table_ref_to_content(word_tble.iloc[0, 0], i)

    def create_table_questions(self):
        for tbl in self.word_tables.values():
            headers = list(tbl.columns)
            # remove duplicates from headers
            headers = list(dict.fromkeys(headers))
            # TODO make clean_headerS function that combines remove '' and remove newline char and removes duplicates
            # remove blank column
            headers.remove('')
            # remove newline character in header
            self.remove_newline_from_headers(headers)
            # values of 5 point scale found in first row starting at second column
            scale = tbl.iloc[0, 1:].values.tolist()
            # remove duplicates from scale
            scale = list(dict.fromkeys(scale))

    def add_table_ref_to_content(self, first_tbl_q, tbl_id):
        # remove trailing and ending spaces from question
        first_tbl_q = first_tbl_q.strip()
        # find index of row with text that matches table question
        # then subtract 0.5 from index
        # TODO important handle error where first table question can't be found
        ref_indx = self.content[self.content['text'].str.contains(
            first_tbl_q, regex=False)].index[0] - 0.5
        tbl_ref = {'text': f'word_tbl:{tbl_id}'}
        # append table refrence to content
        self.content.loc[ref_indx] = tbl_ref
        # sort table ref into correct position
        self.content = self.content.sort_index().reset_index(drop=True)

    def clean_data(self):
        # characters to filter out
        chars_to_replace = {'“': '\"', '”': '\"',
                            '’': '\'', '–': '-', '…': '...', 'é': 'e', '\s+': ' '}
        # replace all white space ('\s+') with single space.This will handle strings with consecutive white spaces
        # replace all docx characters above with callweb recognized equivalents
        self.content.replace({'text': chars_to_replace},
                             regex=True, inplace=True)
        # remove white space at start and end of string
        self.content = self.content['text'].str.strip()
        # convert content back to dataframe
        self.content = self.content.to_frame()
        # replace all whitespace with NAN
        self.content.replace('', np.nan, inplace=True)
        # drop NAN rows
        self.content.dropna(inplace=True)

    def parse(self):
        # iterate over each row in data frame
        for row in self.content.itertuples():
            # get value in text column
            r_text = getattr(row, 'text')
            # check if row is question text, eg: 1) the......
            if (self.is_q_text(r_text)):
                # check if previous rows are related to the survey section
                self.check_for_section(row.Index)
                # get number from start of question
                q_num = self.get_num(r_text)
                # remove number from start of question
                # callweb questions dont start with a number
                r_text = self.remove_qnum(r_text)
                # create new question and add it to questions dictionary
                self.questions[q_num] = Question(
                    q_num, self.cur_sec_header, self.cur_sec_desc, r_text, {}, None, None)
                # set the current question to this question
                self.cur_q = self.questions[q_num]
            # checking for reference to table question
            elif (self.is_flag(self.flags['tbl_ref'], r_text)):
                # extract table id from table reference
                tbl_id = int(self.get_num(r_text))
                # get table questions
                tbl_qs = self.get_tbl_qs(tbl_id)
                self.update_tbl_qs(self.cur_q, tbl_qs)
                # TODO improve
                # get the table codes
                tbl_codes = tbl_qs[0]._codes
                # update the current question codes with the above table codes
                self.cur_q.update_codes_from_tbl(tbl_codes)
                continue
            # ensure that there is a current question
            # checking for question code
            elif (self.cur_q and self.is_code(r_text)):
                # remove code flag from row text
                r_text = self.remove_code_flag(r_text)
                # update codes of question
                self.update_question(
                    self.cur_q, 'codes', r_text
                )
        # [print(q) for q in self.questions.values()]
    # update the question based on the property

    def update_question(self, q, prop, r_text):
        match prop:
            case 'sec_header':
                q.sec_header = r_text
            case 'sec_desc':
                q.sec_desc = r_text
            case 'text':
                q.q_text = r_text
            case 'codes':
                q.codes = r_text
    # TODO possibly combine function with one above

    def update_tbl_qs(self, q, tbl_qs):
        q.tbl_qs = tbl_qs
        # [print(q) for q in tbl_qs]

    def is_flag(self, flag, r_text):
        # check if row text contains flag
        return flag in r_text

    def remove_qnum(self, r_text):
        # replace question number 'Q1.' with space. then remove this space
        return re.sub(self.flags['q_num'], '', r_text).strip()

    # get number from string
    def get_num(self, r_text):
        return re.findall(r'\d+', r_text)[0]

    def remove_code_flag(self, r_text):
        # replace code flag '--' with space. then remove this space
        return r_text.replace(self.flags['code'], '').strip()

    def is_code(self, r_text):
        # true if question text contains code flag
        return self.is_flag(self.flags['code'], r_text)

    def check_for_section(self, row_ind):
        # get index of text column
        text_ind = self.content.columns.get_loc('text')
        # get previous row
        prev_rt = self.content.iloc[row_ind-1, text_ind]
        # get row before previous row
        sec_prev_rt = self.content.iloc[row_ind-2, text_ind]
        # check if second previous row is section header
        if (self.is_sec_header(sec_prev_rt)):
            # set current section header to second prev row
            self.cur_sec_header = sec_prev_rt
            # set the section description to prev row
            # by default the section description is after the header
            self.cur_sec_desc = prev_rt
        # check if previous row is section header
        elif (self.is_sec_header(prev_rt)):
            self.cur_sec_header = prev_rt

    def is_sec_header(self, r_text):
        # section headers are in all caps
        return r_text.isupper()

    # looking for text that starts with 'Q' followed by single or double digit followed by '.'
    def starts_with_q_num(self, r_text):
        return re.search(self.flags['q_num'], r_text)

    # TODO: possibly add conditions (question is immediately after q flag or after sec header or after sec desc)
    def is_q_text(self, r_text):
        # true if paragraph starts with number immediately followed by ')'
        return self.starts_with_q_num(r_text)

    def get_tbl_qs(self, tbl_id):
        # tbl qs to return
        qs = []
        # get table based on id
        tbl = self.word_tables[tbl_id]
        # header of table e.g.: Strongly agree 5, 4, ...
        headers = list(tbl.columns)
        # remove duplicates from headers
        headers = list(dict.fromkeys(headers))
        # TODO make clean_headerS function that combines remove '' and remove newline char and removes duplicates
        # remove blank column
        headers.remove('')
        # remove newline character in header
        self.remove_newline_from_headers(headers)
        # values of 5 point scale found in first row starting at second column
        scale = tbl.iloc[0, 1:].values.tolist()
        # remove duplicates from scale
        scale = list(dict.fromkeys(scale))
        # TODO: replace with pandas vectorization
        # iterate over questions found in first column and create table questions
        for i, q_text in enumerate(tbl.iloc[0:, 0].values):
            # create letter for question e.g.: A,B,C...
            q_letter = chr(i+65)
            tbl_q = TableQuestion(q_text, headers, q_letter, scale)
            qs.append(tbl_q)
        return qs
    # TODO remove extra spaces in DK/ NA/NR
    # TODO add <br> to text eg: Strongly<br />Disagree <br> 1
    # TODO fix: Very Satisfied5

    def remove_newline_from_headers(self, headers):
        for header in headers:
            # get index of header
            indx = headers.index(header)
            # remove newline char
            header = header.replace('\n', '')
            # update header in list
            headers[indx] = header


class TableQuestion:

    def __init__(self, q_text, headers, q_letter, scale):
        self._q_text = q_text
        self._headers = headers
        self._scale = scale
        self._q_letter = q_letter
        self._codes = {self._scale[i]: self._headers[i]
                       for i in range(len(self.scale))}

    def __str__(self):
        return str(self.__class__) + '\n' + '\n'.join(('{} = {}'.format(item, self.__dict__[item]) for item in self.__dict__))

    # @property
    # def codes(self):
    #     if(len(self.headers) != len(self.scale)):
    #         return None
    #     else:
    #         return {self._scale[i]: self._headers[i]
    #                    for i in range(len(self.scale))}

    @ property
    def q_text(self):
        return self._q_text

    @ q_text.setter
    def q_text(self, val):
        self._q_text = val

    @ property
    def q_letter(self):
        return self._q_letter

    @ q_letter.setter
    def q_letter(self, val):
        self._q_letter = val

    @ property
    def headers(self):
        return self._headers

    @ headers.setter
    def headers(self, val):
        self._headers = val

    @ property
    def scale(self):
        return self._scale

    @ scale.setter
    def scale(self, val):
        self._scale = val


class Question:
    def __init__(self, num, sec_header, sec_desc, q_text, codes, q_note, tbl_qs):
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

    def update_codes_from_tbl(self, tbl_codes):
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
        self._tbl_qs = val
    # TODO move function out of Question class

    def is_special_code(self, option, flags):
        for flag in flags:
            if flag in option.lower():
                return True
        return False


if __name__ == '__main__':
    parser = Parser()
