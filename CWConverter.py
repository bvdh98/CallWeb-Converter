class CWConverter:
    def __init__(self, survey):
        self._survey = survey
        self.end_of_q = '! =================================================='
        self.main()

    @ property
    def survey(self):
        return self._survey

    @ survey.setter
    def survey(self, val):
        self._survey = val
    # TODO: make generic function for write_tbl_qs and write_codes
    # TODO: important handle cases where prop == None (section header, description, codes)
    # TODO: replace with properties, eg: q.letter instead of q._letter
    # TODO: important tell user where to place code in full scw
    # TODO: add qcomp question to end

    def write_tbl_qs(self, qs, f):
        # for each table question write out callweb suffix code, eg: "[SUFFIX:_A] question description"
        for q in qs:
            f.write(f'\t[SUFFIX:_{q.letter}]{q.q_text}\n')
    # create text file to store call web code for table groups

    def create_tbl_doc(self):
        with open('table_groups.txt', 'w+') as t:
            t.write('##\tPlace this code underneath the Table section in the SCW\n')

    def append_tbl_group(self, qs, q_num):
        # get first table question
        first_tbl_q = qs[0].letter
        # get last table question
        last_tbl_q = qs[-1].letter
        # append to text document the associated callweb code
        with open('table_groups.txt', 'a') as t:
            t.write(
                f'\n\t#Group GRP_Q{q_num} = Q{q_num}_{first_tbl_q} - Q{q_num}_{last_tbl_q}\n')
            t.write(
                f'\t#Table GRP_Q{q_num} = Q{q_num}_{first_tbl_q} - Q{q_num}_{last_tbl_q}\n')

    def write_codes(self, codes, f):
        for code, value in codes.items():
            f.write(f'\t*{code}*{value}\n')

    def print_skipped_qs(self):
        # get list of survey questions
        keys = list(self.survey.keys())
        # notify user of questions not found in list
        [print(f'question {q} was skipped over. Please refer to the README to on how to structure your questions\n')
         for q in range(keys[0], keys[-1]+1) if q not in keys]

    def write_callweb_code(self):
        # iterate through each question in survey and convert into callweb code
        with open('survey.txt', 'w+') as f:
            f.write(
                '##\tPlace this code underneath the Survey Proper section in the SCW\n')
            f.write(
                '##\tYou will have to program any skips, display conditions, and the intro and closing questions yourself'
                'as well as custom MIN and MAX variable for any question\n')
            for q in self.survey.values():
                # TODO: check for when to set MIN or MAX to different val eg: Multi selects
                f.write(f'Q{q.num} MIN=1 MAX=1\n')
                f.write('% Question\n')
                # if question has section header write out callweb code for section header
                if q.sec_header:
                    f.write(f'\t<H2>{q.sec_header}</H2>\n')
                # if question has section description write out associated callweb code
                if q.sec_desc:
                    f.write(f'\t<P><strong>{q.sec_desc}</strong></P>\n')
                # write out question text
                f.write(f'\t<P><strong>{q.q_text}</strong></P>\n')
                f.write('% Note\n')
                # if question has table questions write out associated callweb code
                if q.tbl_qs:
                    self.write_tbl_qs(q.tbl_qs, f)
                    # append table group to table_groups.txt
                    self.append_tbl_group(q.tbl_qs, q.num)
                # write question codes if applicable
                f.write('% Codes\n')
                if q.codes:
                    self.write_codes(q.codes, f)
                f.write(f'% Skips\n')
                f.write(f'% Condition\n')
                # if question has open-ended option write callweb code for comment box
                if q.has_oe_opt:
                    f.write(f'% Open end\n')
                    f.write('\t66 = C250 5 50\n')
                else:
                    f.write(f'% Open end\n')
                f.write(f'{self.end_of_q}\n\n')

    def main(self):
        # if no questions were parsed, let the user know
        if len(self.survey) == 0:
            print(
                'Could not convert any questions in the survey to CallWeb.'
                'Please refer to the README on how to structure your survey\n')
            return
        # create table section
        self.create_tbl_doc()
        self.print_skipped_qs()
        self.write_callweb_code()
        print(
            'The survey was successfully converted to CallWeb code. '
            'Please look over the code in survey.scw and table_groups.txt. '
            'Note: the code produced may not be correct if you did not follow the template.'
            'Please refer to the README for more information.')
