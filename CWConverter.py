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
    #TODO: make generic function for write_tbl_qs and write_codes
    #TODO: important handle cases where prop == None (section header, description, codes)
    def write_tbl_qs(self, qs, f):
        #for each table question write out callweb suffix code, eg: "[SUFFIX:_A] question description"
        for q in qs:
            f.write(f'\t[SUFFIX:_{q._q_letter}]{q._q_text}\n')
    #create text file to store call web code for table groups
    def create_tbl_doc(self):
        with open('table_groups.txt', 'w+') as t:
            t.write('##\tTables\n')

    def append_tbl_group(self,qs,q_num):
        #get first table question
        first_tbl_q = qs[0]._q_letter
        #get last table question
        last_tbl_q = qs[-1]._q_letter
        #append to text document the associated callweb code
        with open('table_groups.txt', 'a') as t:
            t.write(f'\n\t#Group GRP_Q{q_num} = Q{q_num}_{first_tbl_q} - Q{q_num}_{last_tbl_q}\n')
            t.write(f'\t#Table GRP_Q{q_num} = Q{q_num}_{first_tbl_q} - Q{q_num}_{last_tbl_q}\n')

    def write_codes(self,codes,f):
        f.write('% Codes\n')
        for code,value in codes.items():
            f.write(f'\t*{code}*{value}\n')

    def print_missing_data(self):
        keys = list(self.survey.keys())
        missing_qs = [q for q in range(keys[0], keys[-1]+1) if q not in keys]
        for q in self.survey:
            for attr,val in q.__dict__.items():
                pass

        
    def main(self):
        #create table section
        self.create_tbl_doc()
        #iterate through each question in survey and convert into callweb scw file
        with open('survey.scw', 'w+') as f:
            for q in self.survey.values():
                # TODO: check for when to set MIN or MAX to different val eg: Multi selects
                f.write(f'Q{q._num} MIN=1 MAX=1\n')
                f.write('% Question\n')
                #if question has section header write out callweb code for section header
                if(q._sec_header):
                    f.write(f'\t<H2>{q._sec_header}</H2>\n')
                #if question has section description write out associated callweb code
                if q._sec_desc:
                    f.write(f'\t<P><strong>{q._sec_desc}</strong></P>\n')
                #write out question text
                f.write(f'\t<P><strong>{q._q_text}</strong></P>\n')
                f.write('% Note\n')
                #if question has table questions write out associated callweb code
                if q._tbl_qs:
                    self.write_tbl_qs(q._tbl_qs, f)
                    #append table group to table_groups.txt
                    self.append_tbl_group(q._tbl_qs,q._num)
                #write question codes
                self.write_codes(q._codes,f)
                f.write(f'% Skips\n')
                f.write(f'% Condition\n')
                #if question has open-ended option write callweb code for comment box
                if q._has_oe_opt:
                    f.write(f'% Open end\n')
                    f.write('\t66 = C250 5 50\n')
                else:
                    f.write(f'% Open end\n')
                f.write(f'{self.end_of_q}\n\n')
