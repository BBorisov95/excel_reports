import pandas as pd

#company_new_period = excel_df.loc[excel_df['Unnamed: 0'] == 'Цялостен резултат']
#company_new_period_value = company_new_period.iloc[:, 6].values


class Report():
    def __init__(self, file):
        self.file = file
        self.company_new_period = file.loc[file['Unnamed: 0'] == 'Цялостен резултат']
        self.company_new_period_value = self.company_new_period.iloc[:, 6].values

class ReportOld():
    def __init__(self, file):
        super().__init__(self, file)
    
    def old_period_turnover(self):
        ot3_old_turnover = self.file.iloc[:, 2]
        return ot3_old_turnover.values
    
    def old_period_sold_pcs(self):
        ot3_old_sold_pcs = self.file.iloc[:, 11]
        return ot3_old_sold_pcs.values
    
class ReportNew():
    def __init__(self, file):
        super().__init__(file)
    
    def new_period_turnover(self):
        ot3_new_turnover = self.file.iloc[:, 5]
        return ot3_new_turnover.values
    
    def new_period_sold_pcs(self):
        ot3_new_sold_pcs = self.file.iloc[:, 14]
        return ot3_new_sold_pcs.values
    

        
        
#excel_df = pd.read_excel(r'C:\Users\Bobi\Desktop\Сравнение по групи за периода 22.10-28.10.xls', sheet_name='Table', skiprows=(0, 2, 3))

ot3 = Report(pd.read_excel(r'C:\Users\Bobi\Desktop\Сравнение по групи за периода 22.10-28.10.xls', sheet_name='Table', skiprows=(0, 2, 3)))


ot3_difference_in_percent = (((ot3.new_period_turnover() / ot3.old_period_turnover() - 1) * 100)
ot3_difference_in_bgn = (ot3.new_period_turnover() - ot3.old_period_turnover())
percent_from_company = ((ot3.new_period_turnover() / Report.self.company_new_period_value) * 100)
sold_value_new = (ot3.new_period_turnover() / ot3.new_period_sold_pcs())
sold_value_old = (ot3.old_period_turnover() / ot3.old_period_sold_pcs)
 
