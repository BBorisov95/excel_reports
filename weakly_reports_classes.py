import sys
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing


def show_exception_and_exit(exc_type, exc_value, tb):
    import traceback
    traceback.print_exception(exc_type, exc_value, tb)
    input("Press key to exit.")
    sys.exit(-1)

sys.excepthook = show_exception_and_exit

class Report:
    def __init__(self, file):
        self.file = file.loc[excel_df['Unnamed: 0'] == 'Отдел 3 Кар Аудио']

    def old_period_turnover(self):
        ot3_old_turnover = self.file.iloc[:, 2]
        return ot3_old_turnover.values
    
    def old_period_sold_pcs(self):
        ot3_old_sold_pcs = self.file.iloc[:, 11]
        return ot3_old_sold_pcs.values

    def new_period_turnover(self):
        ot3_new_turnover = self.file.iloc[:, 5]
        return ot3_new_turnover.values

    def new_period_sold_pcs(self):
        ot3_new_sold_pcs = self.file.iloc[:, 14]
        return ot3_new_sold_pcs.values


class ReportWithoutGamAndRec:
    def __init__(self, file):
        self.file_gam = file.loc[excel_df['Unnamed: 0'] == 'ГЕЙМИНГ']
        self.file_rec = file.loc[excel_df['Unnamed: 0'] == 'КНИГИ И ДИСКОВЕ']

    def gam_new_period_turnover(self):
        new_period_gam = self.file_gam.iloc[:, 5].values[0]
        return new_period_gam

    def gam_old_period_turnover(self):
        gam_old_turnover = self.file_gam.iloc[:, 2].values[0]
        return gam_old_turnover

    def rec_new_period_turnover(self):
        rec_new_trunover = self.file_rec.iloc[:, 5].values
        return rec_new_trunover

    def rec_old_period_turnover(self):
        rec_old_turnover = self.file_rec.iloc[:, 2].values
        return rec_old_turnover

    def gam_new_period_sold_pcs(self):
        gam_sold_pcs_new = self.file_gam.iloc[:, 14].values[0]
        return gam_sold_pcs_new

    def gam_old_period_sold_pcs(self):
        gam_sold_pcs_old = self.file_gam.iloc[:, 11].values[0]
        return gam_sold_pcs_old

    def rec_new_period_sold_pcs(self):
        rec_sold_pcs_new = self.file_rec.iloc[:, 14].values
        return rec_sold_pcs_new

    def rec_old_period_sold_pcs(self):
        rec_sold_pcs_old = self.file_rec.iloc[:, 11].values
        return rec_sold_pcs_old

    def final_reports_without_gam_and_rec(self):
        ot3_without_gam_rec_new = Report.new_period_turnover(ot3) - ReportWithoutGamAndRec.gam_new_period_turnover(self.file_gam) \
                                  - ReportWithoutGamAndRec.rec_new_period_turnover(self.file_rec)
        ot3_without_gam_rec_old = Report.old_period_turnover(ot3) - ReportWithoutGamAndRec.gam_old_period_turnover(self.file_gam) \
                                  - ReportWithoutGamAndRec.rec_old_period_turnover(self.file_rec)
        ot3_sold_pcs_without_gam_rec_new = Report.new_period_sold_pcs(ot3) - ReportWithoutGamAndRec.gam_new_period_sold_pcs(self.file_gam) \
                                           - ReportWithoutGamAndRec.rec_new_period_sold_pcs(self.file_rec)
        ot3_sold_pcs_without_gam_rec_old = Report.old_period_sold_pcs(ot3) - ReportWithoutGamAndRec.gam_old_period_sold_pcs(self.file_gam) \
                                          - ReportWithoutGamAndRec.rec_old_period_sold_pcs(self.file_rec)
        return ot3_without_gam_rec_new, ot3_without_gam_rec_old, ot3_sold_pcs_without_gam_rec_new, ot3_sold_pcs_without_gam_rec_old


def final_print():
    print(f'{"=" * 50}')
    print(f'Отдел 3 Кар Аудио с гейминг и рекърдс е с {ot3_difference_in_percent} в %')
    print(f'Отдел 3 Кар Аудио с гейминг и рекърдс е с {ot3_difference_in_bgn} в bgn')
    print(f'Отдел 3 Кар Аудио с гейминг и рекърдс е {percent_from_company} % от Технополис')
    print(f'Средна продажна цена с гейминг и рекърдс е {sold_value_new} лв без ДДС.', end=' ')
    print(f'За минал период сумата е била {sold_value_old}')
    print(f'{"=" * 50}')
    print(f'{"=" * 50}')
    print(f'Отдел 3 Кар Аудио БЕЗ гейминг и рекърдс е с {ot3_without_gam_rec_difference_in_percent} в %')
    print(f'Отдел 3 Кар Аудио БЕЗ гейминг и рекърдс е с {ot3_without_gam_rec_difference_in_bgn} в bgn')
    print(f'Отдел 3 Кар Аудио БЕЗ гейминг и рекърдс е {percent_from_company_without_gam_rec} % от Технополис')
    print(f'Средна продажна цена БЕЗ гейминг и рекърдс е {sold_value_without_gam_rec_new} лв без ДДС.', end=' ')
    print(f'За минал период сумата е била {sold_value_without_gam_rec_old}')
    print(f'{"=" * 50}')

excel_file = askopenfilename()
excel_df = pd.read_excel(excel_file, sheet_name='Table', skiprows=(0, 2, 3))

company_new_period = excel_df.loc[excel_df['Unnamed: 0'] == 'Цялостен резултат']  # catching row with specific name
company_new_period_value = company_new_period.iloc[:, 6].values  # return values of the column

ot3 = Report(excel_df)
ot3_gam_rec = ReportWithoutGamAndRec(excel_df)

#  Third department reports
ot3_difference_in_percent = ((Report.new_period_turnover(ot3) / Report.old_period_turnover(ot3) - 1) * 100)
ot3_difference_in_bgn = (Report.new_period_turnover(ot3) - Report.old_period_turnover(ot3))
percent_from_company = ((Report.new_period_turnover(ot3) / company_new_period_value) * 100)
sold_value_new = (Report.new_period_turnover(ot3) / Report.new_period_sold_pcs(ot3))
sold_value_old = (Report.old_period_turnover(ot3) / Report.old_period_sold_pcs(ot3))

#  Third department reports without Gaming and Records category

ot3_without_gam_rec_new = Report.new_period_turnover(ot3) - \
                          ReportWithoutGamAndRec.gam_new_period_turnover(ot3_gam_rec) \
                          - ReportWithoutGamAndRec.rec_new_period_turnover(ot3_gam_rec)
ot3_without_gam_rec_old = Report.old_period_turnover(ot3) - \
                          ReportWithoutGamAndRec.gam_old_period_turnover(ot3_gam_rec) \
                          - ReportWithoutGamAndRec.rec_old_period_turnover(ot3_gam_rec)
ot3_sold_pcs_without_gam_rec_new = Report.new_period_sold_pcs(ot3) - \
                                   ReportWithoutGamAndRec.gam_new_period_sold_pcs(ot3_gam_rec) \
                                   - ReportWithoutGamAndRec.rec_new_period_sold_pcs(ot3_gam_rec)
ot3_sold_pcs_without_gam_rec_old = Report.old_period_sold_pcs(ot3) - \
                                   ReportWithoutGamAndRec.gam_old_period_sold_pcs(ot3_gam_rec) \
                                   - ReportWithoutGamAndRec.rec_old_period_sold_pcs(ot3_gam_rec)

ot3_without_gam_rec_difference_in_percent = (((ot3_without_gam_rec_new / ot3_without_gam_rec_old) - 1) * 100)
ot3_without_gam_rec_difference_in_bgn = (ot3_without_gam_rec_new - ot3_without_gam_rec_old)
percent_from_company_without_gam_rec = ((ot3_without_gam_rec_new / company_new_period_value) * 100)
sold_value_without_gam_rec_new = (ot3_without_gam_rec_new / ot3_sold_pcs_without_gam_rec_new)
sold_value_without_gam_rec_old = (ot3_without_gam_rec_old / ot3_sold_pcs_without_gam_rec_old)

final_print()
