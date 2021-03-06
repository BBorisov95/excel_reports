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

# importing table
excel_file = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
excel_df = pd.read_excel(excel_file, sheet_name='Table', skiprows=(0, 2, 3))

# company total always in new period
company_new_period = excel_df.loc[excel_df['Unnamed: 0'] == 'Цялостен резултат']
company_new_period_value = company_new_period.iloc[:, 6].values

ot3 = excel_df.loc[excel_df['Unnamed: 0'] == 'Отдел 3 Кар Аудио']

ot3_new_period = ot3.iloc[:, 5]
ot3_old_period = ot3.iloc[:, 2]
sold_pcs_new = ot3.iloc[:, 14]
sold_pcs_old = ot3.iloc[:, 11]
ot3_difference_in_percent = (((ot3_new_period / ot3_old_period) - 1) * 100).values
ot3_difference_in_bgn = (ot3_new_period - ot3_old_period).values
percent_from_company = ((ot3_new_period / company_new_period_value) * 100).values
sold_value_new = (ot3_new_period / sold_pcs_new).values
sold_value_old = (ot3_old_period / sold_pcs_old).values
print(f'{"="*50}')
print(f'Отдел 3 Кар Аудио с гейминг и рекърдс е с {ot3_difference_in_percent} в %')
print(f'Отдел 3 Кар Аудио с гейминг и рекърдс е с {ot3_difference_in_bgn} в bgn')
print(f'Отдел 3 Кар Аудио с гейминг и рекърдс е {percent_from_company} % от Технополис')
print(f'Средна продажна цена с гейминг и рекърдс е {sold_value_new} лв без ДДС.', end=' ')
print(f'За минал период сумата е била {sold_value_old}')
print(f'{"="*50}')
# gam and rec
gam = excel_df.loc[excel_df['Unnamed: 0'] == 'ГЕЙМИНГ']
rec = excel_df.loc[excel_df['Unnamed: 0'] == 'КНИГИ И ДИСКОВЕ']
gam_new_period = gam.iloc[:, 5].values[0]
gam_old_period = gam.iloc[:, 2].values[0]
rec_new_period = rec.iloc[:, 5].values
rec_old_period = rec.iloc[:, 2].values
gam_sold_pcs_new = gam.iloc[:, 14].values[0]
gam_sold_pcs_old = gam.iloc[:, 11].values[0]
rec_sold_pcs_new = rec.iloc[:, 14].values
rec_sold_pcs_old = rec.iloc[:, 11].values

# without gam and rec

ot3_without_gam_rec_new = ot3_new_period - gam_new_period - rec_new_period
ot3_without_gam_rec_old = ot3_old_period - gam_old_period - rec_old_period
ot3_sold_pcs_without_gam_rec_new = sold_pcs_new - gam_sold_pcs_new - rec_sold_pcs_new
ot3_sold_pcs_without_gam_rec_old = sold_pcs_old - gam_sold_pcs_old - rec_sold_pcs_old

ot3_without_gam_rec_difference_in_percent = (((ot3_without_gam_rec_new / ot3_without_gam_rec_old) - 1) * 100).values
ot3_without_gam_rec_difference_in_bgn = (ot3_without_gam_rec_new - ot3_without_gam_rec_old).values
percent_from_company_without_gam_rec = ((ot3_without_gam_rec_new / company_new_period_value) * 100).values
sold_value_without_gam_rec_new = (ot3_without_gam_rec_new / ot3_sold_pcs_without_gam_rec_new).values
sold_value_without_gam_rec_old = (ot3_without_gam_rec_old / ot3_sold_pcs_without_gam_rec_old).values
print(f'{"="*50}')
print(f'Отдел 3 Кар Аудио БЕЗ гейминг и рекърдс е с {ot3_without_gam_rec_difference_in_percent} в %')
print(f'Отдел 3 Кар Аудио БЕЗ гейминг и рекърдс е с {ot3_without_gam_rec_difference_in_bgn} в bgn')
print(f'Отдел 3 Кар Аудио БЕЗ гейминг и рекърдс е {percent_from_company_without_gam_rec} % от Технополис')
print(f'Средна продажна цена БЕЗ гейминг и рекърдс е {sold_value_without_gam_rec_new} лв без ДДС.', end=' ')
print(f'За минал период сумата е била {sold_value_without_gam_rec_old}')
print(f'{"="*50}')

# user choices

group_choice = ''

while group_choice != 'exit':
    group_choice = input('Insert individual filter or type exit: ')
    print()
    group_filter = excel_df.loc[excel_df['Unnamed: 0'] == group_choice]
    # new period

    new_period = group_filter.iloc[:, 5]
    # old period
    old_period = group_filter.iloc[:, 2]
    # grow
    difference_in_percent = (((new_period / old_period) - 1) * 100).values
    difference_in_bgn = (new_period - old_period).values
    group_choice_percent_from_company = ((new_period / company_new_period_value) * 100).values

    print(f'{group_choice} е {difference_in_percent} в %')
    print(f'{group_choice} е {difference_in_bgn} в bgn')
    print(f'{group_choice} е {group_choice_percent_from_company} % от Технополис')
    print()
    if group_choice == 'EXIT':
        break
