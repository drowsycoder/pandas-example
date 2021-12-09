# This module needs following libraries to be installed (by pip install):
# - beautifulsoup4 (for parsing of internet pages)
# - openpyxl (for xlsx support)
# - pandas (for dataframes operations)
# - requests (for download of internet pages)
# - xlsxwriter (for column width adjustment in xlsx)

from bs4 import BeautifulSoup as bs
import pandas as pd
import requests

BASE_URL = 'http://duma.gov.ru/duma/deputies'
CURRENT_DUMA_CONVOCATION = 8
CONVOCATION_NUMBER_ENDING = '-й созыв'
NAME_DEPUTY_SHEET = 'Депутаты Госдумы'
NAME_DEPUTY_COLUMN = 'Депутат'
NAME_TOTAL_COLUMN = 'Всего созывов'


def form_urls_for_all_duma_times(current_time):
    urls = []
    for s in range(current_time + 1):
        urls.append(f'{BASE_URL}/{str(s)}/')
    return urls


def retrieve_deputies_list(url):
    page = requests.get(url)
    soup = bs(page.text, 'html.parser')

    deputies_names = []
    for person_item in soup.findAll('li', class_='list-persons__item'):
        fio = name = fam_name = ''
        for name_item in person_item.findAll('span', itemprop='name'):
            fam_name = name_item.find('strong').text
            name = name_item.find('span', class_='second-name').text
            full_name = fam_name + ' ' + name
            deputies_names.append(full_name)

    return deputies_names


def save_to_excel_with_columns_adjustment(df, filename):
    writer = pd.ExcelWriter(filename)
    df.to_excel(writer, sheet_name=NAME_DEPUTY_SHEET, index=False, na_rep='')

    # Auto-adjust columns' width
    for column in df:
        column_width = max(
            df[column].astype(str).map(len).max() + 10, len(column)
        )
        col_idx = df.columns.get_loc(column)
        writer.sheets[NAME_DEPUTY_SHEET].set_column(
            col_idx, col_idx, column_width
        )

    writer.save()


def main():
    all_deputies_df = pd.DataFrame(columns=[NAME_DEPUTY_COLUMN])

    # For time saving while testing descrease range size here.
    for convocation_number in range(1, CURRENT_DUMA_CONVOCATION + 1):
        url = f'{BASE_URL}/{convocation_number}'
        current_df = pd.DataFrame(
            retrieve_deputies_list(url), columns=[NAME_DEPUTY_COLUMN]
        )
        current_df[f'{convocation_number}{CONVOCATION_NUMBER_ENDING}'] = '+'
        print(current_df)

        all_deputies_df = pd.merge(
            all_deputies_df,
            current_df,
            how='outer',
            left_on=NAME_DEPUTY_COLUMN,
            right_on=NAME_DEPUTY_COLUMN
        )

    all_deputies_df[NAME_TOTAL_COLUMN] = all_deputies_df[(
        all_deputies_df.columns[1:]
    )].notnull().sum(axis=1)

    all_deputies_df.drop_duplicates(inplace=True, ignore_index=True)

    # EXAMPLE OF STATISTICS

    longest_full_name = all_deputies_df[NAME_DEPUTY_COLUMN].str.len().max()
    shortest_full_name = all_deputies_df[NAME_DEPUTY_COLUMN].str.len().min()
    print(f'Самое длинное Ф.И.О. ({longest_full_name - 2} символ(а,ов)):',
        all_deputies_df[all_deputies_df[NAME_DEPUTY_COLUMN].str.len() ==
        longest_full_name][NAME_DEPUTY_COLUMN],
        sep='\n', end='\n\n'
    )
    print(f'Самое короткое Ф.И.О. ({shortest_full_name - 2} символ(а,ов)):',
        all_deputies_df[all_deputies_df[NAME_DEPUTY_COLUMN].str.len() ==
        shortest_full_name][NAME_DEPUTY_COLUMN],
        sep='\n', end='\n\n'
    )
    print(f'Депутаты, которые были во всех созывах',
        all_deputies_df[
            ~all_deputies_df.isnull().any(axis=1)
        ][NAME_DEPUTY_COLUMN],
        sep='\n', end='\n\n'
    )

    # DATA EXPORT:

    all_deputies_df = all_deputies_df.sort_values(
        by=NAME_DEPUTY_COLUMN
    ).reset_index(drop=True)
    save_to_excel_with_columns_adjustment(
        all_deputies_df, 'Deputies_by_name.xlsx'
    )

    all_deputies_df = all_deputies_df.sort_values(
        by=NAME_TOTAL_COLUMN, ascending=False
    ).reset_index(drop=True)
    save_to_excel_with_columns_adjustment(
        all_deputies_df, 'Deputies_by_times.xlsx'
    )


if __name__ == '__main__':
   main()
