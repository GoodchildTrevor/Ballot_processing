#imports
import pandas as pd
import numpy as np
import re
import warnings

warnings.filterwarnings("ignore")

def extract_year(title):
    '''
    :param title: movie titles, from which I extract year using regular expressions
    :return: year as integer or NaN
    '''
    match = re.search(r'\b(19|20)\d{2}\b', title)
    if match:
        return int(match.group(0))
    return np.nan

def define_decade(year):
    '''
    :param year: find decade using dividing by 10
    :return: decade or NaN
    '''
    if not np.isnan(year):
        return (year // 10) * 10
    return np.nan

name = 'Criminal' #mark genre of poll

data = pd.read_excel(f'{name}.xlsx') #reading file

users=data['Ваш ник на Форуме Кинопоиска:'] #get users' names from the dataframe

data = data.drop(columns = ['Отметка времени','Ваш ник на Форуме Кинопоиска:']) #delete useless column with dates

results = [] #create list with rows of new df

for x in range(len(data)):
    for y in range(0, len(data.iloc[x])):
        row = {
            'пользователь': users[x],
            'фильмы': data.at[x, f'Лучший фильм {y+1} место'],
            'упоминания': 1,
            'баллы': 25 - y,
            'позиция': y+1
        }
        results.append(row)

top = pd.DataFrame(results) #create new df to aggregate results

top['фильмы'] = top['фильмы'].apply(lambda x: x.strip()) #delete useless spaces
top['год'] = top['фильмы'].apply(extract_year)
top['декада'] = top['год'].apply(define_decade) #find decade using functions

grouped_top = top.groupby('фильмы').agg({
    'упоминания': 'sum',
    'баллы': 'sum'
}).reset_index() #get sum of mentions and points for every movie
grouped_positions = top.groupby('фильмы')['позиция'].apply(list).reset_index() #collecting positions
merged_group = pd.merge(grouped_top, grouped_positions, on='фильмы') #merging dfs

merged_group['позиции'] = merged_group['позиция'].apply(lambda x: tuple(sorted(x))) #sorting positins
sorted_top = merged_group.sort_values(by=['баллы', 'упоминания', 'позиции'], ascending=[False, False, True]) #final sorting

sorted_top = sorted_top.drop(columns=['позиция']) #removing source column
sorted_top.reset_index(drop=True)

sorted_top = pd.merge(sorted_top, top[['фильмы', 'декада']], on='фильмы', how='left').drop_duplicates() #adding decades to final df

sorted_top.to_excel(f'Results_{name}.xlsx', index=False) #saving file


