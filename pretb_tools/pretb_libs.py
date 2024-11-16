'''
----------------------------------------------------------------------------------------------------------------------------------------
Author: Sherif Mehralivand
Email: sherif.mehralivand@ukdd.de
Github: https://github.com/smehralivand
Twitter: @smehralivand
Date: 11/16/2024
----------------------------------------------------------------------------------------------------------------------------------------
'''

from pathlib import Path
from tqdm import tqdm
import pandas as pd

def tb_txt_parser(txt_path):
    # Initiliaze variables
    context_col = []
    recommend_col = []
    txt_path = Path(txt_path)
    txt_files = list(txt_path.glob('*.txt'))

    # Iterate through txt files and seperate context from recommendation and append according list variables
    for txt_file in tqdm(txt_files):
        with open(txt_file, 'r', encoding='utf-8') as file:
            output = file.read()
            context = output.partition('Therapieempfehlung')[0]
            context_col.append(context)
            recommend = output.partition('Therapieempfehlung')[2]
            recommend = recommend.partition('Verantwortlich')[0]
            recommend_col.append(recommend)
    
    # Create dictionary of lists and pandas dataframe
    output_dict = {'context': context_col, 'recommendation': recommend_col}

    # Save output dictionary in pandas dataframe
    output_df = pd.DataFrame(output_dict)

    # Save dataframe as csv file
    output_df.to_csv('pre_tb.csv', sep=';', encoding='utf-8', index=False, header=True)
    
