import os
import pandas as pd
from pprint import pprint

curpath = os.path.dirname(__file__)
fn = os.path.relpath('docs\\master_mpn_list.xlsx', curpath)
df = pd.read_excel(fn, converters={'MPN': lambda x: str(x)}, index=False)
mpn_dict = df.set_index('MPN').to_dict()['Description']

pprint(mpn_dict)