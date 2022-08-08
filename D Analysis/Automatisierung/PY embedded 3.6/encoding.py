# -*- coding: latin-1 -*-
"""
Created on Mon Apr  9 21:58:44 2018

@author: DEMREI01
"""

import pandas as pd

BSEG= r'BSEG.csv'


dfBSEG = pd.read_csv(BSEG, delimiter='¬', engine='python')


print(dfBSEG)

del dfBSEG


