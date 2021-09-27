import os
import pandas as pd

files = os.listdir('..\Input')
print(files)
files_xls = [f for f in files if f[-4:] == 'xlsx']
print(files_xls)