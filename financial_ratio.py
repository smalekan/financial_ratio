import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from statistics import stdev
from statistics import mean
import seaborn as sns
from __future__ import unicode_literals
import arabic_reshaper
from bidi.algorithm import get_display
from matplotlib.backends.backend_pdf import PdfPages

def search(company_name,year):
  xls_file = pd.ExcelFile('financial_ratio.xlsx')
  df = xls_file.parse('Sheet1')
  df = df.replace('-', np.nan)
  dff = df
  df = df.rename(columns={'سال': 'year'})
  df = df.rename(columns={'نام شرکت': 'company'})
  df = df.rename(columns={'id شرکت': 'comp_id'})
  
  df_96 = df.loc[df['year'] == 96]
  df_97 = df.loc[df['year'] == 97]

  mean_96 = df_96.mean(axis = 0,skipna = True)
  med_96 = df_96.median(axis = 0,skipna = True)

  mean_97 = df_97.mean(axis = 0,skipna = True)
  med_97 = df_97.median(axis = 0,skipna = True)

  year=int(year)
  select = df.loc[(df.company == company_name) & (df.year == year)]

  if(select.shape[0]==1):
    if(year==96):
      select.loc[1] = mean_96
      select.set_value(1, 'company', 'میانگین')
      select.set_value(1, 'comp_id',0 )
      select.set_value(1, 'year', 0)
      
      select.loc[2] = med_96
      select.set_value(2, 'company', 'میانه')
      select.set_value(2, 'comp_id',0 )
      select.set_value(2, 'year', 0)      
    if(year==97):
      select.loc[1] = mean_97
      select.set_value(1, 'company', 'میانگین')
      select.set_value(1, 'comp_id',0 )
      select.set_value(1, 'year', 0)
      
      select.loc[2] = med_97
      select.set_value(2, 'company', 'میانه')
      select.set_value(2, 'comp_id',0 )
      select.set_value(2, 'year', 0)

  kwargs = dict(hist_kws={'alpha':.7}, kde_kws={'linewidth':3})
  df = dff.dropna()
  pp = PdfPages('ملی نفت 96.pdf')
  for i in range(3, 12):
    
    z = df.iloc[:,i]
    # print(select.iloc[0,i]) 
    # plt.suptitle('test title')
    v = z.name
    z = z.values.tolist()
    pn = plt.figure(figsize=(20,7), dpi= 200)
    # plt.xlabel.encode('ascii', 'ignore')
    # plt.axes.get_xaxis().set_visible(False)
    # plt.axis.remove(self)
    sn = sns.distplot(z, color='#E69F00', label="Compact", **kwargs)
    tit = get_display( arabic_reshaper.reshape(v))
    plt.title(tit)
    for k in range(len(sn.patches)):
      #print(sn.patches[k].get_x())
      #print(sn.patches[k].get_x()+sn.patches[k].get_width())
      if sn.patches[k].get_x()<=select.iloc[0,i]<=sn.patches[k].get_x()+sn.patches[k].get_width():
        sn.patches[k].set_color('r')
    pp.savefig(pn)   
    plt.show()
    #plt.savefig("foo.pdf", bbox_inches='tight')
  pp.close()
  return select

y = search('ملی نفت','96')
y.head()
y.to_excel('ملی نفت 96.xlsx')