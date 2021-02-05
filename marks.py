import pandas as pd
from numbers import Number
import math
import sys
import seaborn as sns
from matplotlib import pyplot as plt
import statsmodels.api as sm
from statsmodels.graphics.regressionplots import abline_plot
import os, fnmatch


#### ASSUMPTIONS
## COURSE NAME LOCATION IS IN POSITION 1,0 (row 2, column 1)
## The column marked "Comments" has the last "Number" label in it
## Sudents start on the row after the "Comments" header
## And end before the word "Number" (which does not appear in comments)
## The last number colum in the same row as Number indicates the CGS mark column

COURSENAMELOCATION=(1,0)

def read_file(filename):
  # Note that to read .xlsx files using pandas, it requires `openpyxl' and also
  # need engine='openpyxl'.
  df=pd.read_excel(filename,engine='openpyxl',sheet_name=0,header=None,na_values=["MC","GC","NP"])

  #start by pulling out course name
  coursename=df.iloc[COURSENAMELOCATION].split(' ')[0]

  #now find the index of the column containing the word "Comment"
  s=df.isin(["Comment"]).any()
  col=s[s].index[0]

  #work out starting row
  startrow=df[df[col]=='Comment'].index[0]

  #find the row containing the word Number in that column
  numrow=df[df[col]=='Number'].index[0]

  markcol=-1
  for i in range(col-1,0,-1): #-1 as we start a column before
      if isinstance(df.loc[numrow,i],Number) and not math.isnan(df.loc[numrow,i]):
          markcol=i
          break
  if markcol==-1:
      return
  
  #so now we have the column for marks (markcol), the start row for students (startrow and approximate end row: numrow. We need to get studentid and mark to a dataframe
  retdf=df.iloc[startrow+1:numrow][[0,markcol]]
  retdf.columns=['student_id',coursename]

  #clean up by dropping students with a NAN name or mark
  retdf.dropna(axis='index',inplace=True)

  retdf=retdf.astype({coursename:'float'})
  retdf.set_index('student_id',inplace=True)
  retdf.index=retdf.index.astype(int)
  return retdf

########################################################
def create_scatter(df,module):
  #to create a scatter plot, we find each student, get their average mark across all modules except this one and plot it against this one.
  df2=pd.DataFrame({"avg":(df.sum(axis=1)-df[module])/(df.count(axis=1)-1),module:df[module]})
  
  plt.clf()
  sns.set_context("notebook",font_scale=0.5)
  sns.set(style="white")
  ax=df2.plot.scatter(x=module,y="avg",c='Black')
  model=sm.OLS(df2["avg"],sm.add_constant(df2[module]),missing='drop')
  abline_plot(model_results=model.fit(),ax=ax,c="Red")
  abline_plot(intercept=0,slope=1,ax=ax,c="Blue")

  ax.set(xlabel=module,ylabel="Average Mark")
  ax.set_xlim(0,22)
  ax.set_ylim(0,22)
  for l in [9,12,15,18]:
    ax.axhline(y=l,c='Blue')
    ax.axvline(x=l,c='Blue')
  plt.savefig(module+"Scatter.png")

def create_violin(module):
  #module is the column
  plt.clf()
  sns.set_context("notebook",font_scale=0.5)
  sns.set(style="white")
  ax=sns.violinplot(y=module,cut=0,fontsize=8)
  ax.set_ylim(0,22)
  plt.savefig(module.name+"Violin.png")


########################################################

def read_files(filearray):
    df=pd.DataFrame({'student_id':[]})
    df=df.set_index("student_id")
    for i in filearray:
        print('parsing',i)
        df2=read_file(i)
        if df2 is None:
            continue
        df=pd.DataFrame.merge(df,df2,on='student_id',how='outer')
    return df

########################################################

# find all *.xlsx files 
def find_xlsx():
    listOfAllFiles = os.listdir('.')
    pattern = "*.xlsx"
    listFiles = []
    for file in listOfAllFiles:
        if fnmatch.fnmatch(file,pattern):
            listFiles.append(file)
    return listFiles

listFiles = find_xlsx()
df=read_files(listFiles)
#df=read_files(sys.argv[1:])

#now df contains students and courses, so we can just create the relevant violin plots and scatter plots  
for m in df.columns:
   try:
     create_violin(df[:][m])
     create_scatter(df,m)
   except:
     print("unable to generate plot for ",m)


