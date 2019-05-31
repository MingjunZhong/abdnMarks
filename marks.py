import sys
import matplotlib.pyplot as plt
import os
import numpy as np
import pandas
import statsmodels.api as sm
from statsmodels.graphics.regressionplots import abline_plot
import seaborn


def convertI(i):
  return str(i)

metadatafile=sys.argv[1]

#start by reading metadata file. Format for each row is filename, start line, number of students. There also needs to be a header line (first line) with text in first 4 cells of row saying "Filename" "Start row" "Number of students" and "Marks column". The last cell should say "Short name" and encodes the abbreviated name for the course which will appear on the graphs.
print("Reading metadata file ",metadatafile)
md=pandas.read_excel(metadatafile,sheet_name=0,header=0,index_col=0)

df=pandas.DataFrame()
for index,row in md.iterrows():
    print("Parsing ",index)    
    p=pandas.read_excel(index,sheet_name=0,header=None,index_col=0,skiprows=row["Start row"]-1,nrows=row["Number of students"],na_values=["MC","GC","NP"],converters={0:convertI},usecols="A:"+row["Marks column"])
    r=p.iloc[:,-1]
    #df=pandas.merge(df,r.rename(index.replace(".xlsx","")).to_frame(),how='outer',left_index=True,right_index=True)
    df=pandas.merge(df,r.rename(row["Short name"]).to_frame(),how='outer',left_index=True,right_index=True)


plotDf=pandas.DataFrame()
plotDf['total']=df.sum(axis=1)

for c in df.columns:
        plotDf[c+"y"]=(plotDf['total']-df[c])/(df.count(axis=1)-1)
        plotDf[c]=df[c]

plotDf=plotDf.replace(to_replace=[0,np.inf,-np.inf],value=None)        

for c in df.columns:
        ax=plotDf.plot.scatter(x=c,y=c+"y",c='Black')
        try: #handle case where LSF can't occur
          model=sm.OLS(plotDf[c+"y"],sm.add_constant(plotDf[c]),missing='drop')
          abline_plot(model_results=model.fit(),ax=ax,c='Red')
          abline_plot(intercept=0,slope=1,ax=ax,c='Blue')
        except:
          pass  

        ax.set(xlabel=c,ylabel="Average Mark")
        #ax.set_xlim(0,100)
        #ax.set_ylim(0,100)
        #for l in [40,50,60,70]: #draw lines for first etc boundaries
        ax.set_xlim(0,22) #use cgs marks instead
        ax.set_ylim(0,22)
        for l in [9,12,15,18]:
          ax.axhline(y=l,c='Blue')
          ax.axvline(x=l,c='Blue')
        plt.savefig(c+".pdf")

for c in df.columns:
  plt.clf()
  seaborn.set_context("notebook",font_scale=0.5)
  ax=seaborn.violinplot(y=df[c],cut=0,fontsize=8)
  ax.set_ylim(0,22)
  #seaborn.violinplot(data=df,cut=0,fontsize=8)
  plt.savefig(c+"ViolinPlot.pdf")
