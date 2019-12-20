import matplotlib.pyplot as plt
import pandas as pd
import matplotlib.dates as mdates
import pyodbc
from pandas.plotting import register_matplotlib_converters
register_matplotlib_converters()

import matplotlib.ticker as plticker


conn = pyodbc.connect('Driver={SQL Server Native Client 11.0};'
                            'Server=RES_ATELIER_1\SERVER4AUTO;'
                            'Database=DB4Auto;'
                            'UID=sa;'
                            'PWD=aaaaaa;'
                            'Trusted_Connection=No;')

#date = pd.read_sql_query(
#'''SELECT CAST(DATEADD(HOUR, -6, GETDATE()) AS DATE) AS currentDate, CAST(DATEADD(HOUR, 18, GETDATE()) AS DATE) AS nextDate, GETDATE() AS currentTimestamp''', conn)

#SQL_Query = pd.read_sql_query(
#'''EXEC [dbo].[QAGATE_1_Event_Jour]''', conn)

#df = pd.DataFrame(SQL_Query, columns=['currentOF','code','etat','timeStamp'])
#pd.set_option('display.max_rows', df.shape[0]+1)

#df1 = pd.DataFrame([[None, None, df[-1:]['etat'].to_string(index=False), date['currentTimestamp'].to_string(index=False)]], columns=['currentOF','code','etat','timeStamp'])

#df = df.append(df1)

#fig, ax = plt.subplots(1)
#fig.autofmt_xdate()
#plt.step(df['timeStamp'], df['etat'], df['timeStamp'], df['etat'].astype(float) - 0.2, where='post')

#loc = plticker.MultipleLocator(base=1.0) # this locator puts ticks at regular intervals
#ax.yaxis.set_major_locator(loc)

#labels = [item.get_text() for item in ax.get_yticklabels()]
#labels[1] = 'Stop'
#labels[2] = 'Setup'
#labels[3] = 'Run'

#ax.set_yticklabels(labels)

#xfmt = mdates.DateFormatter('%H:%M')
#ax.xaxis.set_major_formatter(xfmt)
#ax.set_xlim(str(date['currentDate'].to_string(index=False)) + ' 06:00:00', str(date['nextDate'].to_string(index=False)) + ' 06:00:00')

#plt.show()





date = pd.read_sql_query(
'''SELECT GETDATE() AS currentTimestamp''', conn)

SQL_Query = pd.read_sql_query(
'''EXEC [dbo].[QAGATE_1_Event_OF]''', conn)

df = pd.DataFrame(SQL_Query, columns=['currentOF','code','etat','timeStamp'])
pd.set_option('display.max_rows', df.shape[0]+1)

df1 = pd.DataFrame([[None, None, df[-1:]['etat'].to_string(index=False), date['currentTimestamp'].to_string(index=False)]], columns=['currentOF','code','etat','timeStamp'])

df = df.append(df1)
print(df)

fig, ax = plt.subplots(1)
fig.autofmt_xdate()
plt.step(df['timeStamp'], df['etat'], where='post')

loc = plticker.MultipleLocator(base=1.0) # this locator puts ticks at regular intervals
ax.yaxis.set_major_locator(loc)

labels = [item.get_text() for item in ax.get_yticklabels()]
labels[1] = 'Stop'
labels[2] = 'Setup'
labels[3] = 'Run'

ax.set_yticklabels(labels)

xfmt = mdates.DateFormatter('%d-%m %H:%M')
ax.xaxis.set_major_formatter(xfmt)

df['timeStamp'] = pd.to_datetime(df['timeStamp'])
datemax = df['timeStamp'].max()
datemin = df['timeStamp'].min()

ax.set_xlim(datemin, datemax)

plt.show()