import pandas as pd
df1 = pd.DataFrame([['a', 1], ['b', 2]],
                   columns=['letter', 'number'])
df3 = pd.DataFrame([['cat'], ['dog']],
                   columns=['animal'])

print(pd.concat([df1, df3], axis=1))

rowNumb = 1000
df = pd.DataFrame(columns = ['OF', 'Reference'])
for x in range(rowNumb):
    df = df.append({'OF' : '30543' , 'Reference' : '490035-2000'}, ignore_index=True )
print(df)