import pandas
import xlrd
import win32ui
import math
def put():
    o = win32ui.CreateFileDialog(1, '', '', 0, 'Любой файл |*.*')
    o.DoModal()
    return o.GetPathName()
def transposition(sdf):
    i=0
    tt=0
    z=[]
    zzz=[]
    z1=[]
    z2=[]
    lensdf=len(sdf.columns)
    while i < lensdf:
        tt=tt+1
        z.append(sdf.iloc[0][i+1])
        zzz.append(sdf.iloc[1][i+1])
        z1.append(sdf.columns[i])
        z2.append(i)
        i=i+2
    i=0
    print(len(sdf))
    print(len(z))
    while i < len(z)-1:
        if (z[i]<z[i+1]) or ((z[i]==z[i+1]) and (zzz[i]<zzz[i+1])):
            zz=z[i]
            z[i]=z[i+1]
            z[i+1]=zz
            zz=zzz[i]
            zzz[i]=zzz[i+1]
            zzz[i+1]=zz
            zz=z1[i]
            z1[i]=z1[i+1]
            z1[i+1]=zz
            zz=z2[i]
            z2[i]=z2[i+1]
            z2[i+1]=zz
            i=i-2
        i = abs(i+1)
    i=0
    sdf2=[]
    sdf3=[]
    sdf3=pandas.DataFrame(sdf3)
    sdf2 = pandas.DataFrame(sdf2)
    while i< len(z)-1:
        sdf3=sdf[sdf.columns[z2[i]]]
        sdf2 = pandas.concat([sdf2, sdf3], axis=1)
        sdf3 = sdf[sdf.columns[z2[i]+1]]
        sdf2 = pandas.concat([sdf2, sdf3], axis=1)
        #print(sdf2)
        i=i+1
    return sdf2
# def sum1(sdf,df):
#     i=0
#     stromax=len(df)
#     stolmaxsdf=sdf.shape[1]
#     while i < stolmaxsdf:
#
def delcopy(sdf):
    i=1
    while i<len(sdf.columns):
        u=0
        while u<len(sdf.columns[i]):
            if sdf.iloc[u][i] == sdf.iloc[u][i + 2]:
                break
            if u==len(sdf.columns[i])-1:
                del sdf[sdf.columns[i]]
                del sdf[sdf.columns[i + 1]]
            u = u+1
        i=i+2
    return(sdf)
def consist(df, sdf):  # дополняем таблицу
    i = 1
    #while i < variant+1:
        #aa = df.columns[i]  # имя колонки по номеру
        #namestolb =
    #binn=bin(i+variant+1)  # split into binary code
        # u=3
        # sch=0
        # msch=4
        # fig=0
        # lenbin=len(binn)
        # while u<lenbin:
        #     if binn[u]=="1":
        #         sch+=1# involved column counter
        #     if sch==msch:
        #         fig=1
        #         break
        #     u=u+1
        # u=3
    maxx = len(bin(variant)) - 3
    sm = 0
    t = 1
    nid = 4  # number or grouping options
    n = 1
    met = 0
    y = 1
    fafa = 0
    u=0
    for kol in range(nid):
        a = [1 for y in range(kol + 1)]
        b = [0 for y in range(maxx - kol)]
        a.extend(b)
        sm = kol
        n = kol + 1
        while True:  # sm<maxx:
            if met == 0:
                print(a)
                u=0
                gr = []
                ns = ""
                sch = 0
                while u < maxx+1:  # !!!                !!!!         start of grouping # collecting names for grouping
                    if a[u] == 1:
                        gr.append(df.columns[u])  # collecting names for grouping
                        # ns = ns + df.columns[u] + ' + '  # usual namestolb
                        if ns:  # usual namestolb
                            ns = " + ".join([ns, df.columns[u]])
                        else:
                            ns = df.columns[u]
                        sch+=1
                    u += 1
                sn = pandas.DataFrame(df.groupby(gr).size().reset_index(name=('Итог' + str(i))))  # сама группировка
                sn[ns] = sn[sn.columns[:sch]].apply(lambda x: '_'.join(x.dropna().astype(str)), axis=1)  # складываем все кроме суммы
                sn = sn[[ns, ('Итог' + str(i))]]  # set and cut to stay 2 columns
                # sn = sn.loc[sn['штук'] != 1]                                         #clear from value
                stromax = len(df)
                howmach = sn[('Итог' + str(i))].sum()  # the amount of column total to count the rest
                sn = sn.sort_values(sn.columns[1], ascending=False)  # sorting
                sn.loc[len(sn)] = ['[Пустых]', stromax - howmach]  # adding value of empty row
                sn = sn.reset_index(drop=True)  # reset index for excel to save result of sorting
                i+=1
                if not sn.iloc[0][1] == 1:  # not len(sn.index)==0:
                    sdf = pandas.concat([sdf, sn], axis=1)  # the end of the grouping of the next combination of columns
            else:  # !!!                   !!!             !!!                     lower if the rearrangement itself
                met = 0
            if sm > maxx - 1:
                break
            if a[sm + 1] == 0:
                if n == 1:
                    a[sm] = 0
                    a[sm + 1] = 1
                    sm += 1
                else:
                    a[sm + 1] = 1
                    while sm > 0:
                        a[sm] = 0
                        sm = sm - 1
                    for t in range(0, n - 1):
                        a[t] = 1
                        t += 1
                    sm = 0
                    n = 1
            else:
                sm += 1
                n += 1
                met = 1
    return sdf
def otrezki(df, KOL_OTREZKOV, NumberStrok, COLUMN_NAME):
  df2 = []
  max = 0
  min = 9999999999999999999999
  myValues = df[COLUMN_NAME]
  for i in range(NumberStrok, len(df.index)):
      try:
          if min > myValues[i]:
              min = myValues[i]
          if max < myValues[i]:
              max = myValues[i]
      except Exception:
          return df
  myShag = (max - min) / KOL_OTREZKOV
  #print('min=', min, " ", "max=", max)
  #print("shag=", myShag)
  for i in range(NumberStrok, len(df.index)):
      try:
          nizGran = min + int((myValues[i] - min) / myShag) * myShag
          vehGran = nizGran + myShag
          if vehGran > max:
              nizGran = nizGran - myShag
              vehGran = vehGran - myShag
          k = str(nizGran) + "-" + str(vehGran)
          df2.append(k)
          df.loc[i, COLUMN_NAME] = k
      except Exception:
          df2.append(myValues[i])
  return df2
gr = []  # the basis for filling names of columns
ns = ""
tt=0
i=0
ii=0
filef=put()
sdf2=[]
# "C:\\иван\\кластер\\пит\\90.xlsx"
df = pandas.read_excel(filef)
stolmax = df.shape[1]-1                      # количество столбцов
variant = 2 ** (stolmax+1) - 1  # число всех перестановок и включений начиная с первой
KOL_OTREZKOV=5                               # задаем размерность арифметической кластеризации
print('process begin')
# for u in range(1, stolmax):
#     COLUMN_NAME=df.columns[u]                   # какую колонку смотрим
#     NumberStrok = 1                         # начиная с какой строки начинаем смотреть для числовой разбивки на кластеры
#     df2=pandas.DataFrame(otrezki(df, KOL_OTREZKOV, NumberStrok, COLUMN_NAME))  # функция кластеризации
sdf=pandas.DataFrame()
sdf=consist(df, sdf)
print('group finished')
sdf=transposition(sdf)
print('trnspositions finished')
#sdf=delcopy(sdf)
#sum1(sdf,df)
#sdf.to_excel(r'File_Name.xlsx')
sdf.to_csv(r'File Name.csv', sep='\t', encoding='cp1251')
print('save finished')
