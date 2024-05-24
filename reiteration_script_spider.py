# -*- coding: utf-8 -*-

# importation des fichiers dans un dossier
import os
chemin_dossier = 'C:\\Users\\sst\\Downloads\\REITERATION\\source'
fichiers = os.listdir(chemin_dossier)
nom_du_fichier = "resultat"
intervalle = '60T'
n = 2

# lire tous les fichiers csv
import pandas as pd
fichier_csv = [fichier for fichier in fichiers if fichier.endswith('.csv')]
dfs = []
for fichier_csv in fichier_csv :
    chemin_fichier = os.path.join(chemin_dossier,fichier_csv)
    df = pd.read_csv(chemin_fichier,delimiter=";")
    dfs.append(df)
data_final = pd.concat(dfs,ignore_index=True)

def converti_date(data_final) :
    data_final['date'] = pd.to_datetime(data_final['date_appel'], format='%Y-%m-%d %H:%M:%S').dt.date
    return data_final
data_final = converti_date(data_final)

data_final_tranche = data_final.copy()
liste_date = list(data_final_tranche['date'].unique())
data_jour = []
for date in liste_date :
    data_jour.append(data_final_tranche[data_final_tranche['date']==date])

# calcul de la réitération par heure
tranche = data_final.copy()
def transformation(tranche) :
    tranche['date_appel'] = pd.to_datetime(tranche['date_appel'])
    tranche.set_index('date_appel')
    hourly_counts = tranche.groupby([pd.Grouper(key='date_appel', freq='D'),
                                pd.Grouper(key='date_appel', freq=intervalle)])['appelant'].count()
    hourly_unique = tranche.groupby([pd.Grouper(key='date_appel', freq='D'),
                                pd.Grouper(key='date_appel', freq=intervalle)])['appelant'].nunique()
    programme = tranche.groupby([pd.Grouper(key='date_appel', freq='D'),
                                pd.Grouper(key='date_appel', freq=intervalle)])['Programme'].nunique()
    
    def transform(hourly_counts,name) :
        hourly_counts.index.names = ['Date', 'Hour']
        hourly_counts = pd.DataFrame(hourly_counts)
        hourly_counts.columns = [name]
        hourly_counts = hourly_counts.reset_index()
        hourly_counts['Hour'] = hourly_counts['Hour'].dt.strftime('%H:%M:%S')
        hourly_counts = hourly_counts.set_index(['Date', 'Hour'])
        return hourly_counts

    hourly_counts = transform(hourly_counts,'Calls')
    hourly_unique = transform(hourly_unique,'Callers')
    programme = transform(programme,'Called')
    data_tranche = pd.concat([hourly_counts,hourly_unique],axis=1)
    data_tranche['Reiteration'] = round(100*((data_tranche['Calls']-data_tranche['Callers'])/data_tranche['Calls']),2)
    data_tranche = pd.concat([data_tranche,programme],axis=1)
    return data_tranche
data_tranche = transformation(tranche)

data_final = converti_date(data_final)
list_programmes = list(data_final['Programme'].unique())
list_dataframes = []
for i in range(len(list_programmes)) :
    df_prog = data_final[data_final['Programme']==list_programmes[i]].copy()
    list_dataframes.append(df_prog)

# triage des données par le nombre de jour de réitération choisi
def triage(data_final) :
    pd.options.mode.copy_on_write = True
    reiteration = data_final['date'].unique()
    reiteration = pd.Series(reiteration)
    callers = []
    for i in range(len(reiteration)) :
        jours_glissant = reiteration[i:i+n]
        if len(jours_glissant) < n :
            break
        jours_glissant = list(jours_glissant)
        callers.append(jours_glissant)
    n_callers = []
    for i in range(len(callers)) :
        n_callers.append(data_final[data_final['date'].isin(callers[i])])
    n_unique = [caller[-1] for caller in callers]
    for caller, date in zip(n_callers, n_unique):
        caller['date'] = date
    n_callers = pd.concat(n_callers)
    reiteration = n_callers.set_index('date')
    return reiteration

top_calls = triage(data_final)
top_calls = top_calls.groupby(level=0)['appelant'].value_counts()
top_calls = pd.DataFrame(top_calls)
top_calls = top_calls.rename(columns={'appelant':'count'}) 
top_calls = top_calls.reset_index()
top_calls = top_calls[top_calls['count']>=2]
top_calls = top_calls.sort_values(by=['count'],ascending=False)
top_calls = top_calls.rename(columns={'appelant':'callers'})
top_unique = top_calls['callers'].unique()
top_callers = data_final[data_final['appelant'].isin(top_unique)]
top_callers = top_callers.sort_values(by=['appelant','date_appel'],ascending=True).drop(columns=['date'])
top_callers = top_callers.rename(columns={'date_appel':'date','appelant':'Calling','Programme':'Called'})

reiteration = triage(data_final)
reiteration_par_programme = []
for i in range(len(list_dataframes)) :
    reit_p = reiteration[reiteration['Programme']==list_programmes[i]].copy()
    reiteration_par_programme.append(reit_p)

def resultat(reiteration) :
    temp = reiteration.drop(columns=['Programme','date_appel','appelant'])
    temp = temp.groupby(level=0).value_counts()
    temp = pd.DataFrame(temp)
    temp = temp.rename(columns={0:'Calls'})
    reit = reiteration.drop(columns=['Programme','date_appel'])
    reit = reit.groupby(level=0).nunique()
    Reiteration = pd.concat([temp,reit],axis=1)
    Reiteration = Reiteration.rename(columns={'appelant':'Callers'})
    Reiteration['Reiteration'] = round(100*((Reiteration['Calls']-Reiteration['Callers'])/Reiteration['Calls']),2)
    return Reiteration

globales = resultat(reiteration)
resultat_par_programme = list_dataframes.copy()
for i in range(len(list_dataframes)) :
    resultat_par_programme[i] = resultat(reiteration_par_programme[i])
    resultat_par_programme[i]['Called'] = list_programmes[i]

# exporation du résultat dans un fichier excel
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
data_finals = data_final.drop(columns=['date'])
with pd.ExcelWriter("C:\\Users\\sst\\Downloads\\REITERATION\\"+nom_du_fichier+".xlsx") as writer :
    data_finals.to_excel(writer, sheet_name='Data',index=False)
    data_tranche.to_excel(writer,sheet_name='SummaryH',index=True)
    globales.to_excel(writer,sheet_name='SummaryD',index=True)
    top_calls.to_excel(writer,sheet_name='top_calls',index=False)
    top_callers.to_excel(writer,sheet_name='top_calls_Details',index=False)
    for i in range(len(resultat_par_programme)) :
        resultat_par_programme[i].to_excel(writer,sheet_name=list_programmes[i],index=True)















