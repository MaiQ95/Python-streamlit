import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

def statistics():

#Pobierz numery tygodni z excela "Dane.xlsx"
    dfWeek = pd.read_excel(
    io='ExcelData\Data.xlsx',
    engine='openpyxl',
    sheet_name='WeekNumber',
    usecols=['Date','Week']
    )

    dfWeek = dfWeek.astype({
    'Week':'category'})

#Pobierz dane ze statystyk -> FirstLane
    dfFirstLane = pd.read_excel(
    io='ExcelData\Statistics.xlsm',
    engine='openpyxl',
    sheet_name='FirstLane',
    usecols=['Date','Month','Shift','TypeClient','Client','ProductId',
    'Param1','Param2','Param3','Param4'],
    )
    dfFirstLane['Lane'] = 'FirstLane'
    dfFirstLane = dfFirstLane.merge(dfWeek, on='Date',how="left")

    dfFirstLane = dfFirstLane[dfFirstLane['Month']!='nan']
    dfFirstLane = dfFirstLane[dfFirstLane['Month'].isin([1,2,3,4,5,6,7,8,9,10,11,12])]
    dfFirstLane = dfFirstLane.astype({
    'Date':'datetime64',
    'Month':'int64',
    'Shift':'category',
    'TypeClient':'category',
    'Client':'str',
    'ProductId':'int64',
    'Param1':'int64',
    'Param2':'int64',
    'Param3':'float64',
    'Param4':'str'
    })

#Pobierz dane ze statystyk -> SecondLine
    dfSecondLine = pd.read_excel(
    io='ExcelData\Statistics.xlsm',
    engine='openpyxl',
    sheet_name='SecondLine',
    usecols=['Date','Month','Shift','TypeClient','Client','ProductId',
    'Param1','Param2','Param3','Param4'],
    )
    dfSecondLine['Lane'] = 'SecondLine'
    dfSecondLine = dfSecondLine.merge(dfWeek, on='Date',how="left")

    dfSecondLine = dfSecondLine[dfSecondLine['Month']!='nan']
    dfSecondLine = dfSecondLine[dfSecondLine['Month'].isin([1,2,3,4,5,6,7,8,9,10,11,12])]
    dfSecondLine = dfSecondLine.astype({
    'Date':'datetime64',
    'Month':'int64',
    'Shift':'category',
    'TypeClient':'category',
    'Client':'str',
    'ProductId':'int64',
    'Param1':'int64',
    'Param2':'int64',
    'Param3':'float64',
    'Param4':'str'
    })

#Pobierz dane ze statystyk -> ThirdLine
    ThirdLine = pd.read_excel(
    io='ExcelData\Statistics.xlsm',
    engine='openpyxl',
    sheet_name='ThirdLine',
    usecols=['Date','Month','Shift','TypeClient','Client','ProductId',
    'Param1','Param2','Param3','Param4'],
    )
    ThirdLine['Lane'] = 'ThirdLine'
    ThirdLine = ThirdLine.merge(dfWeek, on='Date',how="left")

    ThirdLine = ThirdLine[ThirdLine['Month']!='nan']
    ThirdLine = ThirdLine[ThirdLine['Month'].isin([1,2,3,4,5,6,7,8,9,10,11,12])]
    ThirdLine = ThirdLine.astype({
    'Date':'datetime64',
    'Month':'int64',
    'Shift':'category',
    'TypeClient':'category',
    'Client':'str',
    'ProductId':'int64',
    'Param1':'int64',
    'Param2':'int64',
    'Param3':'float64',
    'Param4':'str'
    })

    dfAll = pd.concat([dfFirstLane,dfSecondLine,ThirdLine], ignore_index=True)


    CAll_1,CAll_2 = st.columns(2)

    with CAll_1:

#Wybierz unikalne linie
        lane = dfAll["Lane"].unique()
        lane_all = st.multiselect(
        "Lane :",
        options=list(lane),
        default=list(lane)
        )

#Wybierz unikalne zmiany
        shift = dfAll["Shift"].unique()
        shift_all = st.multiselect(
        "Shifts :",
        options=list(shift),
        default=list(shift)
        )

    with CAll_2:


        month_all = st.multiselect(
        "Month :",
        options=dfAll["Month"].unique(),
        default=dfAll["Month"].unique()
        )

        typeClient = dfAll["TypeClient"].unique()
        typeClient_all = st.multiselect(
        "Client Type :",
        options=list(typeClient),
        default=list(typeClient)
        )

# Pobierz raport z rozszerzeniem xlsx
    df_selection = dfAll.query(
    "Lane == @lane_all & Month == @month_all & Shift == @shift_all & TypeClient == @typeClient_all")
    st.dataframe(df_selection)
    if st.button("Download Full Report xlsx"):
        df_selection.to_excel("statistics.xlsx")
