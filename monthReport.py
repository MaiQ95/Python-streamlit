import streamlit as st
import pandas as pd
from PIL import Image
from openpyxl import Workbook, load_workbook
import numpy as np
import plotly.express as px


def monthReport():
    dfFirstLane = pd.read_excel(
         io='ExcelData\Statistics.xlsm',
         sheet_name='FirstLane',
         usecols=['Date','Month','Shift','TypeClient','Client','ProductId',
         'Param1','Param2','Param3','Param4'],
         )
    dfFirstLane = dfFirstLane.squeeze()
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

    #Count new parameter
    dfFirstLane['Calc1'] = dfFirstLane['Param1']-dfFirstLane['Param2']-dfFirstLane['Param3']

    shifts = dfFirstLane.groupby(by=['Month'])


    report = shifts.agg({"Param1":"sum",
    "Param1":"sum",
    "Param2":"sum",
    "Param3":"sum",
    "Date":"count",
    "Calc1":"sum"
    })

    dfBudget = pd.read_excel(
    io='ExcelData\Data.xlsx',
    engine='openpyxl',
    sheet_name='Budget',
    usecols=['Month','FirstLane']
    )

    dfBudget = dfBudget.squeeze()

    dfShiftCalc = pd.read_excel(
    io='ExcelData\Data.xlsx',
    engine='openpyxl',
    sheet_name='ShiftCalc',
    usecols=['Month','FirstLane']
    )

    dfShiftCalc = dfShiftCalc.squeeze()

    dfShiftSystem = pd.read_excel(
    io='ExcelData\Data.xlsx',
    engine='openpyxl',
    sheet_name='ShiftSystem',
    usecols=['Month','FirstLane']
    )

    dfShiftSystem = dfShiftSystem.squeeze()

    report.rename(columns={'Date':'Products'},inplace=True)

    report = report.merge(dfBudget, on="Month",how="left")
    report.rename(columns={'FirstLane':'Budget'},inplace=True)

    report = report.merge(dfShiftCalc, on="Month",how="left")
    report.rename(columns={'FirstLane':'Max shifts'},inplace=True)

    report = report.merge(dfShiftSystem, on="Month",how="left")
    report.rename(columns={'FirstLane':'Shift system'},inplace=True)


    dfShiftsMonth = pd.read_excel(
         io='ExcelData\Statistics.xlsm',
         sheet_name='ShiftsMonth',
         usecols=['Month','FirstLane'])

    report = report.merge(dfShiftsMonth, on="Month",how="left")
    report.rename(columns={'FirstLane':'Shifts worked'},inplace=True)

    report['production [Param1]/[h]'] = report['Param1']/report['Shifts worked']
    report['production [pcs]/[h] full time'] = report['Param1']/report['Max shifts']

    report['production [pcs]/[h]'] = report['Products']/report['Shifts worked']
    report['production [pcs]/[h] full time'] = report['Products']/report['Max shifts']

    report['Order realization'] = report['Param1']/report['Budget']

    report['Average Shift number'] = report['Shifts worked']/report['Max shifts']*report['Shift system']


    report = report.transpose()
    st.dataframe(report)

    if st.button("Download month report xlsx"):
        report.to_excel('monthReport.xlsx')
