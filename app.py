import streamlit as st
import pandas as pd
from PIL import Image
from openpyxl import Workbook, load_workbook
import numpy as np
import plotly.express as px
from statistics import statistics
from monthReport import monthReport


img = Image.open("Logo/Logo.PNG")
st.set_page_config(page_title='Production',
page_icon=img,
layout="wide"
)


def main():
    menu = (["Home site","Statistics","Reports"])
    choice = st.sidebar.selectbox("Menu",menu)


    if choice == "Statistics":
        st.title("Statistics")
        statistics()



    elif choice == "Reports":
        st.title("Reports")
        with st.expander("Month Report"):
            monthReport()

    else:
        st.title("Production - Home Site")


if __name__ == '__main__':
    main()
