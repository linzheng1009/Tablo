import streamlit as st
import pandas as pd
from collections import Counter
import xlsxwriter
from io import BytesIO



st.set_page_config(page_title="Tablo",
                   page_icon=":crystal_ball",
                   layout="wide")

showWarningOnDirectExecution = False

# ---------- Header Section ---------- #

with st.container():
    st.title("Welcome to Tablo :wave:")
    st.write("__Tablo__ is a cross tabulation generator with add-in charts on Excel that greatly reduces your administrative burden.")
    st.write("---")



# -------------- STEP 1 -------------- #

with st.container():
    st.write("__STEP 1__: Upload survey responses data (csv/xlsx).")
    df = st.file_uploader("Please ensure the data are cleaned and weighted (if need be) prior to uploading.")
    if df:
        df_name = df.name
        if df_name[-3:] == 'csv':
            df = pd.read_csv(df, na_filter = False)
        else:
            df = pd.read_excel(df, na_filter = False)
    
    st.write("---")



# -------------- STEP 2 -------------- #

with st.container():
    st.write("__STEP 2__: Specify your __weighted__ and __demographic__ columns.")
    weight = st.selectbox('Specify your __weighted__ column', [''] + list(df.columns))
    demo = st.selectbox('Specify your __demographic__ column', [''] + list(df.columns))

    st.write("---")



# -------------- STEP 3 -------------- #

with st.container():
    st.write("__STEP 3__: Specify your __first__ and __last__ columns.")
    st.write("(NOTE: Place the __demographic__ columns after the survey questions.)")
    start = st.selectbox('Select the __first__ column of the dataset', [''] + list(df.columns))
    end = st.selectbox('Select the __last__ column of the dataset', [''] + list(df.columns))
    start_idx = df.columns.get_loc(start)
    end_idx = df.columns.get_loc(end)
    st.write("---")



# -------------- STEP 4 -------------- #

# Select by row, column, or both
with st.container():
    st.write("__STEP 4__: Select how the figures in the crosstabs should be presented.")
    values = st.selectbox('Show value as:', [''] + ["% of Column Total", "% of Row Total", "Both"])
    
    if (values == "% of Column Total"):
        pctof = "columns"
    elif (values == "% of Row Total"):
        pctof = "index"
    else:
        pctof = "all"



# ------------ PROCESSING ------------ #

# get the first and last survey question to form a range for the codes to run their analyses
# extract the question numbers infront of "."
col_list = [i[:i.index(".")] for i in df.columns[start_idx:end_idx+1]]

# create unique worksheet names 
sheet_name = []
for i, j in Counter(col_list).items():
    if j == 1:
        sheet_name.append(i)
    else:
        x = range(1,j+1)
        for y in x:
            z = f"{i}.{y}"
            sheet_name.append(z)

# create crosstabs
output = BytesIO()
with pd.ExcelWriter(output,
                    engine='xlsxwriter',
                    options= {'strings_to_numbers': True}) as writer:

    for j, i in enumerate(df.columns[start_idx:end_idx+1]):
        a = pd.crosstab(df[i],
                        df[demo],
                        values=df[weight],
                        aggfunc='sum',
                        normalize=pctof,
                        margins=True,
                        margins_name="Grand Total").applymap(lambda x: "{:.2f}".format(100*x)
                        ).to_excel(writer,sheet_name=sheet_name[j],startrow=0 , startcol=0)
    
        workbook = xlsxwriter.Workbook('output.xlsx')
    writer.save()

    st.write('__Number of crosstabs__: ' + str(end_idx - start_idx + 1))

    df_xlsx = output.getvalue()
    df_name = df_name[:df_name.find('.')]
    st.subheader('Your crosstabs have been successfully invoked :sparkles:')
    st.download_button(label='📥 Download Crosstabs Only', data=df_xlsx, file_name= df_name + f'_{demo}' + f'_{pctof}' + ' _crosstabs.xlsx')

    st.write("---")



# -------------- STEP 5 -------------- #

with st.container():
    st.write("__STEP 5__: Visualise your crosstabs.")
    dfcharts = st.file_uploader("Please upload the downloaded xlsx file from __STEP 4__.")
    if dfcharts:
        dfcharts_name = dfcharts.name
        if dfcharts_name[-3:] == 'csv':
            # "sheet_name=None" is used here to turn our dataset into an Ordered Dictionary
            dfcharts = pd.read_csv(dfcharts, na_filter = False, sheet_name=None)
        else:
            dfcharts = pd.read_excel(dfcharts, na_filter = False, sheet_name=None)


# Create charts of crosstabs accross every worksheets
output = BytesIO()
df_sheets = list(dfcharts.keys()) 
workbook = xlsxwriter.Workbook(output,{'in_memory': True})
string = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

for idx, i in enumerate(df_sheets[0:]):
    worksheet = workbook.add_worksheet(i)
    df = dfcharts.get(i)
    length = df.shape[0]
    for j, col in enumerate(df.columns):
        worksheet.write_row('A1', df.columns)
        worksheet.write_column(f"{string[j]}2", df[col])
        chart = workbook.add_chart({'type':'bar'})
        for k in range(1, df.shape[1]-1):
            chart.set_style(11)
            chart.add_series({
                'name':       f'={i}!${string[k]}$1',                        
                'categories': f'={i}!$A$2:$A${length+1}',                    
                'values':     f'={i}!${string[k]}$2:${string[k]}${length+1}' 
                })
            worksheet.insert_chart('O2', chart, {'x_offset': 25, 'y_offset': 10})
workbook.close() 

dfcharts_xlsx = output.getvalue()
dfcharts_name = dfcharts_name[:dfcharts_name.find('.')]
st.subheader('Your crosstabs are now imbued with charts :bar_chart:')
st.download_button(
    label='📥 Download Crosstabs with Charts', 
    data=dfcharts_xlsx, 
    file_name= dfcharts_name + '_charts.xlsx', 
    mime="application/vnd.ms-excel"
    )





