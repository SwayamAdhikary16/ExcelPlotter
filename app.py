import streamlit as st
import pandas as pd
import plotly.express as px
import base64
from io import StringIO, BytesIO

def generate_excel_download_link(df):
    towrite = BytesIO()
    df.to_excel(towrite, encoding='utf-8', index=False, header=True)
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="data.xlsx">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_html_download_link(fig):
    towrite = StringIO()
    fig.write_html(towrite, include_plotlyjs="cdn")
    towrite = BytesIO(towrite.getvalue().encode())
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:text/html;charset=utf-8; base64, {b64}" download="plot.html"> Download Graph </a>'
    return st.markdown(href, unsafe_allow_html=True)

st.set_page_config(page_title='Excel Plotter')
st.title('Excel Plotter 📊')
st.subheader('Input your Excel File')

uploaded_file = st.file_uploader('Choose a XLSX file', type='xlsx')

if uploaded_file:
    st.markdown('---')

    df = pd.read_excel(uploaded_file, engine='openpyxl')
    st.dataframe(df)

    data_top = []

    for col in df.columns:
        data_top.append(str(col))
    
    groupby_column = st.selectbox(
        'Column',
        tuple(data_top),
    )

    #====GROUP DATAFRAME====

    output_columns = []
    my_type = ['int64', 'float64']
    dtypes = df.dtypes.to_dict()

    for col_name, typ in dtypes.items():
        if typ in my_type:
            output_columns.append(str(col_name))
    
    df_grouped = df.groupby(by=[groupby_column], as_index=False)[output_columns].sum()

    int_col = st.selectbox(
        'Data',
        tuple(output_columns),
    )

    st.dataframe(df_grouped)

    #====PLOT DATAFRAME====

    fig_type = st.selectbox(
        'Chart Type',
        ('Bar Chart', 'Line Chart', 'Pie Chart')
    )

    if fig_type == 'Bar Chart':
        fig = px.bar(
            df_grouped,
            x = groupby_column,
            y = int_col,
            color = str(int_col),
            color_continuous_scale=['red', 'yellow', 'green'],
            template = 'plotly_white',
            title=f'<b>{int_col} by {groupby_column}</b>'
        )
    elif fig_type == 'Line Chart':
        fig = px.line(
            data_frame=df_grouped,
            x = groupby_column,
            y = int_col,
            title=f'<b>{int_col} by {groupby_column}</b>'
        )
    elif fig_type == 'Pie Chart':
        fig = px.pie(
            df_grouped,
            names=groupby_column,
            values=int_col,
            title=f'<b>{int_col} by {groupby_column}</b>'
        )
        fig.update_traces(
            textposition='inside',
            textinfo='percent+label'
        )
        fig.update_layout(
            title_font_size=42
        )

    st.plotly_chart(fig)

    st.subheader('Downloads')
    generate_excel_download_link(df_grouped)
    generate_html_download_link(fig)
    