import pandas as pd 
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="Ratio analyse Reynders",
    page_icon="logo.png",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ---- Read Excel ----
@st.cache
def get_activa_from_excel():
    df = pd.read_excel(
        io="data/GIP_analyse van de jaarrekening_start (1).xlsx",
        engine="openpyxl",
        sheet_name="verticale analyse balans",
        usecols="A:E",
        nrows=100,
        header=2
    )

    # filter row on column value
    activa = ["VASTE ACTIVA","VLOTTENDE ACTIVA"]
    df = df[df['ACTIVA'].isin(activa)]

    return df

df_activa = get_activa_from_excel()
df_activa = df_activa.round({"Boekjaar 1":2, "Boekjaar 2":2, "Boekjaar 3":2})

#@st.cache
st.subheader("Liquiditeit")
def get_liq_from_excel():
    df = pd.read_excel(
        io="data/GIP_analyse van de jaarrekening_start (1).xlsx",
        engine="openpyxl",
        sheet_name="Liquiditeit",
        usecols="A:D",
        nrows=100,
        header=1
    )
    # change column names
    df.columns = ["Type","Boekjaar 1","Boekjaar 2","Boekjaar 3"]
    # filter row on column value
    liq = ["Liquiditeit in ruime zin","Liquiditeit in enge zin"]
    df = df[df["Type"].isin(liq)]

    df = df.T #Transponeren
    df = df.rename(index={"Boekjaar 1":"1","Boekjaar 2":"2",
                    "Boekjaar 3":"3"})
    df = df.iloc[1: , :] # Drop first row 
    df.insert(0,"Boekjaar",["Boekjaar 1","Boekjaar 2",
                    "Boekjaar 3"],True)
    df.columns = ["Boekjaar","Liquiditeit in ruime zin","Liquiditeit in enge zin"] # change column names
    
    
    return df


df_liq = get_liq_from_excel()

fig_liq = px.line(df_liq, x="Boekjaar", y=["Liquiditeit in ruime zin","Liquiditeit in enge zin"], markers=True)
fig_liq.update_layout({
'plot_bgcolor': 'rgba(0, 0, 0, 0)',
'paper_bgcolor': 'rgba(0, 0, 0, 0)',})

fig_liq.update_traces(line=dict(width=3))
st.plotly_chart(fig_liq, use_container_width=True)

st.subheader("Solvabiliteit")
def get_solv_from_excel():
    df = pd.read_excel(
        io="data/GIP_analyse van de jaarrekening_start (1).xlsx",
        engine="openpyxl",
        sheet_name="Solvabiliteit",
        usecols="A:D",
        nrows=100
    )
    # change column names
    df.columns = ["Type","Boekjaar 1","Boekjaar 2","Boekjaar 3"]
    # filter row on column value
    solv = ["EIGEN VERMOGEN","TOTAAL VAN DE PASSIVA","Solvabiliteit"]
    df = df[df["Type"].isin(solv)]
    
    df = df.T #Transponeren
    df = df.rename(index={"Boekjaar 1":"1","Boekjaar 2":"2",
                    "Boekjaar 3":"3"})
    df = df.iloc[1: , :] # Drop first row 
    df.insert(0,"Boekjaar",["Boekjaar 1","Boekjaar 2",
                    "Boekjaar 3"],True)
    df.columns = ["Boekjaar","Eigen vermogen","Totaal van de passiva","Solvabiliteit"] # change column names
    
    return df

df_solv = get_solv_from_excel()

fig_Solv = px.bar(df_solv,x="Boekjaar",y="Solvabiliteit",range_y=[0.8,0.85], text_auto=True)
fig_Solv.update_layout({
'plot_bgcolor': 'rgba(0, 0, 0, 0)',
'paper_bgcolor': 'rgba(0, 0, 0, 0)',})

st.plotly_chart(fig_Solv, use_container_width=True)

st.subheader("REV")
def get_Rev_from_excel():
    df = pd.read_excel(
        io="data/GIP_analyse van de jaarrekening_start (1).xlsx",
        engine="openpyxl",
        sheet_name="REV",
        usecols="A:D",
        nrows=100
    )
    # change column names
    df.columns = ["Type","Boekjaar 1","Boekjaar 2","Boekjaar 3"]
    # filter row on column value
    solv = ["Te bestemmen winst van het boekjaar","EIGEN VERMOGEN","REV"]
    df = df[df["Type"].isin(solv)]

    df = df.T #Transponeren
    df = df.rename(index={"Boekjaar 1":"1","Boekjaar 2":"2",
                    "Boekjaar 3":"3"})
    df = df.iloc[1: , :] # Drop first row 
    df.insert(0,"Boekjaar",["Boekjaar 1","Boekjaar 2",
                    "Boekjaar 3"],True)
    df.columns = ["Boekjaar","Te bestemmen winst van het boekjaar","EIGEN VERMOGEN","REV"] # change column names
    
    return df

df_Rev = get_Rev_from_excel()

fig_REV = px.line(df_Rev, x="Boekjaar", y="REV", markers=True)
fig_REV.update_layout({
'plot_bgcolor': 'rgba(0, 0, 0, 0)',
'paper_bgcolor': 'rgba(0, 0, 0, 0)',})

fig_REV.update_traces(line=dict(width=3))
st.plotly_chart(fig_REV, use_container_width=True)

st.subheader("Omlooptijd van de voorraad")
def get_Vrd_from_excel():
    df = pd.read_excel(
        io="data/GIP_analyse van de jaarrekening_start (1).xlsx",
        engine="openpyxl",
        sheet_name="Voorraad",
        usecols="A:D",
        nrows=100
    )
    # change column names
    df.columns = ["Type","Boekjaar 1","Boekjaar 2","Boekjaar 3"]
    # filter row on column value
    vrd = ["Handelsgoederen, grond- en hulpstoffen","Voorraden en bestellingen in uitvoering","Omlooptijd","Omloopsnelheid voorraden"]
    df = df[df["Type"].isin(vrd)]

    df = df.T #Transponeren
    df = df.rename(index={"Boekjaar 1":"1","Boekjaar 2":"2",
                    "Boekjaar 3":"3"})
    df = df.iloc[1: , :] # Drop first row 
    df.insert(0,"Boekjaar",["Boekjaar 1","Boekjaar 2",
                    "Boekjaar 3"],True)
    df.columns = ["Boekjaar","Omlooptijd","Omloopsnelheid voorraden","Voorraden en bestellingen in uitvoering","Handelsgoederen, grond- en hulpstoffen"] # change column names
    
    return df

df_Vrd = get_Vrd_from_excel()

fig_Vrd = px.line(df_Vrd, x="Boekjaar", y="Omlooptijd", markers=True)
fig_Vrd.update_layout({
'plot_bgcolor': 'rgba(0, 0, 0, 0)',
'paper_bgcolor': 'rgba(0, 0, 0, 0)',})

fig_Vrd.update_traces(line=dict(width=3))
st.plotly_chart(fig_Vrd, use_container_width=True)

st.subheader("Klant- en Leverancierskrediet")
def get_Klv_from_excel():
    df = pd.read_excel(
        io="data/GIP_analyse van de jaarrekening_start (1).xlsx",
        engine="openpyxl",
        sheet_name="KlantLevKrediet",
        usecols="A:D",
        nrows=100,
        header=1
    )
    # change column names
    df.columns = ["Type","Boekjaar 1","Boekjaar 2","Boekjaar 3"]
    # filter row on column value
    Klv = ["Klantenkrediet","Leverancierskrediet","Totaal aantal dagen voorraad+klantenkrediet"]
    df = df[df["Type"].isin(Klv)]

    df = df.T #Transponeren
    df = df.rename(index={"Boekjaar 1":"1","Boekjaar 2":"2",
                    "Boekjaar 3":"3"})
    df = df.iloc[1: , :] # Drop first row 
    df.insert(0,"Boekjaar",["Boekjaar 1","Boekjaar 2",
                    "Boekjaar 3"],True)
    df.columns = ["Boekjaar","Klantenkrediet","Totaal aantal dagen voorraad+klantenkrediet","Leverancierskrediet"] # change column names
    df = df.astype({'Klantenkrediet': 'float64','Leverancierskrediet': 'float64','Totaal aantal dagen voorraad+klantenkrediet': 'float64'})
    df = df.round({'Klantenkrediet':2, 'Leverancierskrediet':2, 'Totaal aantal dagen voorraad+klantenkrediet':2})
    return df


df_Klv = get_Klv_from_excel()

fig_Klv = px.bar(df_Klv, x=["Klantenkrediet","Totaal aantal dagen voorraad+klantenkrediet","Leverancierskrediet"], y="Boekjaar",
    orientation='h',
    barmode='group',
    range_x=[0,100],
    labels={'value':'aantal dagen'},
    text_auto=True)
fig_Klv.update_traces()
fig_Klv.update_layout()
st.plotly_chart(fig_Klv, use_container_width=True)

#---- Sidebar ----
st.sidebar.header("Filter:")
Liquiditeit = st.sidebar.multiselect(
    "Kies welk boekjaar:",
    options=df_liq["Boekjaar"].unique(),
    default=df_liq["Boekjaar"].unique()
)
# ---- Hide streamlit features----
hide_st_style = """
                <style>
                #Mainmenu {visibility: hidden;}
                footer {visibility: hidden;}
                header {visibility: hidden;}
                </style>"""
st.markdown(hide_st_style, unsafe_allow_html=True)