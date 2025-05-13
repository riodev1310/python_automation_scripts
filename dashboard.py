import streamlit as st 
import pandas as pd 
import matplotlib.pyplot as plt
import seaborn as sns


st.set_page_config(page_title="Superstore", page_icon=":bar_chart:", layout="wide")

st.title(":bar_chart: Sample Superstore EDA")

# Read data from csv file
data = pd.read_csv("https://raw.githubusercontent.com/riodev1310/rio_datasets/refs/heads/main/Sample-Superstore.csv")
data["Order Date"] = pd.to_datetime(data["Order Date"])

# Create side bar
st.sidebar.header("Choose your filter: ")

# Create your Region
region = st.sidebar.multiselect("Pick your Region", data["Region"].unique())

if not region:
    data2 = data.copy()
else:
    # Create separate dataset for region
    data2 = data[data["Region"].isin(region)]

state = st.sidebar.multiselect("Pick the State", data2["State"].unique())

if state:
    data2 = data2[data2["State"].isin(state)]

city = st.sidebar.multiselect("Pick the city", data2["City"].unique())

# Filter the data based on Region, State, and City
# if not region and not state and not city:
#     filtered_data = data
# elif not state and not city:
#     filtered_data = data[data["Region"]]

if city:
    data2 = data2[data2["City"].isin(city)]

col1, col2 = st.columns((2))

# Column chart
with col1:
    st.subheader("Category-wise Sales")
    category_df = data2.groupby(by=["Category"], as_index=False)["Sales"].sum()
    fig, ax = plt.subplots()
    sns.barplot(data=category_df, x="Category", y="Sales", ax=ax)
    ax.set_title("Category-wise Sales")
    st.pyplot(fig)

with col2:
    st.subheader("Region-wise Sales")
    region_sales = data2.groupby(by="Region", as_index=False)["Sales"].sum()
    fig, ax = plt.subplots(figsize=(2, 4))
    # wedges, texts, autotexts = 
    ax.pie(region_sales["Sales"], labels=region_sales["Region"], autopct="%1.f%%", wedgeprops={'edgecolor': 'white'})
    # centre_circle = plt.Circle((0, 0), 0.30, fc='white')  # fc = face color
    # fig.gca().add_artist(centre_circle)
    ax.set_title("Region-wise Sales", fontsize=6)
    st.pyplot(fig)


data2["month_year"] = data2["Order Date"].dt.to_period("M")
linechart = pd.DataFrame(data2.groupby(data2["month_year"].dt.strftime("%Y-%b"))["Sales"].sum()).reset_index()
st.subheader("Time Series Analysis")
fig, ax = plt.subplots(figsize=(10, 6))
# sns.lineplot(data=linechart, x="month_year", y="Sales", ax=ax)
ax.plot(linechart['month_year'], linechart['Sales'])
ax.set_title("Sales Over Time")
ax.tick_params(axis='x', rotation=45)
st.pyplot(fig)

st.subheader("Hierachical View of Sales")
treemap_data = data2.groupby(["Region", "Category"])["Sales"].sum().unstack()
st.dataframe(treemap_data.style.background_gradient(cmap="viridis"))

st.subheader("Sales vs Profit")
fig, ax = plt.subplots()
sns.scatterplot(data=data2, x="Sales", y="Profit", size="Quantity", sizes=(20, 200), ax=ax)
ax.set_title("Sales vs Profit")
st.pyplot(fig)

csv = data2.to_csv(index=False).encode('utf-8')
st.download_button('Download Data', data=csv, file_name="Data.csv", mime="text/csv")