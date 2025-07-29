import cosmic_data_science.clean.standardize as standardize
import streamlit

streamlit.set_page_config(
    page_title="COSMIC Inventory Data Science",
    page_icon=":rocket:",
    layout="wide",
    initial_sidebar_state="expanded",
)

def main():
    streamlit.title("COSMIC Inventory Data Science")
    streamlit.write("Welcome to the COSMIC Inventory Data Science app!")
    streamlit.write("This app allows you to standardize and clean the COSMIC Technology Inventory Database Snapshot.")
    file_path = streamlit.file_uploader("Upload the COSMIC Technology Inventory Database Snapshot Excel file", type=["xlsx"])
    if file_path:
        df = standardize.standardize_inventory_data(file_path)
        streamlit.write("Data standardized successfully!")
        streamlit.dataframe(df)

if __name__ == "__main__":
    main()