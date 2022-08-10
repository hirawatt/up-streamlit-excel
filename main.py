import streamlit as st
import pandas as pd


def main() -> None:
    st.write("Streamlit Excel")
    excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])

    if excel_file:
        df = pd.read_excel(excel_file)
        modified_df = df.astype(str)

        ## Operations on Excel
        modified_df.drop(columns=["Unnamed: 0"], inplace=True)
        st.write("Modified Excel", modified_df)

        ## Output to file
        modified_df.to_excel("output.xlsx")


if __name__ == "__main__":
    st.set_page_config(
        "Streamlit Excel",
        "🕴️",
        initial_sidebar_state="expanded",
        layout="wide",
    )
    main()