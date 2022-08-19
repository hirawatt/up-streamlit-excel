import streamlit as st
import pandas as pd


def main() -> None:
    st.write("Streamlit Excel")
    excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])

    if excel_file:
        input_df = pd.read_excel(excel_file, header=1)
        #input_df = df.astype(str)
        st.write(input_df)
        input_df.shape

        # 2
        input_df = input_df.dropna(how = "all")
        input_df.shape

        # 3
        input_df.head()
        input_df.columns

        # 4
        output_df = input_df[["Date Echéance ", "E/C", "Montant TTC"]]

        grouped_df = output_df.groupby(["Date Echéance ", "E/C"]).sum()

        grouped_df = grouped_df.reset_index(level=['E/C'])

        ec_df = grouped_df.pivot(columns = 'E/C').droplevel(0, axis=1).fillna(0)

        #ec_df["CF"] = ec_df["Encaissement"] - ec_df["Décaissement"]

        ec_df.index = pd.to_datetime(ec_df.index, format = '%m/%d/%Y') #.strftime('%Y/%m/%d')
        ec_df["Total"] = ""
        ec_df = ec_df[["Encaissement", "Décaissement", "Total"]]
        #ec_df.loc['Total']= ec_df.sum(numeric_only=True, axis=0)
        ec_dfT = ec_df.T
        ec_dfT.head()

        # 5
        output_df = input_df[["Date Echéance ", "E/C", "Nature ", "Montant TTC"]]

        output_df.loc[output_df['E/C'] == 'Décaissement', 'Montant TTC'] = output_df["Montant TTC"]

        grouped_df = output_df.groupby(["Date Echéance ", "Nature "]).sum()

        grouped_df = grouped_df.reset_index(level=['Nature '])

        nature_df = grouped_df.pivot(columns = 'Nature ').droplevel(0, axis=1).fillna(0)

        nature_df.index = pd.to_datetime(nature_df.index, format = '%m/%d/%Y').strftime('%Y-%m-%d')
        nature_df["Total"] = ""
        nature_dfT = nature_df.T
        #nature_dfT.head()

        # 6
        current_date = ec_dfT.columns[0].strftime("%m%Y")
        decalage = 0
        index_totaux = []

        for index, column in enumerate(ec_dfT.columns):
            new_date = column.strftime("%m%Y")
            if current_date == new_date:
                pass
            else:
                ec_dfT.insert(index+decalage, "Total" + str(current_date), "")
                nature_dfT.insert(index+decalage, "Total" + str(current_date), "")
                index_totaux.append(index+decalage)
                decalage += 1
                current_date = new_date

        ec_dfT["Total" + current_date] = ""
        nature_dfT["Total" + current_date] = ""

        index_totaux.append(ec_dfT.shape[1]-1)
        #st.write(ec_dfT)

        # 7
        #nature_dfT
        #nature_dfT.index

        # 8
        new_index = ['VENTE - E BATTERIE', 'VENTE-POSE',
        'Vente - E Batterie', 'DEPENSES POSE', 'FOURNITURE ELECTRICITE',
        'FRAIS DEPLACEMENT', 'HONORAIRES', 'ASSURANCE', 'LOCATION', 'LOYERS',
        'REGULARISATIONS FOURNITURE', 'REMBOURSEMENT DETTE', 'SALAIRES',
        'TAXES DIVERSES', 'CVAE', 'TVA', 'Total']


        # 9
        nature_dfT.reindex(new_index)

        # 10
        #index_totaux


        # 11
        assert ec_df.shape[0] == nature_df.shape[0]
        assert ec_dfT.shape[1] == nature_dfT.shape[1]

        # 12
        import xlsxwriter

        writer = pd.ExcelWriter('outputv3.xlsx', engine='xlsxwriter')
        ec_dfT.to_excel(writer, sheet_name = "EC", header = True)

        workbook  = writer.book
        worksheet = writer.sheets['EC']

        START_COLUMN_TOTAL = 1
        END_COLUMN_TOTAL = ec_dfT.shape[1] + 1


        # Gestion des formats des totaux
        cell_format_totaux = workbook.add_format()
        cell_format_totaux.set_bg_color('#FBECE9')
        cell_format_totaux.set_bold()

        # Gestion du format de la date (pour format date et non str)

        format_date = workbook.add_format({'num_format': 'dd/mm/yy'})
        date_list = ec_dfT.columns
        for index, col_num in enumerate(range(START_COLUMN_TOTAL, END_COLUMN_TOTAL)):
            worksheet.write(xlsxwriter.utility.xl_col_to_name(col_num) + '1', date_list[index], format_date)       # 28/02/13
            

        # TOTAL E/C (final row)

        ROW_ENCAISSEMENT = 1
        ROW_DECAISSEMENT = 2
        ROW_TOTAL_EC = ROW_DECAISSEMENT + 1

        for col_num in range(START_COLUMN_TOTAL, END_COLUMN_TOTAL):
            col_letter = xlsxwriter.utility.xl_col_to_name(col_num)
            formula = '=' + col_letter + str(ROW_ENCAISSEMENT+1) + '-' + col_letter + str(ROW_DECAISSEMENT+1)
            worksheet.write_formula(ROW_TOTAL_EC, col_num, formula, cell_format_totaux)

        # TOTAL E/ C (by month)

        initial_col_num = START_COLUMN_TOTAL

        for col_num in index_totaux:
            initial_col_letter = xlsxwriter.utility.xl_col_to_name(initial_col_num)
            end_col_letter = xlsxwriter.utility.xl_col_to_name(col_num)
            for row_num in range(1, ROW_TOTAL_EC):
                formula = '=SUM(' + initial_col_letter + str(row_num+1) + ':' + end_col_letter + str(row_num+1) + ")" 
                worksheet.write_formula(row_num, col_num+1, formula, cell_format_totaux)
                initial_col_num = col_num + 2
                
        # CLOSING BALANCE 

        ROW_CLOSING_BALANCE = ROW_TOTAL_EC + 2

        # Initialisation de la première formule
        worksheet.write_formula(ROW_CLOSING_BALANCE, START_COLUMN_TOTAL, 
                                "=" + str(xlsxwriter.utility.xl_col_to_name(START_COLUMN_TOTAL)) + str(ROW_TOTAL_EC+1))

        for col_num in range(START_COLUMN_TOTAL, END_COLUMN_TOTAL-1):
            col_previous_letter = xlsxwriter.utility.xl_col_to_name(col_num)
            if col_num in index_totaux:
                formula = "="+col_previous_letter+str(ROW_CLOSING_BALANCE+1)
                worksheet.write(ROW_CLOSING_BALANCE, col_num+1, formula, cell_format_totaux)
            else:
                col_letter = xlsxwriter.utility.xl_col_to_name(col_num+1)
                formula = "="+col_previous_letter+str(ROW_CLOSING_BALANCE+1)+"+"+col_letter+str(ROW_TOTAL_EC+1)
                worksheet.write_formula(ROW_CLOSING_BALANCE, col_num+1, formula)

        worksheet.write('A' + str(ROW_CLOSING_BALANCE+1), 'Closing Balance')

        # PAR NATURE

        START_ROW_NATURE = ROW_CLOSING_BALANCE + 2
        ROW_TOTAL_NATURE = START_ROW_NATURE + nature_df.shape[1]

        nature_dfT.to_excel(writer, sheet_name = "EC", startrow=START_ROW_NATURE, header = False)

        for col_num in range(START_COLUMN_TOTAL, END_COLUMN_TOTAL):
            col_letter = xlsxwriter.utility.xl_col_to_name(col_num)
            formula = '=SUM(' + col_letter + str(START_ROW_NATURE+1) + ':' + col_letter + str(ROW_TOTAL_NATURE-1) +')'
            worksheet.write_formula(ROW_TOTAL_NATURE-1, col_num, formula, cell_format_totaux)

        # Total nature (by month)

        initial_col_num = START_COLUMN_TOTAL

        for col_num in index_totaux:
            initial_col_letter = xlsxwriter.utility.xl_col_to_name(initial_col_num)
            end_col_letter = xlsxwriter.utility.xl_col_to_name(col_num)
            for row_num in range(START_ROW_NATURE, ROW_TOTAL_NATURE-1):
                formula = '=SUM(' + initial_col_letter + str(row_num+1) + ':' + end_col_letter + str(row_num+1) + ")" 
                worksheet.write_formula(row_num, col_num+1, formula, cell_format_totaux)
                initial_col_num = col_num + 2

        # Gestion du format A

        cell_format_A = workbook.add_format()
        cell_format_A.set_bold()
        cell_format_A.set_align("center")

        worksheet.set_column('A:A', None, cell_format_A)

        # GROUPING DES COLONNES
        start_col = "B"

        for col_num in index_totaux:
            end_col_letter = xlsxwriter.utility.xl_col_to_name(col_num)
            formula = start_col + ":" + end_col_letter
            worksheet.set_column(formula, None, None, {'level': 1, 'hidden': True})
            start_col = xlsxwriter.utility.xl_col_to_name(col_num+2)
            
        writer.save()
        writer.close()
        print("over")
        
        import os
        # Edit this with the excel file variable
        output_file = os.getcwd() + '/output3.xlsx'
        st.write(output_file)
        st.download_button("Press to Download output", data=output_file, file_name="output.xlsx", mime="application/vnd.ms-excel")

        st.write("End")

if __name__ == "__main__":
    st.set_page_config(
        "Streamlit Excel",
        "🕴️",
        initial_sidebar_state="expanded",
        layout="wide",
    )
    main()