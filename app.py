import streamlit as st
import pandas as pd

def main():
    st.title("Blitz Report Data Cleaner!")

    uploaded_file = st.file_uploader("Choose an XLSX file", type="xlsx")

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file, sheet_name=None)

            # Filter sheets to include only "Blitz Summary" and "IR Detail"
            allowed_sheets = ["Blitz Summary", "IR Detail"]
            filtered_df = {sheet: df[sheet] for sheet in df if sheet in allowed_sheets}

            st.sidebar.header("Select Sheet:")
            selected_sheet = st.sidebar.selectbox("Sheet Name", list(filtered_df.keys()))

            st.header(f"Sheet: {selected_sheet}")

            if selected_sheet == "Blitz Summary":
                # Rename specific columns by index
                rename_dict = {
                    filtered_df[selected_sheet].columns[141]: "Domestic Data Calls",
                    filtered_df[selected_sheet].columns[153]: "Domestic Data Minutes Of Use",
                    filtered_df[selected_sheet].columns[165]: "Domestic Data Usage (MB)",
                    filtered_df[selected_sheet].columns[177]: "Domestic Data Excess Usage ($)",
                    filtered_df[selected_sheet].columns[189]: "Overseas Data Calls",
                    filtered_df[selected_sheet].columns[201]: "Overseas Data Minutes Of Use",
                    filtered_df[selected_sheet].columns[213]: "Overseas Data Usage (MB)",
                    filtered_df[selected_sheet].columns[225]: "Overseas Data Excess Usage ($)",
                    filtered_df[selected_sheet].columns[237]: "Domestic Voice Calls",
                    filtered_df[selected_sheet].columns[249]: "Domestic Voice Minutes Of Use",
                    filtered_df[selected_sheet].columns[261]: "Domestic Voice Spend ($)",
                    filtered_df[selected_sheet].columns[273]: "Domestic SMS",
                    filtered_df[selected_sheet].columns[285]: "Bill Total",
                    filtered_df[selected_sheet].columns[373]: "Total Discount Amount"
                }
                filtered_df[selected_sheet].rename(columns=rename_dict, inplace=True)

                # Delete columns by index ranges
                columns_to_delete = []
                for i in range(142, 153):
                    columns_to_delete.append(filtered_df[selected_sheet].columns[i])
                for i in range(154, 165):
                    columns_to_delete.append(filtered_df[selected_sheet].columns[i])
                for i in range(166, 177):
                    columns_to_delete.append(filtered_df[selected_sheet].columns[i])
                for i in range(178, 189):
                    columns_to_delete.append(filtered_df[selected_sheet].columns[i])
                for i in range(190, 201):
                    columns_to_delete.append(filtered_df[selected_sheet].columns[i])
                for i in range(202, 213):
                    columns_to_delete.append(filtered_df[selected_sheet].columns[i])
                for i in range(214, 225):
                    columns_to_delete.append(filtered_df[selected_sheet].columns[i])
                for i in range(226, 237):
                    columns_to_delete.append(filtered_df[selected_sheet].columns[i])
                for i in range(238, 249):
                    columns_to_delete.append(filtered_df[selected_sheet].columns[i])
                for i in range(250, 261):
                    columns_to_delete.append(filtered_df[selected_sheet].columns[i])
                for i in range(262, 273):
                    columns_to_delete.append(filtered_df[selected_sheet].columns[i])
                for i in range(274, 285):
                    columns_to_delete.append(filtered_df[selected_sheet].columns[i])
                for i in range(286, 302):
                    columns_to_delete.append(filtered_df[selected_sheet].columns[i])
                columns_to_delete += [
                    filtered_df[selected_sheet].columns[344], 
                    filtered_df[selected_sheet].columns[346], 
                    filtered_df[selected_sheet].columns[347], 
                    filtered_df[selected_sheet].columns[349], 
                    filtered_df[selected_sheet].columns[350], 
                    filtered_df[selected_sheet].columns[364], 
                    filtered_df[selected_sheet].columns[365], 
                    filtered_df[selected_sheet].columns[371], 
                    filtered_df[selected_sheet].columns[372], 
                    filtered_df[selected_sheet].columns[374], 
                    filtered_df[selected_sheet].columns[375], 
                    filtered_df[selected_sheet].columns[395], 
                    filtered_df[selected_sheet].columns[396],
                    filtered_df[selected_sheet].columns[397],
                ]

                filtered_df[selected_sheet].drop(columns=columns_to_delete, inplace=True)

                st.dataframe(filtered_df[selected_sheet])
                
                # Download button
                csv = filtered_df[selected_sheet].to_csv(index=False)
                st.download_button(
                    label="Download as CSV",
                    data=csv,
                    file_name=f"{selected_sheet}.csv",
                    mime='text/csv'
                )
                
                # Download Blitz Report as XLSX
                with pd.ExcelWriter(f"{selected_sheet}.xlsx", engine='openpyxl') as writer:
                    filtered_df[selected_sheet].to_excel(writer, sheet_name=selected_sheet, index=False)
                with open(f"{selected_sheet}.xlsx", "rb") as f:
                    xlsx = f.read()
                st.download_button(
                    label="Download Blitz Report as XLSX",
                    data=xlsx,
                    file_name=f"{selected_sheet}.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

            if selected_sheet == "IR Detail":
                # Rename specific columns by index
                rename_dict = {
                    filtered_df[selected_sheet].columns[170]: "Number Of $5 IR Daypass",
                    filtered_df[selected_sheet].columns[171]: "Number Of $10 IR Daypass",
                    filtered_df[selected_sheet].columns[172]: "Number Of $30 IR Daypass",
                    filtered_df[selected_sheet].columns[173]: "Number Of Monthly IR Pass",
                    filtered_df[selected_sheet].columns[174]: "Monthly IR Pass Amount"
                }
                filtered_df[selected_sheet].rename(columns=rename_dict, inplace=True)

                # Keep only selected columns
                filtered_df[selected_sheet] = filtered_df[selected_sheet].iloc[:, [3, 170, 171, 172, 173, 174]] 
                st.dataframe(filtered_df[selected_sheet])
                # Download button
                csv = filtered_df[selected_sheet].to_csv(index=False)
                st.download_button(
                    label="Download as CSV",
                    data=csv,
                    file_name=f"{selected_sheet}.csv",
                    mime='text/csv'
                )
                
                # Download IR Detail as XLSX
                with pd.ExcelWriter(f"{selected_sheet}.xlsx", engine='openpyxl') as writer:
                    filtered_df[selected_sheet].to_excel(writer, sheet_name=selected_sheet, index=False)
                with open(f"{selected_sheet}.xlsx", "rb") as f:
                    xlsx = f.read()
                st.download_button(
                    label="Download IR Detail as XLSX",
                    data=xlsx,
                    file_name=f"{selected_sheet}.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                
        except Exception as e:
            st.error(f"Error processing file: {e}")

if __name__ == "__main__":
    main()