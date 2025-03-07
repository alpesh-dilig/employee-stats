import streamlit as st
import pandas as pd
import numpy as np
import io
import matplotlib.pyplot as plt
import matplotlib.patheffects as path_effects

pd.set_option('future.no_silent_downcasting', True)

# Table names for easier manipulation
table_names = ["zaposleni", "otkazi", "clanovi_skupstine_radni", "clanovi_skupstine_van_radni"]

# Define age groups
bins = [15, 30, 40, 50, 60, 70]
labels = ["Od 15-30", "Od 31-40", "Od 41-50", "Od 51-60", "Od 61-70"]

# -------------------- Functions for DataFrame manipulation-------------------------------
def read_sheet_and_extract_data(uploaded_file, sheet_name):
    # Read the sheet into a DataFrame without headers (some tables may have different headers)
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)

        # Replace empty cells (NaN) with an empty string to facilitate processing
        df = df.fillna("")

        # Identify blank rows (rows where all columns are empty)
        blank_rows = df.apply(lambda row: all(row == ""), axis=1)

        # Find indexes of blank rows
        blank_row_indices = df.index[blank_rows].tolist()

        # Add start and end indices to extract tables dynamically
        table_start_indices = [0] + [idx + 1 for idx in blank_row_indices if idx + 1 < len(df)]
        table_end_indices = blank_row_indices + [len(df)]

        # Extract tables dynamically
        tables = []
        for start, end in zip(table_start_indices, table_end_indices):
            table = df.iloc[start:end].reset_index(drop=True)
            
            # Ignore completely empty tables (if there are multiple blank rows together)
            if not table.replace("", np.nan).dropna(how="all").empty:
                table.columns = table.iloc[0]
                table = table[1:].reset_index(drop=True)
                tables.append(table)

        tables.pop(-1)

        # Tables dictionary
        tables_dict = {}
        for i, table in enumerate(tables):
            tables_dict.update({table_names[i]: table})

        return df, tables_dict

# Function to convert dataframe to Excel in memory
def to_excel(df):
    # Create a BytesIO buffer to save the Excel file
    output = io.BytesIO()
    with pd.ExcelWriter(output) as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    return output.getvalue()


def calculate_number_emplyees(data_df)-> pd.DataFrame:
    """
    Calculates number of emplyees and distribution of men and women.
    """

    # Count by gender
    gender_counts = data_df["POL"].value_counts().reset_index()
    gender_counts.columns = ["POL", "BROJ"]

    # Calculate percentage
    total_count = gender_counts["BROJ"].sum()
    gender_counts["PROCENAT"] = (gender_counts["BROJ"] / total_count * 100).round(2)

    # Append total row
    total_row = pd.DataFrame([["UKUPNO", total_count, 100.0]], columns=["POL", "BROJ", "PROCENAT"])
    result_df = pd.concat([gender_counts, total_row], ignore_index=True)

    return result_df


def calculate_age_structure(data_df) -> pd.DataFrame:
    """
    Calcualtes age structutre of employes and distribution of men and women.
    """

    # Categorize ages into bins
    data_df["Starosna granica"] = pd.cut(data_df["GODINE STAROSTI"], bins=bins, labels=labels, right=True)

    # Create a complete dataframe with all age categories
    all_ranges = pd.DataFrame({"Starosna granica": labels})

    # Convert categorical to string
    data_df["Starosna granica"] = data_df["Starosna granica"].astype(str)
    
    # first column name : ZAPOSLENI, etc...
    first_column_name = data_df.columns[0]
    # Total employes per age group
    age_struct_df = data_df.groupby("Starosna granica", observed=False)[first_column_name].count().reset_index()
    age_struct_df.columns = ["Starosna granica", "Ukupan broj"]

    # Merge with all possible age ranges
    age_struct_df = all_ranges.merge(age_struct_df, on="Starosna granica", how="left").fillna(0)

    # Caluclate total percentage
    total_employees = age_struct_df["Ukupan broj"].sum()
    age_struct_df["%"] = (age_struct_df["Ukupan broj"] / total_employees * 100).round(2)

    # Count females and males in each group
    gender_counts = data_df.groupby(["Starosna granica", "POL"], observed=False)[first_column_name].count().unstack(fill_value=0)

    # Add female & male counts
    age_struct_df["Žene"] = age_struct_df["Starosna granica"].map(gender_counts.get("Ž", {})).fillna(0).astype(int)
    age_struct_df["% Žene"] = (age_struct_df["Žene"] / age_struct_df["Ukupan broj"] * 100).round(2)

    age_struct_df["Muškarci"] = age_struct_df["Starosna granica"].map(gender_counts.get("M", {})).fillna(0).astype(int)
    age_struct_df["% Muškarci"] = (age_struct_df["Muškarci"] / age_struct_df["Ukupan broj"] * 100).round(2)

    # Replace NaN with 0
    age_struct_df.fillna(0, inplace=True)

    age_struct_df["Ukupan broj"] = pd.to_numeric(age_struct_df["Ukupan broj"], errors='coerce').fillna(0).astype(int)

    return age_struct_df


def calculate_executive_positions(data_df)-> pd.DataFrame:
    """
    Calculate number and percentage of executive positions and distribution of men and women.
    """

    # Define job categories
    executive_titles = ["Direktor"]  # Adjust this list based on actual executive titles
    
    # Categorize jobs
    data_df["Job Category"] = data_df["RADNO MESTO"].apply(lambda x: "Rukovodeca mesta" if x in executive_titles else "Izvršilačka mesta")
    
    # Total employees per category
    job_counts = data_df.groupby("Job Category")["ZAPOSLENI"].count().reset_index()
    job_counts.columns = ["Radna mesta", "Ukupan broj"]
    
    # Calculate percentage of total
    total_employees = job_counts["Ukupan broj"].sum()
    job_counts["%"] = (job_counts["Ukupan broj"] / total_employees * 100).round(2).astype(str) + "%"
    
    # Gender breakdown per category
    gender_counts = data_df.groupby(["Job Category", "POL"])["ZAPOSLENI"].count().unstack(fill_value=0).reset_index()
    
    # Ensure "Radna mesta" exists in gender_counts
    gender_counts.rename(columns={"Job Category": "Radna mesta"}, inplace=True)

    # Merge gender counts into job_counts
    job_counts = job_counts.merge(gender_counts, on="Radna mesta", how="left").fillna(0)
    
    # Compute gender percentages
    job_counts["% Žene"] = (job_counts.get("Ž", 0) / job_counts["Ukupan broj"] * 100).round(2).astype(str) + "%"
    job_counts["% Muškarci"] = (job_counts.get("M", 0) / job_counts["Ukupan broj"] * 100).round(2).astype(str) + "%"
    

    # Ensure required columns exist
    if "Ž" in job_counts.columns and "% Žene" in job_counts.columns:
        job_counts["Ukupan broj Žena"] = job_counts["Ž"].astype(str)
    if "M" in job_counts.columns and "% Muškarci" in job_counts.columns:
        job_counts["Ukupan broj Muškaraca"] = job_counts["M"].astype(str)

    
    return job_counts[["Radna mesta", "Ukupan broj", "%", "Ukupan broj Žena", "% Žene", "Ukupan broj Muškaraca", "% Muškarci"]]


def positions_statistics(data_df) -> pd.DataFrame:
    """
    Calculates statistics for each position and distribution between men and women.
    """
    # Count total employees per position
    position_counts = data_df['RADNO MESTO'].value_counts().reset_index()
    position_counts.columns = ['Naziv pozicije', 'Br ljudi po poziciji']
    
    # Count male and female employees per position
    gender_counts = data_df.groupby(['RADNO MESTO', 'POL']).size().unstack(fill_value=0)
    gender_counts.columns = ['Ž', 'M']  # Ensure correct ordering of columns
    
    # Merge both counts
    stats_df = position_counts.merge(gender_counts, left_on='Naziv pozicije', right_index=True, how='left').fillna(0)
    
    # Compute percentages
    total_employees = stats_df['Br ljudi po poziciji'].sum()
    stats_df['% Br ljudi po poziciji'] = (stats_df['Br ljudi po poziciji'] / total_employees * 100).round(2)
    stats_df['% M'] = (stats_df['M'] / total_employees * 100).round(2)
    stats_df['% Ž'] = (stats_df['Ž'] / total_employees * 100).round(2)
    
    # Add total row
    total_row = {
        'Naziv pozicije': 'Ukupno',
        'Br ljudi po poziciji': stats_df['Br ljudi po poziciji'].sum(),
        'M': stats_df['M'].sum(),
        'Ž': stats_df['Ž'].sum(),
        '% Br ljudi po poziciji': 100.00,
        '% M': stats_df['% M'].sum(),
        '% Ž': stats_df['% Ž'].sum()
    }
    stats_df = pd.concat([stats_df, pd.DataFrame([total_row])], ignore_index=True)
    
    return stats_df


def calculate_qualification(df) -> pd.DataFrame:
    """
    This function calculates qualifications for employees and dispresion between male and female.
    """
    # Group by gender and qualification level (counts)
    grouped = df.groupby(["POL", "STRUCNA SPREMA"]).size().unstack(fill_value=0)

    # Add total per gender
    grouped["UKUPNO"] = grouped.sum(axis=1)

    # Convert column names from Arabic to Roman numerals
    qualification_map = {4: "IV", 5: "V", 6: "VI", 7: "VII"}
    grouped = grouped.rename(columns=lambda x: qualification_map.get(x, x))

    # Calculate total employees
    total_employees = grouped["UKUPNO"].sum()

    # Ensure numeric operations apply only to numeric columns
    numeric_columns = grouped.columns  # Select all numeric columns

    # Calculate percentage per row (avoid string conversion issue)
    grouped_percentage = (grouped[numeric_columns].div(total_employees) * 100).round(2)

    # Interleave count and percentage columns
    formatted_df = pd.DataFrame()
    for col in grouped.columns:
        formatted_df[col] = grouped[col]
        formatted_df[f"{col} %"] = grouped_percentage[col]

    # Reset index and rename 'POL' to 'Strucna sprema'
    formatted_df = formatted_df.reset_index().rename(columns={"POL": "Strucna sprema"})

    # Rename gender values
    formatted_df["Strucna sprema"] = formatted_df["Strucna sprema"].replace({"M": "Muškarci", "Ž": "Žene"})

    # Compute total row separately, ensuring numeric columns are correctly handled
    total_row = grouped.sum().to_frame().T
    total_row["Strucna sprema"] = "UKUPNO"

    # Ensure only numeric columns are used for percentage calculations
    total_percentage = (total_row[numeric_columns].div(total_employees) * 100).round(2)

    # Create total row with interleaved percentage columns
    total_row_formatted = pd.DataFrame()
    for col in grouped.columns:
        total_row_formatted[col] = total_row[col]
        total_row_formatted[f"{col} %"] = total_percentage[col]

    # Append total row
    total_row_formatted["Strucna sprema"] = "UKUPNO"
    formatted_df = pd.concat([formatted_df, total_row_formatted], ignore_index=True)

    formatted_df = formatted_df.drop(formatted_df.columns[-1], axis=1)

    return formatted_df


def calculate_dismissals(df) -> pd.DataFrame:
    """
    This function calculates statiscitc for employees that are dismissed.
    """
    return calculate_age_structure(df)


def calculate_new_employes(df, target_date) -> pd.DataFrame:
    """
    This function calulates age structure and dispersion for new employees.
    """
    # Ensure the 'date' column is in datetime format
    df['STARTNI DATUM'] = pd.to_datetime(df['STARTNI DATUM'], format='%d/%m/%Y')
    
    # Convert target_date to datetime if it's not already
    target_date = pd.to_datetime(target_date)
    
    # Get today's date
    today = pd.to_datetime('today')
    
    # Filter rows where 'date' is between target_date and today (inclusive)
    filtered_df = df[(df['STARTNI DATUM'] >= target_date) & (df['STARTNI DATUM'] <= today)]

    result_df = calculate_age_structure(filtered_df)

    return result_df


# --------------- Functions for Streamlit display -------------------------
def show_number_of_employees(number_of_empolyes_df: pd.DataFrame):
     # Number of employees
        st.write("**Broj zaposlenih i raspodela izmedju muskih i zenskih zaposlenih**")
        container_employes = st.container(border=True)
        container_employes.dataframe(number_of_empolyes_df, use_container_width=True)
            # Buttons
        #download, show = container_employes.columns(spec=2)
        container_employes.download_button(
            label="Preuzmi kao Excel",
            data=num_employes_excel,
            file_name="number_of_employes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        with st.expander("Prikazi Chartove"):
            # Charts and data filtering
            employes_df = number_of_empolyes_df[number_of_empolyes_df['POL']!= "UKUPNO"]
            labels = employes_df['POL'].values
            sizes = employes_df['BROJ'].values
                # Plot
            fig = plt.figure(figsize=(2,2))
            plt.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90,wedgeprops={'edgecolor': 'black'})
            plt.title('Raspodela medju polovima')
            st.pyplot(fig=fig, use_container_width=False)


def show_age_structure(age_structure_df: pd.DataFrame):
    # Age structure
        st.write("**Starosna struktura**")
        conatainer_age = st.container(border=True)
        conatainer_age.dataframe(age_structure_df, use_container_width=True)
        age_df = age_structure_df[["Starosna granica", "Žene", "Muškarci"]]
            # Button for download
        conatainer_age.download_button(
            label="Preuzmi kao Excel",
            data=age_excel,
            file_name="age_structre.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        with st.expander("Prikazi Chartove"):
            df_filtered = age_structure_df[age_structure_df["Ukupan broj"]> 0]
            # main pie chart
            fig_main, ax = plt.subplots(figsize=(2,4))
            wedges, texts, autotexts = ax.pie(
                df_filtered['Ukupan broj'], 
                autopct='%1.1f%%', 
                startangle=90, 
                wedgeprops={'edgecolor': 'black'},
                labeldistance=1.1,
                #pctdistance=0.85
            )
            ax.legend(
                wedges, 
                df_filtered['Starosna granica'], 
                loc="center left",  # Move legend outside
                bbox_to_anchor=(1, 0.5),  # Position legend to the right
                ncol=1,  # Use multiple columns
                fontsize=8  # Adjust font size if needed
            )
            #plt.pie(df_filtered['Ukupan broj'], labels=df_filtered['Starosna granica'], autopct='%1.1f%%', startangle=90, wedgeprops={'edgecolor': 'black'})
            plt.title("Raspodela medju starosnim granicama")
            st.pyplot(fig=fig_main, use_container_width=False)
            
            # Other pie charts
            pie_charts = []

            for _, row in df_filtered.iterrows():
                labels = ['Žene', 'Muškarci']
                sizes = [row['Žene'], row['Muškarci']]

                fig= plt.figure(figsize=(2, 2))
                plt.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, wedgeprops={'edgecolor': 'black'}, colors=["red", "blue"])
                #plt.legend([f"Ukupan broj {row['Ukupan broj']}"], loc="upper right")
                plt.title(f"Raspodela polova po starosnoj granici ({row['Starosna granica']})")
                
                pie_charts.append(fig)
            
            for chart in pie_charts:
                st.pyplot(fig=chart, use_container_width=False)


def show_executive(executive_position_df: pd.DataFrame):
    # Executive
        st.write("**Ukupan broj i procenat izvrsilackih i rukovodecih mesta**")
        container_executive = st.container(border=True)
        container_executive.dataframe(executive_position_df,use_container_width=True)
        container_executive.download_button(
            label="Preuzmi kao Excel",
            data=exec_excel,
            file_name="executive_roles.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


def show_positions(positions_df):
     # Positions
        st.write("**Pozicije statistika**")
        container_positions = st.container(border=True)
        container_positions.dataframe(positions_df, use_container_width=True)
        container_positions.download_button(
            label="Preuzmi kao Excel",
            data=positions_excel,
            file_name="positions_of_employes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        pos_df = positions_df[positions_df['Naziv pozicije'] != "Ukupno"]
        with st.expander("Prikazi Chartove"):
            # main pie chart
            fig_main, ax = plt.subplots(figsize=(8, 6))
            wedges, texts, autotexts = ax.pie(
                pos_df['Br ljudi po poziciji'], 
                autopct='%1.1f%%', 
                startangle=90, 
                wedgeprops={'edgecolor': 'black'},
                labeldistance=1.1,
                pctdistance=1.2
            )

            # Improve the labels by adding leader lines
            for text in texts:
                text.set_fontsize(7)
                text.set_color('black')

            for autotext in autotexts:
                autotext.set_fontsize(6)
                autotext.set_color('black')
            
            ax.legend(
                wedges, 
                pos_df['Naziv pozicije'], 
                loc="center left",  # Move legend outside
                bbox_to_anchor=(1, 0.5),  # Position legend to the right
                ncol=1,  # Use multiple columns
                fontsize=8  # Adjust font size if needed
            )
            # plt.pie(pos_df['Br ljudi po poziciji'], autopct='%1.1f%%', startangle=90, wedgeprops={'edgecolor': 'black'})
            # plt.legend(labels=pos_df['Naziv pozicije'], loc='best', )
            plt.title("Raspodela medju pozicijama")
            st.pyplot(fig=fig_main, use_container_width=False)

            # Other pie charts
            pie_charts = []

            for _, row in pos_df.iterrows():
                labels = ['Ž', 'M']
                sizes = [row['Ž'], row['M']]

                fig= plt.figure(figsize=(2, 2))
                plt.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, wedgeprops={'edgecolor': 'black'}, colors=["red", "blue"])
                plt.title(f"Raspodela polova po pozicijama ({row['Naziv pozicije']})")
                
                pie_charts.append(fig)
            
            for chart in pie_charts:
                st.pyplot(fig=chart, use_container_width=False)


def show_qualifiactions(qualifications_df):
    # Qualifications
        st.write("**Kvalifikacije**")
        container_qualifications = st.container(border=True)
        container_qualifications.dataframe(qualifications_df, use_container_width=True)
        container_qualifications.download_button(
            label="Preuzmi kao Excel",
            data=qualifications_excel,
            file_name="qualifications_of_employes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        qua_df = qualifications_df[qualifications_df['Strucna sprema'] != "UKUPNO"]
        # Exclude columns
        exclude_columns = {"Strucna sprema", "UKUPNO"}

        labels = [col for col in qualifications_df.columns if col not in exclude_columns and "%" not in col]
        muskarci= qualifications_df[qualifications_df["Strucna sprema"] == "Muškarci"][labels].iloc[0].to_list()
        zene= qualifications_df[qualifications_df["Strucna sprema"] == "Žene"][labels].iloc[0].to_list()
        ukupno= qualifications_df[qualifications_df["Strucna sprema"] == "UKUPNO"][labels].iloc[0].to_list()

        x = np.arange(len(labels))  # Lokacije za x os
        width = 0.3  # Širina kolona

        # get colors
        label_indics = [labels.index(label) for label in labels]
        cmap = plt.get_cmap("tab10")
        colors = [cmap(i / max(label_indics)) for i in label_indics]

        with st.expander("Prikazi Chartove"):
            fig, ax = plt.subplots(figsize=(8, 6))

            # Bar chart
            ax.bar(x - width, muskarci, width, label="Muškarci", color='blue')
            ax.bar(x, zene, width, label="Žene", color='red')
            ax.bar(x + width, ukupno, width, label="Ukupno", color='gray')

            ax.set_xlabel("Stručna sprema")
            ax.set_ylabel("Broj zaposlenih")
            ax.set_title("Distribucija zaposlenih prema stručnoj spremi")
            ax.set_xticks(x)
            ax.set_xticklabels(labels)
            ax.legend()
            st.pyplot(fig=fig, use_container_width=False)

            # Pie chart
            fig, ax = plt.subplots(figsize=(6,6))
            ax.pie(ukupno, labels=labels, autopct='%1.1f%%', colors=colors)
            ax.set_title("Ukupna raspodela zaposlenih prema strucnoj spremi")
            st.pyplot(fig=fig, use_container_width=False)


def show_otkazi(dismissales_df):
    # Otkazi
        st.write("**Otkazi**")
        container_quit = st.container(border=True)
        container_quit.dataframe(dismissales_df, use_container_width=True)
        container_quit.download_button(
            label="Preuzmi kao Excel",
            data=dismissales_excel,
            file_name="dismissed_employes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        with st.expander("Prikazi Chartove"):
            df_filtered = dismissales_df[dismissales_df["Ukupan broj"]> 0]
            # main pie chart
            fig_main, ax = plt.subplots(figsize=(3,4))
            wedges, texts, autotexts = ax.pie(
                df_filtered['Ukupan broj'], 
                autopct='%1.1f%%', 
                startangle=90, 
                wedgeprops={'edgecolor': 'black'},
                labeldistance=1.1,
                pctdistance=1.25
            )
            ax.legend(
                wedges, 
                df_filtered['Starosna granica'], 
                loc="center left",  # Move legend outside
                bbox_to_anchor=(1, 0.5),  # Position legend to the right
                ncol=1,  # Use multiple columns
                fontsize=8  # Adjust font size if needed
            )
            #plt.pie(df_filtered['Ukupan broj'], labels=df_filtered['Starosna granica'], autopct='%1.1f%%', startangle=90, wedgeprops={'edgecolor': 'black'})
            plt.title("Raspodela medju starosnim granicama")
            st.pyplot(fig=fig_main, use_container_width=False)
            
            # Other pie charts
            pie_charts = []

            for _, row in df_filtered.iterrows():
                labels = ['Žene', 'Muškarci']
                sizes = [row['Žene'], row['Muškarci']]

                fig= plt.figure(figsize=(2, 2))
                plt.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, wedgeprops={'edgecolor': 'black'}, colors=["red", "blue"])
                #plt.legend([f"Ukupan broj {row['Ukupan broj']}"], loc="upper right")
                plt.title(f"Raspodela polova po starosnoj granici ({row['Starosna granica']})")
                
                pie_charts.append(fig)
            
            for chart in pie_charts:
                st.pyplot(fig=chart, use_container_width=False)


def show_newbees(new_eployees_df):
    # New employees
        st.write("**Novozaposleni**")
        container_newempl = st.container(border=True)
        container_newempl.dataframe(new_eployees_df, use_container_width=True)
        container_newempl.download_button(
            label="Preuzmi kao Excel",
            data=new_eployees_excel,
            file_name="new_employes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        with st.expander("Prikazi Chartove"):
            df_filtered = new_eployees_df[new_eployees_df["Ukupan broj"]> 0]
            # main pie chart
            fig_main, ax = plt.subplots(figsize=(3,4))
            wedges, texts, autotexts = ax.pie(
                df_filtered['Ukupan broj'], 
                autopct='%1.1f%%', 
                startangle=90, 
                wedgeprops={'edgecolor': 'black'},
                labeldistance=1.1,
                #pctdistance=1.25
            )
            ax.legend(
                wedges, 
                df_filtered['Starosna granica'], 
                loc="center left",  # Move legend outside
                bbox_to_anchor=(1, 0.5),  # Position legend to the right
                ncol=1,  # Use multiple columns
                fontsize=8  # Adjust font size if needed
            )
            #plt.pie(df_filtered['Ukupan broj'], labels=df_filtered['Starosna granica'], autopct='%1.1f%%', startangle=90, wedgeprops={'edgecolor': 'black'})
            plt.title("Raspodela medju starosnim granicama")
            st.pyplot(fig=fig_main, use_container_width=False)
            
            # Other pie charts
            pie_charts = []

            for _, row in df_filtered.iterrows():
                labels = ['Žene', 'Muškarci']
                sizes = [row['Žene'], row['Muškarci']]

                fig= plt.figure(figsize=(2, 2))
                plt.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, wedgeprops={'edgecolor': 'black'}, colors=["red", "blue"])
                #plt.legend([f"Ukupan broj {row['Ukupan broj']}"], loc="upper right")
                plt.title(f"Raspodela polova po starosnoj granici ({row['Starosna granica']})")
                
                pie_charts.append(fig)
            
            for chart in pie_charts:
                st.pyplot(fig=chart, use_container_width=False)

# ----------------- Start ---------------------------------
# Streamlit UI
st.title('Informacije o zaposlenima')

# Container for input data
container_input = st.container(border=True)

# Upload file
uploaded_file = container_input.file_uploader("Izaberi Excel fajl", type="xlsx")

# Date from when Novozaposleni
date = container_input.date_input("Izaberi datum od kog se racunaju novozaposleni")



if uploaded_file:
    # Read the Excel file as a dataframe
    try:
        # Extract sheet names
        xls = pd.ExcelFile(uploaded_file)
        sheets = xls.sheet_names
        selected_option = container_input.selectbox("Izaberi sheet:", sheets)

        sheet_name = selected_option

        df, tables_dict = read_sheet_and_extract_data(uploaded_file, sheet_name)

        # --------------- Perform operations and make dataframes -------------------------

        number_of_empolyes_df = calculate_number_emplyees(tables_dict["zaposleni"])
        age_structure_df = calculate_age_structure(tables_dict["zaposleni"])
        executive_position_df = calculate_executive_positions(tables_dict["zaposleni"])
        positions_df = positions_statistics(tables_dict["zaposleni"])
        qualifications_df = calculate_qualification(tables_dict["zaposleni"])
        dismissales_df = calculate_dismissals(tables_dict["otkazi"])
        new_eployees_df = calculate_new_employes(tables_dict["zaposleni"], target_date=date)

        # ------------- Convert DataFrames to Excel ----------------------------
        num_employes_excel = to_excel(number_of_empolyes_df)
        age_excel = to_excel(age_structure_df)
        exec_excel = to_excel(executive_position_df)
        positions_excel = to_excel(positions_df)
        qualifications_excel = to_excel(qualifications_df)
        dismissales_excel = to_excel(dismissales_df)
        new_eployees_excel = to_excel(new_eployees_df)

        # ------------------ Show in the app -------------------------

        show_number_of_employees(number_of_empolyes_df)
        show_age_structure(age_structure_df)
        show_executive(executive_position_df)
        show_positions(positions_df)
        show_qualifiactions(qualifications_df)
        show_otkazi(dismissales_df)
        show_newbees(new_eployees_df)       
        

    except Exception as e:
        st.error(f"Error: {e}")