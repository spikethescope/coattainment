import pandas as pd
import streamlit as st

def process_co_data(df, scale_value):
    # Extract CO columns and weights
    co_columns = df.columns[1:]  # Skip the first column (CO name)
    max_marks = df.iloc[0, 1:].astype(float)  # Max marks
    student_marks = df.iloc[1:-1, 1:].astype(float)  # Student data
    
    # Grouping the CO scores
    co_totals = {}
    for i, co in enumerate(co_columns):
        if co not in co_totals:
            co_totals[co] = []
        co_totals[co].append(float(max_marks[i]))
    
    # Calculate total marks for each CO
    co_final_scores = {co: sum(marks) for co, marks in co_totals.items()}
    
    # Prepare DataFrame for output
    co_final_df = pd.DataFrame.from_dict(co_final_scores, orient='index', columns=['Total Marks'])
    
    # Weighting the student marks based on CO marks
    total_co_marks = sum(co_final_scores.values())
    weighted_student_scores = (student_marks.values @ (list(co_final_scores.values()) / total_co_marks)) * (scale_value / 50)
    
    # Create the output DataFrame
    output_df = pd.DataFrame(weighted_student_scores, columns=['Weighted Score'])
    
    return co_final_df, output_df

# Streamlit interface
st.title("CO Marks Processor")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

scale_value = st.number_input("Enter the scale value (e.g., 50)", min_value=1, value=50)

if uploaded_file:
    # Read the Excel file
    df = pd.read_excel(uploaded_file, header=None)
    
    st.write("Input Data Frame:")
    st.dataframe(df)

    # Process the data
    co_final_df, output_df = process_co_data(df, scale_value)

    # Display processed results
    st.write("Total Marks for Each CO:")
    st.dataframe(co_final_df)
    
    st.write("Weighted Scores for Students:")
    st.dataframe(output_df)

    # Save the processed data to a new Excel file
    output_file = "processed_scores.xlsx"
    with pd.ExcelWriter(output_file) as writer:
        co_final_df.to_excel(writer, sheet_name='CO Totals')
        output_df.to_excel(writer, sheet_name='Student Scores')
    
    st.download_button("Download Processed Excel File", data=output_file, file_name=output_file)
