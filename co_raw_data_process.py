import pandas as pd
import streamlit as st

def process_co_data(df, scale_value):
    # Extract CO columns and weights
    co_columns = df.iloc[0, 1:].values  # Skip the first column for CO names
    print("CO Columns Identified:", co_columns)
    
    max_marks = df.iloc[1, 1:].astype(float).values  # Max marks for each CO
    print("Max Marks:", max_marks)

    # Collect student marks starting from the third row
    student_marks = df.iloc[2:, 1:].astype(float)  
    print("Student Marks Data Frame:\n", student_marks)
    
    # Prepare the grouping for CO scores
    co_totals = {}
    for i, co in enumerate(co_columns):
        if co not in co_totals:
            co_totals[co] = []
        co_totals[co].append(max_marks[i])
    
    print("CO Grouped Totals:", co_totals)
    
    # Calculate total marks for each CO
    co_final_scores = {co: sum(marks) for co, marks in co_totals.items()}
    print("Final Scores Per CO:", co_final_scores)
    
    # Prepare DataFrame for output
    co_final_df = pd.DataFrame.from_dict(co_final_scores, orient='index', columns=['Total Marks'])
    
    # Validate dimensions before matrix operations
    print(f"Student marks shape: {student_marks.values.shape}")
    print(f"CO final scores: {list(co_final_scores.values())}")
    
    # Ensure compatibility of matrix dimensions for the dot product
    total_co_marks = sum(co_final_scores.values())
    co_weights = list(co_final_scores.values())
    if len(co_weights) != student_marks.shape[1]:
        raise ValueError("Mismatch between student marks columns and CO weight length.")
    
    weighted_student_scores = (student_marks.values @ (co_weights / total_co_marks)) * (scale_value / 50)
    print("Weighted Student Scores:", weighted_student_scores)
    
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
    try:
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

    except Exception as e:
        st.error(f"An error occurred: {e}")
