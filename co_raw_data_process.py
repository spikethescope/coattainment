
import streamlit as st
import pandas as pd
import numpy as np
import io

def process_co_data(df, co_weights, round_digits=2):
    # Find the row with CO labels
    co_row = None
    for i, row in df.iterrows():
        if any(co in str(cell) for co in ['CO1', 'CO2', 'CO3', 'CO4', 'CO5', 'CO6'] for cell in row):
            co_row = i
            break
    
    if co_row is None:
        raise ValueError("No CO labels found in the file")

    # Extract headers, CO labels, max marks, and student marks
    headers = df.iloc[:co_row].values.tolist() if co_row > 0 else []
    co_labels = df.iloc[co_row].tolist()
    max_marks = df.iloc[co_row + 1].tolist()
    student_marks = df.iloc[co_row + 2:].values.tolist()
    
    # Get unique CO components and sort them
    unique_cos = sorted([col for col in co_labels if isinstance(col, str) and col.startswith('CO')], 
                        key=lambda x: int(x[2:]))
    
    # Group COs and their corresponding marks
    co_groups = {co: [] for co in unique_cos}
    for co, max_mark in zip(co_labels, max_marks):
        if co in co_groups:
            co_groups[co].append(max_mark)
    
    # Calculate total marks for each CO group
    co_totals = {co: sum(marks) for co, marks in co_groups.items()}
    
    # Calculate total of all CO marks
    total_co_marks = sum(co_totals.values())
    
    # Calculate weighted max marks for each CO (without rounding)
    weighted_max_marks = {co: total_co_marks * co_weights[co] for co in unique_cos}
    
    # Process student marks
    processed_student_marks = []
    for student in student_marks:
        student_co_marks = {}
        for co in unique_cos:
            co_indices = [i for i, label in enumerate(co_labels) if label == co]
            student_co_total = sum(student[i] for i in co_indices)
            max_co_total = sum(co_groups[co])
            # Round only the student marks
            student_co_marks[co] = round((student_co_total / max_co_total) * weighted_max_marks[co], round_digits)
        processed_student_marks.append(student_co_marks)
    
    # Prepare output data
    output_data = headers + [
        ["CO"] + co_labels,
        ["Weighted Max Marks"] + [weighted_max_marks.get(co, "") if co in unique_cos else "" for co in co_labels]
    ]
    for i, student in enumerate(processed_student_marks, start=1):
        output_data.append([f"Student {i}"] + [student.get(co, "") if co in unique_cos else "" for co in co_labels])
    
    # Create output DataFrame
    output_df = pd.DataFrame(output_data)
    
    return output_df

# Streamlit app
st.title('Course Outcome (CO) Data Processor')

# File upload
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    # Read the Excel file
    df = pd.read_excel(uploaded_file, header=None)
    
    # Display input data
    st.subheader("Input Data")
    st.dataframe(df)
    
    # Find CO labels
    co_row = next((i for i, row in df.iterrows() if any(co in str(cell) for co in ['CO1', 'CO2', 'CO3', 'CO4', 'CO5', 'CO6'] for cell in row))), None)
    if co_row is None:
        st.error("No CO labels found in the file")
    else:
        #co_labels = [col for col in df.iloc[co_row] if isinstance(col, str) and col.startswith('CO')]
        co_labels = list({col for col in df.iloc[co_row] if isinstance(col, str) and col.startswith('CO')})

        # Create input fields for CO weights
        st.write(f"The list of CO labels are {co_labels}")
        st.subheader("Enter weightage for each CO component")
        st.write("The sum of all weights should be 1.")
        
        co_weights = {}
        col1, col2 = st.columns(2)
        for i, co in enumerate(co_labels):
            if i % 2 == 0:
                co_weights[co] = col1.number_input(f"Weight for {co}", min_value=0.0, max_value=1.0, value=1.0/len(co_labels), step=0.01, format="%.2f")
            else:
                co_weights[co] = col2.number_input(f"Weight for {co}", min_value=0.0, max_value=1.0, value=1.0/len(co_labels), step=0.01, format="%.2f")
        
        # Input for rounding digits
        round_digits = st.number_input("Number of decimal places for rounding", min_value=0, max_value=6, value=2)
        
        # Check if weights sum to 1
        total_weight = sum(co_weights.values())
        if abs(total_weight - 1.0) > 1e-6:  # Allow for small floating-point errors
            st.warning(f"The sum of weights is {total_weight:.2f}. It should be 1.0.")
        else:
            if st.button("Process Data"):
                # Process the data
                output_df = process_co_data(df, co_weights, round_digits)
                
                # Display output data
                st.subheader("Processed Data")
                st.dataframe(output_df)
                
                # Provide download link for processed data
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    output_df.to_excel(writer, index=False, header=False)
                output.seek(0)
                
                st.download_button(
                    label="Download Processed Data",
                    data=output,
                    file_name="processed_co_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

st.markdown("""
### Instructions:
1. Upload an Excel file containing CO data.
2. The file should have CO labels in a row, with max marks in the row below, and student marks in subsequent rows. Any headers above the CO labels will be included in the output.
3. Enter the weightage for each CO component. The sum of all weights should be 1.
4. Specify the number of decimal places for rounding student marks.
5. Click 'Process Data' to see the results.
6. You can download the processed data as an Excel file.
""")
