import streamlit as st
import pandas as pd
import numpy as np
import io

def process_co_data(df, total_marks):
    # Extract CO labels, max marks, and student marks
    co_labels = df.iloc[0].tolist()
    max_marks = df.iloc[1].tolist()
    student_marks = df.iloc[2:].values.tolist()
    
    # Group COs and their corresponding marks
    co_groups = {}
    for co, max_mark in zip(co_labels, max_marks):
        if co not in co_groups:
            co_groups[co] = []
        co_groups[co].append(max_mark)
    
    # Calculate total marks for each CO group
    co_totals = {co: sum(marks) for co, marks in co_groups.items()}
    
    # Calculate scaling factor
    total_co_marks = sum(co_totals.values())
    scaling_factor = total_marks / total_co_marks
    
    # Scale CO totals
    scaled_co_totals = {co: total * scaling_factor for co, total in co_totals.items()}
    
    # Process student marks
    processed_student_marks = []
    for student in student_marks:
        student_co_marks = {}
        for co, marks in co_groups.items():
            co_indices = [i for i, label in enumerate(co_labels) if label == co]
            student_co_total = sum(student[i] for i in co_indices)
            max_co_total = sum(marks)
            scaled_co_total = scaled_co_totals[co]
            student_co_marks[co] = (student_co_total / max_co_total) * scaled_co_total
        processed_student_marks.append(student_co_marks)
    
    # Prepare output data
    output_data = []
    output_data.append(["CO"] + list(co_totals.keys()))
    output_data.append(["Max Marks"] + list(co_totals.values()))
    output_data.append(["Scaled Max Marks"] + [round(mark, 2) for mark in scaled_co_totals.values()])
    for i, student in enumerate(processed_student_marks, start=1):
        output_data.append([f"Student {i}"] + [round(mark, 2) for mark in student.values()])
    
    # Create output DataFrame
    output_df = pd.DataFrame(output_data[1:], columns=output_data[0])
    
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
    
    # Get total marks from user
    total_marks = st.number_input("Enter total marks to scale to:", min_value=1, value=50)
    
    if st.button("Process Data"):
        # Process the data
        output_df = process_co_data(df, total_marks)
        
        # Display output data
        st.subheader("Processed Data")
        st.dataframe(output_df)
        
        # Provide download link for processed data
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            output_df.to_excel(writer, index=False)
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
2. The file should have CO labels in the first row, max marks in the second row, and student marks in subsequent rows.
3. Enter the total marks you want to scale to.
4. Click 'Process Data' to see the results.
5. You can download the processed data as an Excel file.
""")
