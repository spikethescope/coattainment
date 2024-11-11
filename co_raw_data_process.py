import streamlit as st
import pandas as pd
import numpy as np
import io

def process_co_data(df, co_weights):
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
    
    # Calculate total of all CO marks
    total_co_marks = sum(co_totals.values())
    
    # Calculate weighted max marks for each CO
    weighted_max_marks = {co: total_co_marks * co_weights[co] for co in co_totals.keys()}
    
    # Process student marks
    processed_student_marks = []
    for student in student_marks:
        student_co_marks = {}
        for co, marks in co_groups.items():
            co_indices = [i for i, label in enumerate(co_labels) if label == co]
            student_co_total = sum(student[i] for i in co_indices)
            max_co_total = sum(marks)
            student_co_marks[co] = (student_co_total / max_co_total) * weighted_max_marks[co]
        processed_student_marks.append(student_co_marks)
    
    # Prepare output data
    output_data = []
    output_data.append(["CO"] + list(co_totals.keys()))
    output_data.append(["Weighted Max Marks"] + [round(weighted_max_marks[co], 2) for co in co_totals.keys()])
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
    
    # Get unique CO components
    co_components = sorted(set(df.iloc[0]))
    
    # Create input fields for CO weights
    st.subheader("Enter weightage for each CO component")
    st.write("The sum of all weights should be 1.")
    
    co_weights = {}
    col1, col2 = st.columns(2)
    for i, co in enumerate(co_components):
        if i % 2 == 0:
            co_weights[co] = col1.number_input(f"Weight for {co}", min_value=0.0, max_value=1.0, value=1.0/len(co_components), step=0.01, format="%.2f")
        else:
            co_weights[co] = col2.number_input(f"Weight for {co}", min_value=0.0, max_value=1.0, value=1.0/len(co_components), step=0.01, format="%.2f")
    
    # Check if weights sum to 1
    total_weight = sum(co_weights.values())
    if abs(total_weight - 1.0) > 1e-6:  # Allow for small floating-point errors
        st.warning(f"The sum of weights is {total_weight:.2f}. It should be 1.0.")
    else:
        if st.button("Process Data"):
            # Process the data
            output_df = process_co_data(df, co_weights)
            
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
3. Enter the weightage for each CO component. The sum of all weights should be 1.
4. Click 'Process Data' to see the results.
5. You can download the processed data as an Excel file.
""")
