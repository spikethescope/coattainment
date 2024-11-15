
import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import pandas as pd
import re
from streamlit_echarts import st_echarts


def compute_attainment_both_options(output_df, attain_level_3_min=80,attain_level_2_min=60,attain_level_1_min=50,threshold=0.6, method="threshold"):
    # Extract the CO labels and weighted max marks from the output DataFrame
    co_labels = output_df.iloc[0, 1:].tolist()  # Skipping the first column (CO label row)
    weighted_max_marks = output_df.iloc[1, 1:].tolist()  # Extracting weighted max marks

    # Create a dictionary for max marks per CO
    max_marks = {co: max_mark for co, max_mark in zip(co_labels, weighted_max_marks)}

    # Define expected proficiency thresholds based on the user-specified threshold or average method
    if method == "threshold":
        # Use threshold-based attainment
        thresholds = {co: threshold * max_marks[co] for co in max_marks}
    elif method == "average":
        # Use average score of students for each CO as the threshold
        student_marks_df = output_df.iloc[2:, 1:]  # Skip the first column with student names
        student_marks_df.columns = co_labels  # Set column names to CO labels
        thresholds = student_marks_df.mean().to_dict()  # Set thresholds to averages per CO
    else:
        raise ValueError("Invalid method. Choose either 'threshold' or 'average'.")

    # Extract student marks rows, starting from row 2 (excluding headers)
    student_marks_df = output_df.iloc[2:, 1:]  # Skip the first column with student names
    student_marks_df.columns = co_labels  # Set column names to CO labels
    total_students = len(student_marks_df)

    # Count students who met or exceeded the expected proficiencies for each CO
    results = {}
    for co, min_score in thresholds.items():
        results[co] = (student_marks_df[co] >= min_score).sum()

    # Calculate Course Outcome attainment percentages
    attainment_percentages = {co: (count / total_students) * 100 for co, count in results.items()}

    # Determine CO attainment levels
    attainment_levels = {}
    for co, percentage in attainment_percentages.items():
        if percentage >= attain_level_3_min:
            attainment_levels[co] = 3
        elif percentage >= attain_level_2_min:
            attainment_levels[co] = 2
        elif percentage >= attain_level_1_min:
            attainment_levels[co] = 1
        else:
            attainment_levels[co] = 0

    # Prepare a summary of outcomes
    summary = {
        'CO': co_labels,
        'Expected Proficiency (%)': [threshold * 100 if method == "threshold" else "Average Score"] * len(max_marks),
        'No of Students Scored Expected Marks': [results[co] for co in co_labels],
        'Course Outcome Attainment (%)': [attainment_percentages[co] for co in co_labels],
        'CO Attainment Level': [attainment_levels[co] for co in co_labels]
    }

    # Create a DataFrame for the summary
    summary_df = pd.DataFrame(summary)
    
    return summary_df

   
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
    
    # Get unique CO components while maintaining order
    # Get unique CO components and sort them in 'CO1', 'CO2', etc. order
    unique_cos = sorted({col for col in co_labels if isinstance(col, str) and col.startswith('CO')},
                        key=lambda x: int(x[2:]))  # Sort by numeric part of CO labels
    
    #unique_cos = []
    #seen = set()
    #for co in co_labels:
     #   if isinstance(co, str) and co.startswith('CO') and co not in seen:
      #      unique_cos.append(co)
       #     seen.add(co)
    
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
    
    # Prepare output data with unique CO labels
    output_data = headers + [
        ["CO"] + unique_cos,
        ["Weighted Max Marks"] + [weighted_max_marks.get(co, "") for co in unique_cos]
    ]
    for i, student in enumerate(processed_student_marks, start=1):
        output_data.append([f"Student {i}"] + [student.get(co, "") for co in unique_cos])
    
    # Create output DataFrame
    output_df = pd.DataFrame(output_data)
    
    return output_df

# Initialize session state variables
if 'processed' not in st.session_state:
    st.session_state.processed = False
if 'output_df' not in st.session_state:
    st.session_state.output_df = None
if 'summary_df' not in st.session_state:
    st.session_state.summary_df = None
# Create the data for the DataFrame
data = {
    'CO2': [10, 9, 9, 9],
    'CO1': [10, 8, 9, 4],
    'CO2': [10, 7, 9, 9],
    'CO2': [10, 7, 8, 7],
    'CO3': [10, 8, 9, 8],
    'CO4': [10, 8, 8, 7],
    'CO4': [10, 10, 9, 6]
}
sample_df = pd.DataFrame(data)
# Streamlit app
st.title('Course Outcome (CO) Attainment Computation')
st.subheader("Excel input format below. Note second row is Max. Marks for that Assessment.")
# Display the DataFrame in Streamlit
st.write(sample_df)

# File upload
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    # Read the Excel file
    df = pd.read_excel(uploaded_file, header=None)
    
    # Display input data
    st.subheader("Input Data")
    st.dataframe(df)
    attain_level_3_min = st.number_input(
                        "Enter minimum value for Attainment Level 3:",
                        min_value=0,
                        max_value=100,
                        value=80
                    )
    attain_level_2_min = st.number_input(
                        "Enter minimum value for Attainment Level 2:",
                        min_value=0,
                        max_value=attain_level_3_min,
                        value=60
                    )
    attain_level_1_min = st.number_input(
                        "Enter minimum value for Attainment Level 1:",
                        min_value=0,
                        max_value=attain_level_2_min,
                        value=50
                    )
    # Find CO labels
    co_row = next((i for i, row in df.iterrows() if any(co in str(cell) for co in ['CO1', 'CO2', 'CO3', 'CO4', 'CO5', 'CO6'] for cell in row)), None)
    if co_row is None:
        st.error("No CO labels found in the file")
    else:
        co_labels = list({col for col in df.iloc[co_row] if isinstance(col, str) and col.startswith('CO')})

        # Create input fields for CO weights
        #st.write(f"The list of CO labels are {co_labels}")
        st.subheader("Enter weightage for each CO component. Weightage is the weight of the CO component relative to the entire course. Thus the sum of all weights should be 1.")      
        
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
            if st.button("Process Data") or st.session_state.processed:
                # Process the data
                st.session_state.output_df = process_co_data(df, co_weights, round_digits)
                st.session_state.processed = True
                
                # Display output data
                st.subheader("Processed Data")
                st.write("Download Scaled Marks for Each CO")
                st.dataframe(st.session_state.output_df)
                
                # Provide download link for processed data
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    st.session_state.output_df.to_excel(writer, index=False, header=False)
                output.seek(0)
                
                st.download_button(
                    label="Download Processed Data",
                    data=output,
                    file_name="processed_co_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Attainment calculation section
                st.subheader("Course Outcome Attainment Analysis")

                # Attainment method selection
                method = st.radio("Choose attainment calculation method:", ("Threshold", "Average"))

                # Set threshold value input only if "Threshold" method is chosen
                threshold = None
                if method == "Threshold":
                    threshold = st.number_input(
                        "Enter Proficiency Threshold (as a decimal, e.g., 0.60 for 60%):",
                        min_value=0.0,
                        max_value=1.0,
                        value=0.60
                    )

                # Compute attainment based on user selection
                if st.button("Calculate Attainment"):
                    try:
                        summary_df = compute_attainment_both_options(
                            st.session_state.output_df,                                                       
                            attain_level_3_min,
                            attain_level_2_min,
                            attain_level_1_min,
                            threshold=threshold, 
                            method=method.lower()
                        )
                       
                        
                        # Display the summary DataFrame
                        st.write("### Summary of Course Outcomes")
                        st.write(f"### Attainment Levels: 0 if less than {attain_level_1_min}% of students attains expected proficiency, 1 if between {attain_level_1_min} and {attain_level_2_min}% Students attains expected proficiency , 2 if between {attain_level_2_min}% and {attain_level_3_min}% students attain expected proficiency,3 if more than {attain_level_3_min}% students achieve the expected proficiency.") 
                                                
                        st.dataframe(summary_df)

                        # Download the summary as an Excel file
                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                            summary_df.to_excel(writer, index=False, sheet_name='Summary')
                        buffer.seek(0)  # Move to the beginning of the BytesIO buffer

                        st.download_button(
                            label="Download Summary as Excel",
                            data=buffer,
                            file_name="course_outcome_summary.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.session_state.summary_df = summary_df
                    except ValueError as e:
                        st.error(f"Error: {e}")
                if st.button("Draw Histogram"):
                    try:
                        st.write(st.session_state.summary_df)
                        df = st.session_state.summary_df                                                       
                         
                        st.write("### Data Preview")
                        
                        st.dataframe(df)
                        
                        # Prepare data for ECharts
                        categories = df["CO"].tolist()
                        scores = df["CO Attainment Level"].tolist()
                        percentages = df["Course Outcome Attainment (%)"].tolist()
                        
                        # Options for ECharts
                        options = {
                            "title": {"text": "Category Histogram", "left": "center"},
                            "tooltip": {"trigger": "axis"},
                            "xAxis": {
                                "type": "category",
                                "data": categories,
                                "name": "Categories",
                            },
                            "yAxis": {
                                "type": "value",
                                "name": "Values",
                            },
                            "series": [
                                {
                                    "name": "Average Score",
                                    "type": "bar",
                                    "data": scores,
                                    "label": {"show": True, "position": "top"},
                                },
                                {
                                    "name": "Percentage",
                                    "type": "bar",
                                    "data": percentages,
                                    "label": {"show": True, "position": "top"},
                                },
                            ],
                            "legend": {"data": ["Average Score", "Percentage"], "top": "10%"},
                        }
                        
                        # Display the chart
                        st.write("### Histogram")
                        st_echarts(options=options, height="400px")
                    except ValueError as e:
                        st.error(f"Error: {e}")    

st.markdown("""
### Instructions:
1. Upload an Excel file containing CO data.
2. The file should have CO labels in a row, with max marks in the row below, and student marks in subsequent rows. Any headers above the CO labels will be included in the output.
3. Enter the Attainment Level Percentages and then enter the weightage for each CO component. The sum of all weights should be 1.
4. Specify the number of decimal places for rounding student marks.
5. Click 'Process Data' to see the results.
6. You can download the processed data as an Excel file.
""")
