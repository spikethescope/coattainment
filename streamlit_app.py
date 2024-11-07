import streamlit as st
import pandas as pd
import re
import io

# Sample data
data = {
    "NAME": ["John", "Mathew", "Rema", "Anand", "Yuan"],
    "ROLL NUMBER": ["F22003", "F22005", "F22011", "F22017", "F22018"],
    "CO1 (30)": [24, 16, 22, 24, 24],
    "CO2 (40)": [25, 24, 21, 25, 24],
    "CO3 (20)": [15, 16, 18, 19, 19],
    "CO4 (10)": [6, 7, 5, 6, 2]
}

# Create a DataFrame
df_students = pd.DataFrame(data)


# Function to process the uploaded Excel file
def process_file(uploaded_file, threshold):
    # Load data from the uploaded Excel file
    df = pd.read_excel(uploaded_file)

    # Extract maximum marks for each CO from the column headers
    max_marks = {}
    for column in df.columns:
        if 'CO' in column:
            match = re.search(r'\((\d+)\)', column)
            if match:
                max_marks[column] = int(match.group(1))

    # Define expected proficiency thresholds based on user input
    thresholds = {co: threshold * max_marks[co] for co in max_marks}

    # Count number of students
    total_students = len(df)

    # Count students who met or exceeded expected proficiencies
    results = {}
    for co, min_score in thresholds.items():
        results[co] = (df[co] >= min_score).sum()

    # Calculate Course Outcome attainment percentages
    attainment_percentages = {co: (count / total_students) * 100 for co, count in results.items()}

    # Determine CO attainment levels
    attainment_levels = {}
    for co, percentage in attainment_percentages.items():
        if percentage >= 80:
            attainment_levels[co] = 3
        elif percentage >= 70:
            attainment_levels[co] = 2
        else:
            attainment_levels[co] = 1

    # Summary of outcomes
    summary = {
        'CO': list(max_marks.keys()),
        'Expected Proficiency (%)': [threshold * 100] * len(max_marks),
        'No of Students Scored Expected Marks': [results[co] for co in max_marks],
        'Course Outcome Attainment (%)': [attainment_percentages[co] for co in max_marks],
        'CO Attainment Level': [attainment_levels[co] for co in max_marks]
    }

    # Create a DataFrame for summary
    summary_df = pd.DataFrame(summary)
    
    return summary_df

# Streamlit app layout
st.title("Course Outcome Attainment Analysis")
# Display the DataFrame as a table
st.write("### Student Scores Sample format which contains 4 COs. If you require more components add extra columns but follow same format")
st.write("Please follow it exactly. Maximum marks is specified in brackets. For example CO1 (30) means Max. marks in an assessment is 30.")

st.dataframe(df_students)
st.write("### CO Details.")
# Input for proficiency threshold
threshold = st.number_input("Enter Proficiency Threshold (as a decimal, e.g., 0.60 for 60%) Suggested is 60% threshold:", min_value=0.0, max_value=1.0, value=0.60)

# File uploader
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

# Process and display results when the button is clicked
if st.button("Process File"):
    if uploaded_file is not None:
        summary_df = process_file(uploaded_file, threshold)
        
        # Display the summary DataFrame
        st.write("### Summary of Course Outcomes:")
        st.dataframe(summary_df)

        # Prepare to download the summary as an Excel file
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
    else:
        st.error("Please upload an Excel file.")
