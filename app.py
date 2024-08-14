import streamlit as st
import pandas as pd

# Title of the app
st.title("Teacher Data Processing App")

# Upload files
teacher_data_file = st.file_uploader("Upload the teacher data Excel file", type="xlsx")
criteria_file = st.file_uploader("Upload the subject mapping Excel file", type="xlsx")

if teacher_data_file and criteria_file:
    # Read files
    teacher_df = pd.read_excel(teacher_data_file)
    criteria_df = pd.read_excel(criteria_file)

    # Extract criteria information
    criteria_subjects = criteria_df[['subject', 'subsubjects', 'grades']].values.tolist()
    criteria_grades = criteria_df['grades'].tolist()

    teacher_df['Teacher Name'] = teacher_df['firstName']
    teacher_df['Employee Code'] = teacher_df['employeeCode']
    teacher_df['Highest Grade'] = 'NA'
    teacher_df['Subject'] = 'NA'

    def refined_subject_matches(subject, criteria_subjects, highest_grade):
        subject = subject.lower()
        matched_subjects = set()
        for main_subject, sub_subject, grade in criteria_subjects:
            main_subject = main_subject.lower()
            sub_subject = sub_subject.lower()
            if grade == highest_grade and ((main_subject == subject) or (sub_subject == subject)):
                matched_subjects.add(main_subject)
        return matched_subjects

    def find_highest_grade_subject_refined(row):
        highest_grade = 0
        main_subjects = set()

        for grade_col in range(1, 23):
            grade_col_name = f'grade{grade_col}'
            subject_col_name = f'subject{grade_col}'
            grade = row[grade_col_name]
            subject = row[subject_col_name]

            if pd.notna(grade) and pd.notna(subject):
                grade = int(grade)
                if grade in criteria_grades:
                    matched_subject = refined_subject_matches(subject, criteria_subjects, grade)
                    if matched_subject:
                        if grade > highest_grade:
                            highest_grade = grade
                            main_subjects = matched_subject
                        elif grade == highest_grade:
                            main_subjects.update(matched_subject)

        if highest_grade > 0:
            return highest_grade, ','.join(main_subjects)

        return highest_grade, 'NA'

    # Apply the function to the DataFrame
    for index, row in teacher_df.iterrows():
        highest_grade, main_subject = find_highest_grade_subject_refined(row)
        teacher_df.at[index, 'Highest Grade'] = highest_grade
        teacher_df.at[index, 'Subject'] = main_subject

    teacher_df['Subject'] = teacher_df['Subject'].apply(lambda x: ','.join(sorted(set(x.split(',')))))

    # Save the output file
    output_file_refined_v3 = 'VGOS_Teacher_Data.xlsx'
    teacher_df.to_excel(output_file_refined_v3, index=False)

    # Provide download link for the output file
    with open(output_file_refined_v3, "rb") as file:
        st.download_button(
            label="Download updated teacher data",
            data=file,
            file_name=output_file_refined_v3,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.success("Data processing complete. You can download the file now.")

else:
    st.info("Please upload both the teacher data file and criteria file to proceed.")
