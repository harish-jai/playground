import pandas as pd

def analyze_courses(meng_data, mcs_data):
    """
    Identifies common course codes between two datasets and extracts relevant information.

    Args:
        meng_data (list): List of dictionaries representing the MEng dataset.
        mcs_data (list): List of dictionaries representing the MCS dataset.

    Returns:
        dict: A dictionary where keys are common course codes and values are dictionaries
              containing the extracted course information.
    """

    common_courses_info = {}

    # Create dictionaries for faster lookup
    mcs_data_dict = {course['Course Number']: course for course in mcs_data}

    for meng_course in meng_data:
        meng_course_code = meng_course['Course Number & Name'].split('.')[0]  # Extract course code
        if meng_course_code in mcs_data_dict:
            mcs_course = mcs_data_dict[meng_course_code]

            # Determine MEng program areas and their status (M or E)
            meng_areas = {}
            for area in ['CM', 'CT', 'CV', 'EM', 'ES', 'IC', 'MM', 'OP', 'PE', 'RO', 'SP', 'SS']:
                if area in meng_course and meng_course[area] and not pd.isna(meng_course[area]):
                    meng_areas[area] = meng_course[area]

            # Apply filters
            is_cse_program_qualifying = mcs_course[' Graduate Credit Type'] == 'CSE Program Qualifying'
            has_meng_program_area_m = any(val == 'M' for val in meng_areas.values())

            if is_cse_program_qualifying or has_meng_program_area_m:
                common_courses_info[meng_course_code] = {
                    'Course Title': mcs_course[' Course Name'],
                    'Credits': mcs_course['Credit Value'],
                    'CSE Program Qualifying': 'Yes' if is_cse_program_qualifying else 'No',
                    'MCS Breadth Area': mcs_course['Breadth Area'],
                    'CSE Technical Elective': 'Yes' if isinstance(mcs_course['Technical Elective'], str) and '•' in mcs_course['Technical Elective'] else 'No',
                    'MEng Program Areas': meng_areas
                }

    # Order the results by Course Code
    ordered_courses_info = dict(sorted(common_courses_info.items()))

    return ordered_courses_info


# Load data from Excel files
meng_data_file = './ECE.xlsx'
mcs_data_file = './CSE.xlsx'

# Read the Excel files into pandas DataFrames
meng_data_df = pd.read_excel(meng_data_file)
mcs_data_df = pd.read_excel(mcs_data_file)

# Convert the DataFrames to lists of dictionaries
meng_data = meng_data_df.to_dict(orient='records')
mcs_data = mcs_data_df.to_dict(orient='records')

print(mcs_data_df.columns)
print(meng_data_df.columns)

common_course_data = analyze_courses(meng_data, mcs_data)

# Display the results
for code, info in common_course_data.items():
    print(f"Course Code: {code}")
    for key, value in info.items():
        print(f"  {key}: {value}")
    print("-" * 20)

print("Total common courses found:", len(common_course_data))