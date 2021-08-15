import numpy as np
import pandas as pd
import statistics

from create_pdf import create_pdf


def create_dataset(excel_file_loc):

    try:
        dataset = pd.read_excel(excel_file_loc, header=1)
        read_file = True

    except:
        print("--- Unable to read the file ---")
        dataset = None
        read_file = False

    if read_file:
        dataset.columns = dataset.columns.str.strip()

        dataset['Attempt Status'] = np.NaN
        dataset.loc[dataset['Outcome (Correct/Incorrect/Not Attempted)'] == 'Unattempted', 'Attempt Status'] = 'Unattempted'
        dataset['Attempt Status'].replace(np.NaN, 'Attempted', inplace=True)

        dataset_info = dataset.copy()
        dataset_info.drop_duplicates(subset="Registration Number", keep='first', inplace=True)

        temp = pd.DataFrame(dataset.groupby('Registration Number')['Your score'].sum()).reset_index().rename(columns={'Your score': 'Total score'})
        dataset_info = pd.merge(dataset_info, temp, how='inner', on='Registration Number')

        dataset_info['World Percentile'] = dataset_info['Total score'].rank(pct=True)
        dataset_info['World Percentile'] = dataset_info['World Percentile'] * 100

        dataset_info['Country Percentile'] = dataset_info.groupby(['Country of Residence'])['Total score'].rank(pct=True).to_list()
        dataset_info['Country Percentile'] = dataset_info['Country Percentile'] * 100

        dataset_marksheet_section1 = dataset[['Registration Number', 'Question No.', 'Attempt Status',
                                              'What you marked', 'Correct Answer', 'Outcome (Correct/Incorrect/Not Attempted)',
                                              'Score if correct', 'Your score']].copy()

        dataset_marksheet_section2 = dataset[['Registration Number', 'Question No.', 'Attempt Status',
                                              'What you marked', 'Correct Answer', 'Outcome (Correct/Incorrect/Not Attempted)',
                                              'Your score', '% of students\nacross the world\nwho attempted\nthis question', '% of students (from\nthose who attempted\nthis ) who got it\ncorrect',
                                              '% of students\n(from those who\nattempted this)\nwho got it\nincorrect', 'World Average\nin this question']].copy()

        dataset_marksheet_section3 = dataset[['Registration Number', 'Question No.', 'Attempt Status',
                                              'What you marked', 'Correct Answer', 'Outcome (Correct/Incorrect/Not Attempted)', 'Your score']].copy()

        per_country = dataset[['Registration Number', 'Country of Residence', 'Question No.', 'Attempt Status',
                                              'What you marked', 'Correct Answer', 'Outcome (Correct/Incorrect/Not Attempted)', 'Your score']].copy()

        per_country['Per_attempts'] = np.NaN
        per_country['Per_correct'] = np.NaN
        per_country['Per_incorrect'] = np.NaN
        per_country['Average'] = np.NaN

        merger = []
        for country in per_country['Country of Residence'].unique().tolist():
            country_dataset = per_country[per_country['Country of Residence'] == country]
            per_attempts = []
            per_correct = []
            per_incorrect = []
            average = []
            for question in country_dataset['Question No.'].unique().tolist():
                question_dataset = country_dataset[country_dataset['Question No.'] == question]
                number_attempted = (question_dataset[question_dataset['Attempt Status'] == 'Attempted']['Attempt Status'].count() / question_dataset['Attempt Status'].count()) * 100
                attempt_dataset = question_dataset[question_dataset['Attempt Status'] == 'Attempted']
                correct_attempts = (attempt_dataset[attempt_dataset['Outcome (Correct/Incorrect/Not Attempted)'] == 'Correct']['Outcome (Correct/Incorrect/Not Attempted)'].count() / attempt_dataset[
                    'Outcome (Correct/Incorrect/Not Attempted)'].count()) * 100
                incorrect_attempts = (attempt_dataset[attempt_dataset['Outcome (Correct/Incorrect/Not Attempted)'] == 'Incorrect']['Outcome (Correct/Incorrect/Not Attempted)'].count() /
                                      attempt_dataset['Outcome (Correct/Incorrect/Not Attempted)'].count()) * 100
                avg_score = statistics.mean(question_dataset['Your score'].to_list())
                per_attempts.append(number_attempted)
                per_correct.append(correct_attempts)
                per_incorrect.append(incorrect_attempts)
                average.append(avg_score)
            for reg_no in country_dataset['Registration Number'].unique().tolist():
                new_df = country_dataset[country_dataset['Registration Number'] == reg_no].copy()
                new_df['Per_attempts'] = per_attempts
                new_df['Per_correct'] = per_correct
                new_df['Per_incorrect'] = per_incorrect
                new_df['Average'] = average
                new_df['Per_attempts'].replace(np.NaN, 0, inplace=True)
                new_df['Per_correct'].replace(np.NaN, 0, inplace=True)
                new_df['Per_incorrect'].replace(np.NaN, 0, inplace=True)
                new_df['Average'].replace(np.NaN, 0, inplace=True)
                merger.append(new_df)

        dataset_marksheet_section3 = pd.concat(merger)

        for reg_no in dataset_info['Registration Number'].unique():

            data_on_condition = dataset_info['Registration Number'] == reg_no
            info = dataset_info[data_on_condition]

            dataset_marksheet_section1_student = dataset_marksheet_section1[dataset_marksheet_section1['Registration Number'] == reg_no].copy()
            dataset_marksheet_section1_student.drop(columns=['Registration Number'], inplace=True)

            dataset_marksheet_section2_student = dataset_marksheet_section2[dataset_marksheet_section2['Registration Number'] == reg_no].copy()
            dataset_marksheet_section2_student.drop(columns=['Registration Number'], inplace=True)

            dataset_marksheet_section3_student = dataset_marksheet_section3[dataset_marksheet_section3['Registration Number'] == reg_no].copy()
            dataset_marksheet_section3_student.drop(columns=['Registration Number'], inplace=True)

            img_loc = 'Photo/' + info['Full Name'].item() + '.png'
            create_pdf(img_loc, info, dataset, dataset_info, dataset_marksheet_section1_student, dataset_marksheet_section2_student, dataset_marksheet_section3_student)

