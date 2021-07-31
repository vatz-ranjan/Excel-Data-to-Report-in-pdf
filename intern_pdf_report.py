'''
    *******************************************************************************
    *******************************************************************************

    This program will prepare the report cards of student based on their data...

    *******************************************************************************
    *******************************************************************************
'''


# All the important modules to be imported
import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle, Image, PageBreak, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle


# This function will analyze the given data of students...
def excel_data():

    plot_graph = True
    make_pdf = True

    try:
        dataset = pd.read_excel('Dummy Data.xlsx', header=1)
        excel_read = True

    except():
        print("--- Unable to read the excel file --- /n--- Provide the correct path ---")
        excel_read = False

    if excel_read:
        dataset.columns = dataset.columns.str.strip()
        dataset_info = dataset[
            ['Registration Number', 'Full Name', 'Grade', 'Round', 'Date and time of test', 'Name of School', 'Gender',
             'Date of Birth', 'City of Residence', 'Country of Residence', 'Final result']].copy()
        dataset_info.drop_duplicates(subset="Registration Number", keep='first', inplace=True)
        dataset_score = pd.DataFrame(dataset.groupby('Registration Number')[['Your score', 'Score if correct']].sum())
        dataset_score.reset_index(inplace=True)
        dataset_report = pd.merge(dataset_info, dataset_score, how='inner', on='Registration Number')
        dataset_report['Percentage'] = dataset_report['Your score'] / dataset_report['Score if correct'] * 100
        dataset_report['World Rank'] = dataset_report['Percentage'].rank(ascending=False)
        dataset_report['Country Rank'] = dataset_report.groupby(['Country of Residence'])['Percentage'].rank(ascending=False).to_list()
        dataset_report['School Rank'] = dataset_report.groupby(['Name of School'])['Percentage'].rank(ascending=False).to_list()

        info = {}
        for reg_no in dataset_report['Registration Number'].unique():

            data_on_condition = dataset_report['Registration Number'] == reg_no
            info['Registration Number'] = reg_no
            info['Name'] = dataset_report[data_on_condition]['Full Name'].item()
            info['Grade'] = dataset_report[data_on_condition]['Grade'].item()
            info['Round'] = dataset_report[data_on_condition]['Round'].item()
            info['Date and time of test'] = dataset_report[data_on_condition]['Date and time of test'].item()
            info['Name of School'] = dataset_report[data_on_condition]['Name of School'].item()
            info['Gender'] = dataset_report[data_on_condition]['Gender'].item()
            info['Date of Birth'] = dataset_report[data_on_condition]['Date of Birth'].item().strftime('%b %d, %Y')
            info['City of Residence'] = dataset_report[data_on_condition]['City of Residence'].item()
            info['Country of Residence'] = dataset_report[data_on_condition]['Country of Residence'].item()
            info['Final result'] = dataset_report[data_on_condition]['Final result'].item()
            info['Your score'] = dataset_report[data_on_condition]['Your score'].item()
            info['Score if correct'] = dataset_report[data_on_condition]['Score if correct'].item()
            info['Percentage'] = dataset_report[data_on_condition]['Percentage'].item()
            info['World Rank'] = dataset_report[data_on_condition]['World Rank'].item()
            info['Country Rank'] = dataset_report[data_on_condition]['Country Rank'].item()
            info['School Rank'] = dataset_report[data_on_condition]['School Rank'].item()

            if plot_graph:
                create_graph(dataset_report, info)

            if make_pdf:
                create_pdf(info)


# This program will create several graphs...
def create_graph(dataset_report, info):

    # os.makedirs('Graph/', exist_ok=True)

    width = 7
    height = 6

    # World Rank
    fig1, ax1 = plt.subplots(figsize=(width, height))
    plt.ylim(0, 100)

    ax1.bar(dataset_report['Full Name'],
            dataset_report['Your score'])
    ax1.set_xlabel("Full Name")
    ax1.set_ylabel("Score")
    ax1.set_title("World's Graph")


    folder_name = 'Graph/All/'
    os.makedirs(folder_name, exist_ok=True)
    fig_name = folder_name + str(info['Registration Number']) + ".png"
    fig1.savefig(fig_name)

    # Country Rank
    dataset_temp = dataset_report[dataset_report['Country of Residence'] == info['Country of Residence']]
    fig2, ax2 = plt.subplots(figsize=(width, height))
    plt.ylim(0, 100)

    ax2.bar(dataset_temp['Full Name'],
            dataset_temp['Your score'])
    ax2.set_xlabel("Full Name")
    ax2.set_ylabel("Score")
    ax2.set_title("Country's Graph")

    folder_name = 'Graph/Country/' + info['Country of Residence'] + '/'
    os.makedirs(folder_name, exist_ok=True)
    fig_name = folder_name + str(info['Registration Number']) + ".png"
    fig2.savefig(fig_name)

    # School Rank
    dataset_temp = dataset_report[dataset_report['Name of School'] == info['Name of School']]
    fig3, ax3 = plt.subplots(figsize=(width, height))
    plt.ylim(0, 100)

    ax3.bar(dataset_temp['Full Name'],
            dataset_temp['Your score'])

    ax3.set_xlabel("Full Name")
    ax3.set_ylabel("Score")
    ax3.set_title("School's Graph")

    folder_name = 'Graph/School/' + info['Name of School'] + '/'
    os.makedirs(folder_name, exist_ok=True)
    fig_name = folder_name + str(info['Registration Number']) + ".png"
    fig3.savefig(fig_name)


# This program will start preparing the reports of the students in pdf form...
def create_pdf(info, cover_page=True):

    school_name = '{}'.format(info['Name of School'])

    name = 'Full Name: {}'.format(info['Name'])
    reg_no = 'Registration Number: {}'.format(info['Registration Number'])
    grade = 'Grade: {}'.format(info['Grade'])
    round_no = 'Round: {}'.format(info['Round'])
    date_time = 'Date and time of test: {}'.format(info['Date and time of test'])
    gender = 'Gender: {}'.format(info['Gender'])
    dob = 'Date of Birth: {}'.format(info['Date of Birth'])
    city = 'City of Residence: {}'.format(info['City of Residence'])
    country = 'Country of Residence: {}'.format(info['Country of Residence'])

    marks_obtained = info['Your score']
    total_marks = info['Score if correct']
    percentage = info['Percentage']
    world_rank = info['World Rank']
    country_rank = info['Country Rank']
    school_rank = info['School Rank']

    folder_name = 'Report/'
    os.makedirs(folder_name, exist_ok=True)

    pdf_name = folder_name + str(info['Registration Number']) + ".pdf"
    pdf = SimpleDocTemplate(pdf_name)
    pdf_obj = []
    styles1 = getSampleStyleSheet()

    style_title = styles1["Title"]
    style_title.fontName = 'Times'
    style_title.fontSize = 25

    school_logo_location = 'Photo/school_logo.jfif'
    title = Paragraph(text="REPORT CARD ", style=style_title)
    school_logo = Image(school_logo_location, 500, 500)
    subtitle = Paragraph(text=school_name.upper(), style=style_title)
    pdf_obj.append(title)
    pdf_obj.append(Spacer(0, 20))
    pdf_obj.append(school_logo)
    pdf_obj.append(Spacer(0, 20))
    pdf_obj.append(subtitle)

    pdf_obj.append(PageBreak())

    style_body = styles1["Normal"]
    style_body.fontName = 'Times'
    style_body.fontSize = 10

    name_text = Paragraph(text=name, style=style_body)
    regno_text = Paragraph(text=reg_no, style=style_body)
    gender_text = Paragraph(text=gender, style=style_body)
    dob_text = Paragraph(text=dob, style=style_body)
    grade_text = Paragraph(text=grade, style=style_body)
    school_text = Paragraph(text='School: {}'.format(school_name), style=style_body)
    round_text = Paragraph(text=round_no, style=style_body)
    datetime_text = Paragraph(text=date_time, style=style_body)
    city_text = Paragraph(text=city, style=style_body)
    country_text = Paragraph(text=country, style=style_body)

    student_photo_location = 'Photo/' + info['Name'] + '.png'

    t1_info = Table([[name_text, regno_text],
                [gender_text, dob_text],
                [grade_text, school_text],
                [round_text, datetime_text],
                [city_text, country_text]])

    t1_style = TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER")])
    t1_info.setStyle(t1_style)

    t1 = Table([[t1_info, Image(student_photo_location, 100, 100)]], colWidths=[400, 100])

    pdf_obj.append(t1)

    pdf_obj.append(Spacer(0, 30))

    marks_obtained_text = Paragraph(text="Marks Obtained", style=style_body)
    total_marks_text = Paragraph(text="Total Marks", style=style_body)
    percentage_text = Paragraph(text="Percentage", style=style_body)
    marks_obtained_text_no = Paragraph(text="{}".format(marks_obtained), style=style_body)
    total_marks_text_no = Paragraph(text="{}".format(total_marks), style=style_body)
    percentage_text_no = Paragraph(text="{}".format(percentage), style=style_body)

    t2 = Table([[marks_obtained_text, total_marks_text, percentage_text],
                [marks_obtained_text_no, total_marks_text_no, percentage_text_no]])

    t2_style = TableStyle([("BOX", (0, 0), (-1, -1), 1, colors.black),
                          ("GRID", (0, 0), (-1, -1), 1, colors.black),
                          ("ALIGN", (0, 0), (-1, -1), "CENTER")])

    t2.setStyle(t2_style)
    pdf_obj.append(t2)

    pdf_obj.append(Spacer(0, 30))

    world_rank_text = Paragraph(text="World's Rank", style=style_body)
    country_rank_text = Paragraph(text="Country's Rank", style=style_body)
    school_rank_text = Paragraph(text="School's Rank", style=style_body)
    world_rank_text_no = Paragraph(text="{}".format(world_rank), style=style_body)
    country_rank_text_no = Paragraph(text="{}".format(country_rank), style=style_body)
    school_rank_text_no = Paragraph(text="{}".format(school_rank), style=style_body)

    world_rank_graph_location = 'Graph/All/' + str(info['Registration Number']) + ".png"
    country_rank_graph_location = 'Graph/Country/' + info['Country of Residence'] + '/' + str(info['Registration Number']) + ".png"
    school_rank_graph_location = 'Graph/School/' + info['Name of School'] + '/' + str(info['Registration Number']) + ".png"

    world_rank_graph = Image(world_rank_graph_location, 125, 125)
    country_rank_graph = Image(country_rank_graph_location, 125, 125)
    school_rank_graph = Image(school_rank_graph_location, 125, 125)

    t3 = Table([[world_rank_text, country_rank_text, school_rank_text],
                [world_rank_text_no, country_rank_text_no, school_rank_text_no],
                [world_rank_graph, country_rank_graph, school_rank_graph]])

    t3_style = TableStyle([("BOX", (0, 0), (-1, -1), 1, colors.black),
                           ("GRID", (0, 0), (-1, -2), 1, colors.black),
                           ("ALIGN", (0, 0), (-1, -1), "CENTER")])

    t3.setStyle(t3_style)
    pdf_obj.append(t3)

    pdf.build(pdf_obj)


# Main Function....
if __name__ == '__main__':
    excel_data()


'''
    *******************************************************************************
    *******************************************************************************
    
    ---------------------------- END ----------------------------------------------
    
    *******************************************************************************
    *******************************************************************************
'''