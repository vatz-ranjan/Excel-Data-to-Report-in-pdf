import os
import numpy as np
import pandas as pd
import statistics
import matplotlib.pyplot as plt

from static import PAGE_WIDTH, PAGE_HEIGHT, MARGIN_LEFT, MARGIN_TOP
from reportlab.platypus import Paragraph, Image, Table
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from utils import PDFItem, PDFPage, PDF


PAGE_TOP_2 = 925


def get_page_1(img_loc, info, dataset_marksheet_section1_student):
    page_1 = PDFPage()
    background = Image("back.png", width=PAGE_WIDTH, height=PAGE_HEIGHT, hAlign='CENTER')
    title = Paragraph("<font size=12><b>INTERNATIONAL MATHS OLYMPIAD CHALLENGE</b></font>")
    logo = Image("logo.png", width=7 * cm + 20, height=3.4 * cm, hAlign='LEFT')
    headingl = Paragraph(
        text=f"<b>Round I - Enhanced Score Report</b>:  {info['Full Name'].item()}<br/>\nReg Number: {info['Registration Number'].item()}")
    stu_pic = Image(f"{img_loc}", width=5 * cm, height=4 * cm, hAlign="RIGHT")
    student_name = Paragraph(text=f"<b><font size=12>Round I performance of {info['Full Name'].item()}</font></b>", )
    student_detail = (
        ("Grade", f"{info['Grade'].item()}", "", "Registration No. ", f"{info['Registration Number'].item()}"),
        ("School Name", f"{info['Name of School'].item()}", "", "Gender", f"{info['Gender'].item()}"),
        ("City Of Residence", f"{info['City of Residence'].item()}", "", "Date of Birth",
         f"{info['Date of Birth'].item().strftime('%b %d, %Y')}"),
        ("Country Of Residence", f"{info['Country of Residence'].item()}", "", "Date Of Test",
         f"{info['Date and time of test'].item()}")
    )
    tbl_style = (
        ('FONT', (0, 0), (0, -1), "Helvetica-Bold"),
        ('FONT', (3, 0), (3, -1), "Helvetica-Bold"),
        ('INNERGRID', (0, 0), (1, -1), 1, (0, 0, 0)),
        ('INNERGRID', (-2, 0), (-1, -1), 1, (0, 0, 0)),
        ('BOX', (0, 0), (1, -1), 1, (0, 0, 0)),
        ('BOX', (-2, 0), (-1, -1), 1, (0, 0, 0)),
    )
    table = Table(student_detail, style=tbl_style, colWidths=(None, 5 * cm, 1 * cm, None, 5 * cm))

    styles1 = getSampleStyleSheet()

    heading = styles1["Normal"]
    heading.fontName = 'Times'
    heading.fontSize = 15

    stf = Paragraph(text="Section 1", style=heading)
    desc = Paragraph(text=f"This section describes {info['First Name'].item()}'s performance v/s the Test in Grade {info['Grade'].item()}")
    report_table_data = [
        ("Question No.", 'Attempt\nStatus', f"{info['First Name'].item()}'s\nChoice", 'Correct\nAnswer',
         'Outcome', 'Score if\ncorrect', f"{info['First Name'].item()}'s\nScore"),
    ]

    for _, row in dataset_marksheet_section1_student.iterrows():
        report_table_data.append((
            row['Question No.'], row['Attempt Status'], row['What you marked'], row['Correct Answer'],
            row['Outcome (Correct/Incorrect/Not Attempted)'],
            row['Score if correct'], row['Your score']
        ))
    style = (
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
        ('BACKGROUND', (0, 0), (-1, 0), (0, 0, 0)),
        ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('TEXTCOLOR', (0, 0), (-1, 0), (1, 1, 1)),
        ('BOX', (0, 0), (-1, -1), 0.25, (0, 0, 0)),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, (0, 0, 0))
    )
    table1 = Table(report_table_data, style=style)
    tscore = Paragraph(text=f"<b><i>Total Score</i>: {info['Total score'].item()}</b>")
    data = (
        PDFItem(background, 0, 0),
        PDFItem(title, 245, PAGE_TOP_2 - 1 * cm - 100),
        PDFItem(headingl, MARGIN_LEFT + 6, PAGE_TOP_2 - 100),
        PDFItem(logo, 270, PAGE_TOP_2 - 4.5 * cm - 100),
        PDFItem(student_name, MARGIN_LEFT + 250, PAGE_TOP_2 - (4 * cm + 120)),
        PDFItem(stu_pic, PAGE_WIDTH - 220, PAGE_TOP_2 - 120 - 4 * cm),
        PDFItem(table, MARGIN_LEFT + 110, PAGE_TOP_2 - 330),
        PDFItem(stf, MARGIN_LEFT + 350, PAGE_TOP_2 - 350),
        PDFItem(desc, MARGIN_LEFT + 230, PAGE_TOP_2 - 370),
        PDFItem(table1, MARGIN_LEFT + 160, 50),
        PDFItem(tscore, PAGE_WIDTH - 327, 30),
    )

    page_1.add(data)
    return page_1


def get_page_2(info, dataset_marksheet_section2_student):
    page_2a = PDFPage()

    styles1 = getSampleStyleSheet()

    heading = styles1["Normal"]
    heading.fontName = 'Times'
    heading.fontSize = 15

    background = Image("back.png", width=PAGE_WIDTH, height=PAGE_HEIGHT + 100, hAlign='center')
    desc1 = Paragraph(text=f"Section 2 ", style=heading)
    desc = Paragraph(
        text=f"This section describes {info['Full Name'].item()}'s performance v/s the Rest of the World in Grade 3.")
    report_table_data = [
        ("Question\nNo.", 'Attempt Status', f"{info['First Name'].item()}'s\nChoice", 'Correct \nAnswer',
         'Outcome', f"{info['First Name'].item()}'s\nScore", "% of students\nacross the world\nwho attempted\nthis question",
         "% of students (from\nthose who attempted\nthis ) who got it\ncorrect",
         "% of students\n(from those who\nattempted this)\n"
         "who got it\nincorrect",
         f"World Average\nin this question",),
    ]
    for _, row in dataset_marksheet_section2_student.iterrows():
        report_table_data.append(
            (row['Question No.'], row['Attempt Status'], row['What you marked'], row['Correct Answer'],
             row['Outcome (Correct/Incorrect/Not Attempted)'], row['Your score'],
             row['% of students\nacross the world\nwho attempted\nthis question'],
             row['% of students (from\nthose who attempted\nthis ) who got it\ncorrect'],
             row['% of students\n(from those who\nattempted this)\nwho got it\nincorrect'],
             row['World Average\nin this question'])
        )
    style = (
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
        ('BACKGROUND', (0, 0), (-1, 0), (0, 0, 0)),
        ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('TEXTCOLOR', (0, 0), (-1, 0), (1, 1, 1)),
        ('BOX', (0, 0), (-1, -1), 0.25, (0, 0, 0)),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, (0, 0, 0))
    )
    table = Table(report_table_data, style=style)

    percentile = info['World Percentile'].item()

    dets = Paragraph(
        text=f"<b>{info['First Name'].item()}</b>'s overall percentile in the world is <b>{percentile:.2f}%ile</b>. This indicates "
             f"that {info['First Name'].item()} has scored more than <b>{percentile:.2f}%</b> of students"
             f" in the World and lesser than <b>{100 - percentile:.2f}%</b> of students<br/>in the world.")

    mean = info['Average score of all students across the World'].item()
    attempt_per = info['First name\'s attempts (Attempts x 100 / Total Questions)'].item()
    accuracy_per = info['First name\'s Accuracy ( Corrects x 100 /Attempts )'].item()
    median = info['Median score of all students across the World'].item()
    avg_attempts = info['Average attempts of all students across the Worl'].item()
    avg_accuracy = info['Average accuracy of all students across the World'].item()
    mode = info['Mode score of all students across World'].item()

    tbl_data = (
        (Paragraph(text=f"<b>Overview</b>"),),
        ("Average score of  all\nstudents across the World", Paragraph(text=f"<b>{mean}</b>"), '',
         f"{info['First Name'].item()}'s attempts\n(Attempts x 100 / Total\nQuestions)", Paragraph(text=f"<b>{attempt_per}%</b>"), "",
         f"{info['First Name'].item()}'s Accuracy\n( Corrects x 100 /Attempts )", Paragraph(text=f"<b>{accuracy_per}%</b>")),

        (f"Median score of all \nstudents  across the World", Paragraph(text=f"<b>{median}</b>"), "",
         "Average attempts of all\nstudents across the World", Paragraph(text=f"<b>{avg_attempts}%</b>"), f"",
         "Average accuracy of  all\nstudents across the World", Paragraph(text=f"<b>{avg_accuracy}%</b>")),

        ("Mode score of all students across\nWorld", Paragraph(text=f"<b>{mode:.2f}</b>"), ''),
    )
    det_style = (
        ('ALIGN', (1, 2), (-1, 2), 'LEFT'),
        ('ALIGN', (1, 4), (-1, 4), 'LEFT'),
        ('ALIGN', (1, 6), (-1, 6), 'LEFT'),
        ('INNERGRID', (0, 1), (1, -1), 1, (0, 0, 0)),
        ('INNERGRID', (3, 1), (4, -2), 1, (0, 0, 0)),
        ('INNERGRID', (-2, 1), (-1, -2), 1, (0, 0, 0)),
        ('BOX', (0, 1), (1, -1), 1, (0, 0, 0)),
        ('BOX', (3, 1), (4, -2), 1, (0, 0, 0)),
        ('BOX', (-2, 1), (-1, -2), 1, (0, 0, 0)),
    )
    tbl = Table(tbl_data, colWidths=(170, 60, cm, 170, 60, cm, 170, 60), style=det_style)

    folder_name = 'Image/Bar1/'
    os.makedirs(folder_name, exist_ok=True)
    x = [info['First Name'].item(), 'Average', 'Median', 'Mode']
    y = [info['Total score'].item(), mean, median, mode]
    plt.bar(x, y)
    addlabels(x, y)
    plt.title("Comparison of Scores")
    plt.ylabel("Score")
    plt.savefig(folder_name+info['First Name'].item()+str(info['Registration Number'].item())+'.png')
    plt.close()
    bar1_img = folder_name + info['First Name'].item() + str(info['Registration Number'].item()) + '.png'
    bar1 = Image(bar1_img, width=8 * cm, height=8 * cm, hAlign='LEFT')

    folder_name = 'Image/Bar2/'
    os.makedirs(folder_name, exist_ok=True)
    x = [info['First Name'].item(), 'World']
    y = [attempt_per, avg_attempts]
    plt.bar(x, y)
    addlabels(x, y)
    plt.title("Comparison of Attempts(%)")
    plt.ylabel("Attempts(%)")
    plt.savefig(folder_name + info['First Name'].item() + str(info['Registration Number'].item()) + '.png')
    plt.close()
    bar2_img = folder_name + info['First Name'].item() + str(info['Registration Number'].item()) + '.png'
    bar2 = Image(bar2_img, width=8 * cm, height=8 * cm, hAlign='LEFT')

    folder_name = 'Image/Bar3/'
    os.makedirs(folder_name, exist_ok=True)
    x = [info['First Name'].item(), 'World']
    y = [accuracy_per, avg_accuracy]
    plt.bar(x, y)
    addlabels(x, y)
    plt.title("Comparison of Accuracy(%)")
    plt.ylabel("Accuracy(%)")
    plt.savefig(folder_name + info['First Name'].item() + str(info['Registration Number'].item()) + '.png')
    plt.close()
    bar3_img = folder_name + info['First Name'].item() + str(info['Registration Number'].item()) + '.png'

    bar3 = Image(bar3_img, width=8 * cm, height=8 * cm, hAlign='LEFT')

    data = (PDFItem(background, 0, 0),
            PDFItem(desc1, MARGIN_LEFT + 350, PAGE_TOP_2 - 100),
            PDFItem(desc, MARGIN_LEFT + 200, PAGE_TOP_2 - 130),
            PDFItem(table, MARGIN_LEFT, 250),
            PDFItem(dets, MARGIN_LEFT, 210),
            PDFItem(tbl, MARGIN_LEFT + 10, 80)
            )
    page_2a.add(data)

    page_2b = PDFPage()
    data2 = (PDFItem(background, 0, 0),
            PDFItem(bar1, MARGIN_LEFT + 30, 30),
            PDFItem(bar2, MARGIN_LEFT + PAGE_WIDTH // 3, 30),
            PDFItem(bar3, MARGIN_LEFT + (PAGE_WIDTH // 3) * 2, 30),
            )
    page_2b.add(data2)
    return page_2a, page_2b


def get_page_3(info, dataset, dataset_info, dataset_marksheet_section3_student):
    page_3a = PDFPage()

    styles1 = getSampleStyleSheet()

    heading = styles1["Normal"]
    heading.fontName = 'Times'
    heading.fontSize = 15

    background = Image("back.png", width=PAGE_WIDTH, height=PAGE_HEIGHT + 100, hAlign='center')
    desc1 = Paragraph(text=f"Section 3 ", style=heading)
    desc = Paragraph(
        text=f"This section describes {info['First Name'].item()}'s performance v/s the Rest of {info['Country of Residence'].item()} in Grade 3.")
    report_table_data = [
        ("Question\nNo.", "Attempt Status", f"{info['First Name'].item()}'s\nChoice", 'Correct \nAnswer',
         'Outcome', f"{info['First Name'].item()}'s\nScore", f"% of students\nacross the {info['Country of Residence'].item()}\nwho attempted\nthis question",
         "% of students (from\nthose who attempted\nthis ) who got it\ncorrect",
         "% of students\n(from those who\nattempted this)\n"
         "who got it\nincorrect",
         f"Average of {info['Country of Residence'].item()}\nin this question",),
    ]
    for _, row in dataset_marksheet_section3_student.iterrows():
        report_table_data.append(
            (row['Question No.'], row['Attempt Status'], row['What you marked'], row['Correct Answer'],
             row['Outcome (Correct/Incorrect/Not Attempted)'], row['Your score'],
             f"{row['Per_attempts']}%",
             f"{row['Per_correct']}%",
             f"{row['Per_incorrect']}%",
             f"{row['Average']}")
        )
    style = (
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
        ('BACKGROUND', (0, 0), (-1, 0), (0, 0, 0)),
        ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('TEXTCOLOR', (0, 0), (-1, 0), (1, 1, 1)),
        ('BOX', (0, 0), (-1, -1), 0.25, (0, 0, 0)),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, (0, 0, 0))
    )
    table = Table(report_table_data, style=style)

    percentile = info['Country Percentile'].item()

    dets = Paragraph(
        text=f"<b>{info['First Name'].item()}</b>'s overall percentile in {info['Country of Residence'].item()} is <b>{percentile:.2f}%ile</b>. This indicates "
             f"that {info['First Name'].item()} has scored more than <b>{percentile:.2f}%</b> of students"
             f" in {info['Country of Residence'].item()} and lesser than <b>{100 - percentile:.2f}%</b> of students in {info['Country of Residence'].item()}.")

    overview = get_overview_section3(info, dataset, dataset_info)

    mean = round(overview[0])
    attempt_per = round(overview[4])
    accuracy_per = round(overview[6])
    median = round(overview[1])
    avg_attempts = round(overview[3])
    avg_accuracy = round(overview[5])
    mode = round(overview[2])

    tbl_data = (
        (Paragraph(text=f"<b>Overview</b>"),),
        (f"Average score of  all\nstudents in {info['Country of Residence'].item()}", Paragraph(text=f"<b>{mean}</b>"), '',
         f"{info['First Name'].item()}'s attempts\n(Attempts x 100 / Total\nQuestions)", Paragraph(text=f"<b>{attempt_per}%</b>"), "",
         f"{info['First Name'].item()}'s Accuracy\n( Corrects x 100 /Attempts )", Paragraph(text=f"<b>{accuracy_per}%</b>")),

        (f"Median score of all\nstudents  in {info['Country of Residence'].item()}", Paragraph(text=f"<b>{median}</b>"), "",
         f"Average attempts of all\nstudents in {info['Country of Residence'].item()}", Paragraph(text=f"<b>{avg_attempts}%</b>"),
         '',
         f"Average accuracy of all\nstudents in {info['Country of Residence'].item()}", Paragraph(text=f"<b>{avg_accuracy}%</b>")),

        (f"Mode score of all students in\n{info['Country of Residence'].item()}", Paragraph(text=f"<b>{mode}</b>"), ''),
    )
    det_style = (
        ('ALIGN', (1, 2), (-1, 2), 'LEFT'),
        ('ALIGN', (1, 4), (-1, 4), 'LEFT'),
        ('ALIGN', (1, 6), (-1, 6), 'LEFT'),
        ('INNERGRID', (0, 1), (1, -1), 1, (0, 0, 0)),
        ('INNERGRID', (3, 1), (4, -2), 1, (0, 0, 0)),
        ('INNERGRID', (-2, 1), (-1, -2), 1, (0, 0, 0)),
        ('BOX', (0, 1), (1, -1), 1, (0, 0, 0)),
        ('BOX', (3, 1), (4, -2), 1, (0, 0, 0)),
        ('BOX', (-2, 1), (-1, -2), 1, (0, 0, 0)),
    )

    tbl = Table(tbl_data, colWidths=(170, 60, cm, 170, 60, cm, 170, 60), style=det_style)

    folder_name = 'Image2/Bar1/'
    os.makedirs(folder_name, exist_ok=True)
    x = [info['First Name'].item(), 'Average', 'Median', 'Mode']
    y = [info['Total score'].item(), mean, median, mode]
    plt.bar(x, y)
    addlabels(x, y)
    plt.title("Comparison of Scores")
    plt.ylabel("Score")
    plt.savefig(folder_name+info['First Name'].item()+str(info['Registration Number'].item())+'.png')
    plt.close()
    bar1_img = folder_name + info['First Name'].item() + str(info['Registration Number'].item()) + '.png'
    bar1 = Image(bar1_img, width=8 * cm, height=8 * cm, hAlign='LEFT')

    folder_name = 'Image2/Bar2/'
    os.makedirs(folder_name, exist_ok=True)
    x = [info['First Name'].item(), info['Country of Residence'].item()]
    y = [attempt_per, avg_attempts]
    plt.bar(x, y)
    addlabels(x, y)
    plt.title("Comparison of Attempts(%)")
    plt.ylabel("Attempts(%)")
    plt.savefig(folder_name + info['First Name'].item() + str(info['Registration Number'].item()) + '.png')
    plt.close()
    bar2_img = folder_name + info['First Name'].item() + str(info['Registration Number'].item()) + '.png'
    bar2 = Image(bar2_img, width=8 * cm, height=8 * cm, hAlign='LEFT')

    folder_name = 'Image2/Bar3/'
    os.makedirs(folder_name, exist_ok=True)
    x = [info['First Name'].item(), info['Country of Residence'].item()]
    y = [accuracy_per, avg_accuracy]
    plt.bar(x, y)
    addlabels(x, y)
    plt.title("Comparison of Accuracy(%)")
    plt.ylabel("Accuracy(%)")
    plt.savefig(folder_name + info['First Name'].item() + str(info['Registration Number'].item()) + '.png')
    plt.close()
    bar3_img = folder_name + info['First Name'].item() + str(info['Registration Number'].item()) + '.png'

    bar3 = Image(bar3_img, width=8 * cm, height=8 * cm, hAlign='LEFT')

    data = (PDFItem(background, 0, 0),
            PDFItem(desc1, MARGIN_LEFT + 350, PAGE_TOP_2 - 100),
            PDFItem(desc, MARGIN_LEFT + 200, PAGE_TOP_2 - 130),
            PDFItem(table, MARGIN_LEFT, 250),
            PDFItem(dets, MARGIN_LEFT, 210),
            PDFItem(tbl, MARGIN_LEFT + 10, 80)
            )
    page_3a.add(data)

    page_3b = PDFPage()
    data2 = (PDFItem(background, 0, 0),
            PDFItem(bar1, MARGIN_LEFT + 30, 30),
            PDFItem(bar2, MARGIN_LEFT + PAGE_WIDTH // 3, 30),
            PDFItem(bar3, MARGIN_LEFT + (PAGE_WIDTH // 3) * 2, 30),
            )
    page_3b.add(data2)
    return page_3a, page_3b


def get_overview_section3(info, dataset, dataset_info):
    overview = []

    basic_overview = dataset_info[dataset_info['Country of Residence'] == info['Country of Residence'].item()]
    lst_of_score = basic_overview['Total score'].to_list()
    mean = statistics.mean(lst_of_score)
    median = statistics.median(lst_of_score)
    try:
        mode = statistics.mode(lst_of_score)
    except:
        mode = info['Total score']

    country_dataset = dataset[dataset['Country of Residence'] == info['Country of Residence'].item()]
    per_country_dataset = country_dataset[country_dataset['Registration Number'] == info['Registration Number'].item()]
    avg_attempts_country = (country_dataset[country_dataset['Attempt Status'] == 'Attempted']['Attempt Status'].count() / country_dataset['Attempt Status'].count()) * 100
    avg_attempt_per = (per_country_dataset[per_country_dataset['Attempt Status'] == 'Attempted']['Attempt Status'].count() / per_country_dataset['Attempt Status'].count()) * 100
    avg_accuracy_country = (country_dataset[country_dataset['Outcome (Correct/Incorrect/Not Attempted)'] == 'Correct']['Outcome (Correct/Incorrect/Not Attempted)'].count() / country_dataset[
        'Outcome (Correct/Incorrect/Not Attempted)'].count()) * 100
    avg_accuracy_per = (per_country_dataset[per_country_dataset['Outcome (Correct/Incorrect/Not Attempted)'] == 'Correct']['Outcome (Correct/Incorrect/Not Attempted)'].count() / per_country_dataset[
        'Outcome (Correct/Incorrect/Not Attempted)'].count()) * 100

    overview.append(mean)
    overview.append(median)
    overview.append(mode)
    overview.append(avg_attempts_country)
    overview.append(avg_attempt_per)
    overview.append(avg_accuracy_country)
    overview.append(avg_accuracy_per)

    return overview


def addlabels(x, y):
    for i in range(len(x)):
        plt.text(i, y[i], y[i])


def create_pdf(img_loc, info, dataset, dataset_info, dataset_marksheet_section1_student, dataset_marksheet_section2_student, dataset_marksheet_section3_student):

    page_1 = get_page_1(img_loc, info, dataset_marksheet_section1_student)
    page_2a, page_2b = get_page_2(info, dataset_marksheet_section2_student)
    # page_3a = get_page_3(info, dataset, dataset_info, dataset_marksheet_section3_student)
    page_3a, page_3b = get_page_3(info, dataset, dataset_info, dataset_marksheet_section3_student)

    folder_name = 'Report/'
    os.makedirs(folder_name, exist_ok=True)

    pdf_name = '{0}{1} ({2}).pdf'.format(folder_name, info['Full Name'].item(), info['Registration Number'].item())

    pdf = PDF(dest=pdf_name)
    pdf.add_page(page_1)
    pdf.add_page(page_2a)
    pdf.add_page(page_2b)
    pdf.add_page(page_3a)
    pdf.add_page(page_3b)

    # pdf.prepare()
    pdf.prepare(size=(800, 850))
