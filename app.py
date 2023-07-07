import xlsxwriter
from random import uniform
import datetime
import pandas as pd
from email.message import EmailMessage
import smtplib
from time import sleep

def define_date():
    date_now = {
        "year": "",
        "month": "",
        "friday":"",
        "monday":""
    }
    date_arr = str(datetime.date.today()).split("-")
    date_now["year"]=date_arr[0]
    date_now["month"]=date_arr[1]
    date_now["friday"]=date_arr[2]
    date_now["monday"]=str(int(date_arr[2])-4)
    return date_now

date_dict = define_date()

def create_mail(applicant_email, subject, to):
    files = ['Timesheet_Map.xlsx']
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = applicant_email
    msg['To'] = to
    msg.set_content('')
    for file in files:
        with open(file, 'rb') as file:
            file_data = file.read()
            file_name = file.name
        msg.add_attachment(
            file_data,
            maintype='application',
            subtype='octet-stream',
            filename=file_name
        )
    return msg

def send_mail(applicant_email, my_password, arrival_mail):
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(applicant_email, my_password)
        smtp.send_message(create_mail(
                applicant_email,
                "Timesheet map",
                arrival_mail
            ))



input_df = pd.read_csv("input.csv", delimiter=";")
arrival_email = input_df.loc[10, "Monday_min"]
pw = input_df.loc[11, "Monday_min"]
departure_email = input_df.loc[12, "Monday_min"]
name = str(input_df.iloc[8,1])
staff_number = str(input_df.iloc[9,1])
indexes = list(input_df.iloc[:,0])

timetable = input_df.drop(input_df.columns[[0]], axis=1)

timetable.index=indexes


def format_output(val):
    if int(val) <=0:
        return "-"
    return val

def project_hours(project):
    return (
    round(uniform(float(timetable.loc[project,"Monday_min"]),float(timetable.loc[project,"Monday_max"])),1),
    round(uniform(float(timetable.loc[project,"Tuesday_min"]),float(timetable.loc[project,"Tuesday_max"])),1),
    round(uniform(float(timetable.loc[project,"weddnesday_min"]),float(timetable.loc[project,"weddnesday_max"])),1),
    round(uniform(float(timetable.loc[project,"Thursday_min"]),float(timetable.loc[project,"Thursday_max"])),1),
    round(uniform(float(timetable.loc[project,"Friday_min"]),float(timetable.loc[project,"Friday_max"])),1),
    )


Broccoli_Staffing_hours = project_hours("Project CortexPOS")
TD_Global_hours = project_hours("Project Jobboard")
Dealerocassion_Basucco = project_hours("Project Waive")
Broccoli_Web_Page = project_hours("Broccoli Web Page")
Project_Waive = project_hours("Dealerocassion/Basucco")
Project_Jobboard = project_hours("TD Global")
Project_CortexPOS = project_hours("Broccoli Staffing")
mammography_hours = project_hours("Mamography")

def daily_sum(day_index):
    return sum((mammography_hours[day_index],Broccoli_Staffing_hours[day_index],
        TD_Global_hours[day_index],Dealerocassion_Basucco[day_index],
        Broccoli_Web_Page[day_index],Project_Waive[day_index],Project_Jobboard[day_index],
        Project_CortexPOS[day_index]
    ))



file = xlsxwriter.Workbook("Timesheet_Map.xlsx")
worksheet = file.add_worksheet("Blad1")

#styles
BOLD_AND_LARGE = file.add_format({
    "font_size": 22,
    "bold": True,
})
BOLD=file.add_format({
    "font_size":12,
    "bold":True,
    "bg_color": "#BFBFBF",
    "align": "center"
})
REGULAR_FORMAT = file.add_format({
    "font_size":12,
})
GRAY_BG = file.add_format({
    "font_size":12,
    "bg_color": "#BFBFBF"
})
WEEK_DAY = file.add_format({
    "font_size":12,
    "bold":True,
    "bg_color": "#BFBFBF",
    "align": "right"
})



worksheet.set_column(0,0,16)
worksheet.set_column(1,1,40)
worksheet.set_column(2,7,13)

worksheet.merge_range("A22:B22","Merged Range")
worksheet.merge_range("A30:B30","Merged Range")
worksheet.merge_range("A41:B41","Merged Range")

worksheet.write(0,0,"Weekly Time Registration Form Broccoli", BOLD_AND_LARGE)

worksheet.write(2,0,"Date",REGULAR_FORMAT)
worksheet.write(2,1,
    f'from {date_dict["year"]}/{date_dict["month"]}/{date_dict["monday"]} to {date_dict["year"]}/{date_dict["month"]}/{date_dict["friday"]}',
    REGULAR_FORMAT
)

worksheet.write(3,0,"Name",REGULAR_FORMAT)
worksheet.write(3,1,name,REGULAR_FORMAT)

worksheet.write(4,0,"Staff Number",REGULAR_FORMAT)
worksheet.write(4,1,staff_number,REGULAR_FORMAT)

worksheet.write(7,0,f"", BOLD)
worksheet.write(7,1,f"Time accountint", BOLD)
worksheet.write(7,2,f"", BOLD)

worksheet.write(8,1,f"Hours of absense",REGULAR_FORMAT)
worksheet.write(8,2,f"-",REGULAR_FORMAT)

worksheet.write(9,1,f"Indirect hours",REGULAR_FORMAT)
worksheet.write(9,2,f"-",REGULAR_FORMAT)

worksheet.write(10,1,"hours on projects",REGULAR_FORMAT)
worksheet.write(10,2,sum(mammography_hours),REGULAR_FORMAT)

worksheet.write(11,0,f"",GRAY_BG)
worksheet.write(11,1,f"total",GRAY_BG)
worksheet.write(11,2, 40,GRAY_BG)

worksheet.write(13,1,f"Weekly standard hours",REGULAR_FORMAT)
worksheet.write(13,2,40,REGULAR_FORMAT)

worksheet.write(15,0,f"",GRAY_BG)
worksheet.write(15,1,f"Difference",GRAY_BG)
worksheet.write(15,2,sum(mammography_hours)-40,GRAY_BG)

worksheet.write(21,0,f"",BOLD)
worksheet.write(21,0,f"Hours of absense",BOLD)
worksheet.write(21,2,f"Monday",WEEK_DAY)
worksheet.write(21,3,f"Tuesday",WEEK_DAY)
worksheet.write(21,4,f"Wednesday",WEEK_DAY)
worksheet.write(21,5,f"Thursday",WEEK_DAY)
worksheet.write(21,6,f"Friday",WEEK_DAY)
worksheet.write(21,7,f"Week-end",WEEK_DAY)

worksheet.write(22,1,f"Holidays",REGULAR_FORMAT)
worksheet.write(22,2,f"-",REGULAR_FORMAT)
worksheet.write(22,3,f"-",REGULAR_FORMAT)
worksheet.write(22,4,f"-",REGULAR_FORMAT)
worksheet.write(22,5,f"-",REGULAR_FORMAT)
worksheet.write(22,6,f"-",REGULAR_FORMAT)
worksheet.write(22,7,f"-",REGULAR_FORMAT)

worksheet.write(23,1,f"Leave",REGULAR_FORMAT)
worksheet.write(23,2,f"-",REGULAR_FORMAT)
worksheet.write(23,3,f"-",REGULAR_FORMAT)
worksheet.write(23,4,f"-",REGULAR_FORMAT)
worksheet.write(23,5,f"-",REGULAR_FORMAT)
worksheet.write(23,6,f"-",REGULAR_FORMAT)
worksheet.write(23,7,f"-",REGULAR_FORMAT)

worksheet.write(24,1,f"Illnes",REGULAR_FORMAT)
worksheet.write(24,2,f"-",REGULAR_FORMAT)
worksheet.write(24,3,f"-",REGULAR_FORMAT)
worksheet.write(24,4,f"-",REGULAR_FORMAT)
worksheet.write(24,5,f"-",REGULAR_FORMAT)
worksheet.write(24,6,f"-",REGULAR_FORMAT)
worksheet.write(24,7,f"-",REGULAR_FORMAT)

worksheet.write(25,1,f"Doctors visit",REGULAR_FORMAT)
worksheet.write(25,2,f"-",REGULAR_FORMAT)
worksheet.write(25,3,f"-",REGULAR_FORMAT)
worksheet.write(25,4,f"-",REGULAR_FORMAT)
worksheet.write(25,5,f"-",REGULAR_FORMAT)
worksheet.write(25,6,f"-",REGULAR_FORMAT)
worksheet.write(25,7,f"-",REGULAR_FORMAT)

worksheet.write(26,0,f"",GRAY_BG)
worksheet.write(26,1,f"total",GRAY_BG)
worksheet.write(26,2,f"-",GRAY_BG)
worksheet.write(26,3,f"-",GRAY_BG)
worksheet.write(26,4,f"-",GRAY_BG)
worksheet.write(26,5,f"-",GRAY_BG)
worksheet.write(26,6,f"-",GRAY_BG)
worksheet.write(26,7,f"-",GRAY_BG)

worksheet.write(29,0,f"Indirect Hours",BOLD)
worksheet.write(29,2,f"Monday",WEEK_DAY)
worksheet.write(29,3,f"Tuesday",WEEK_DAY)
worksheet.write(29,4,f"Wednesday",WEEK_DAY)
worksheet.write(29,5,f"Thursday",WEEK_DAY)
worksheet.write(29,6,f"Friday",WEEK_DAY)
worksheet.write(29,7,f"Week-end",WEEK_DAY)

worksheet.write(30,0,f"",REGULAR_FORMAT)
worksheet.write(30,1,f"Meeting",REGULAR_FORMAT)
worksheet.write(30,2,f"-",REGULAR_FORMAT)
worksheet.write(30,3,f"-",REGULAR_FORMAT)
worksheet.write(30,4,f"-",REGULAR_FORMAT)
worksheet.write(30,5,f"-",REGULAR_FORMAT)
worksheet.write(30,6,f"-",REGULAR_FORMAT)
worksheet.write(30,7,f"-",REGULAR_FORMAT)

worksheet.write(31,1,f"Education",REGULAR_FORMAT)
worksheet.write(31,2,f"-",REGULAR_FORMAT)
worksheet.write(31,3,f"-",REGULAR_FORMAT)
worksheet.write(31,4,f"-",REGULAR_FORMAT)
worksheet.write(31,5,f"-",REGULAR_FORMAT)
worksheet.write(31,6,f"-",REGULAR_FORMAT)
worksheet.write(31,7,f"-",REGULAR_FORMAT)

worksheet.write(32,1,f"Office taks as demandec by manager",REGULAR_FORMAT)
worksheet.write(32,2,f"-",REGULAR_FORMAT)
worksheet.write(32,3,f"-",REGULAR_FORMAT)
worksheet.write(32,4,f"-",REGULAR_FORMAT)
worksheet.write(32,5,f"-",REGULAR_FORMAT)
worksheet.write(32,6,f"-",REGULAR_FORMAT)
worksheet.write(32,7,f"-",REGULAR_FORMAT)

worksheet.write(34,1,f"Other",REGULAR_FORMAT)
worksheet.write(34,2,f"-",REGULAR_FORMAT)
worksheet.write(34,3,f"-",REGULAR_FORMAT)
worksheet.write(34,4,f"-",REGULAR_FORMAT)
worksheet.write(34,5,f"-",REGULAR_FORMAT)
worksheet.write(34,6,f"-",REGULAR_FORMAT)
worksheet.write(34,7,f"-",REGULAR_FORMAT)

worksheet.write(37,0,f"",GRAY_BG)
worksheet.write(37,1,f"total",GRAY_BG)
worksheet.write(37,2,f"-",GRAY_BG)
worksheet.write(37,3,f"-",GRAY_BG)
worksheet.write(37,4,f"-",GRAY_BG)
worksheet.write(37,5,f"-",GRAY_BG)
worksheet.write(37,6,f"-",GRAY_BG)
worksheet.write(37,7,f"-",GRAY_BG)

worksheet.write(40,0,f"Direct Hours",BOLD)
worksheet.write(40,2,f"Monday",WEEK_DAY)
worksheet.write(40,3,f"Tuesday",WEEK_DAY)
worksheet.write(40,4,f"Wednesday",WEEK_DAY)
worksheet.write(40,5,f"Thursday",WEEK_DAY)
worksheet.write(40,6,f"Friday",WEEK_DAY)
worksheet.write(40,7,f"Week-end",WEEK_DAY)

worksheet.write(41,1,f"Project CortexPOS",REGULAR_FORMAT)
worksheet.write(41,2,format_output(Project_CortexPOS[0]),REGULAR_FORMAT)
worksheet.write(41,3,format_output(Project_CortexPOS[1]),REGULAR_FORMAT)
worksheet.write(41,4,format_output(Project_CortexPOS[2]),REGULAR_FORMAT)
worksheet.write(41,5,format_output(Project_CortexPOS[3]),REGULAR_FORMAT)
worksheet.write(41,6,format_output(Project_CortexPOS[4]),REGULAR_FORMAT)
worksheet.write(41,7,"-",REGULAR_FORMAT)

worksheet.write(42,1,f"Project Jobboard",REGULAR_FORMAT)
worksheet.write(42,2,format_output(Project_Jobboard[0]),REGULAR_FORMAT)
worksheet.write(42,3,format_output(Project_Jobboard[1]),REGULAR_FORMAT)
worksheet.write(42,4,format_output(Project_Jobboard[2]),REGULAR_FORMAT)
worksheet.write(42,5,format_output(Project_Jobboard[3]),REGULAR_FORMAT)
worksheet.write(42,6,format_output(Project_Jobboard[4]),REGULAR_FORMAT)
worksheet.write(42,7,"-",REGULAR_FORMAT)

worksheet.write(43,1,f"Project Waive",REGULAR_FORMAT)
worksheet.write(43,2,format_output(Project_Waive[0]),REGULAR_FORMAT)
worksheet.write(43,3,format_output(Project_Waive[1]),REGULAR_FORMAT)
worksheet.write(43,4,format_output(Project_Waive[2]),REGULAR_FORMAT)
worksheet.write(43,5,format_output(Project_Waive[3]),REGULAR_FORMAT)
worksheet.write(43,6,format_output(Project_Waive[4]),REGULAR_FORMAT)
worksheet.write(43,7,"-",REGULAR_FORMAT)

worksheet.write(44,1,f"Broccoli Web Page",REGULAR_FORMAT)
worksheet.write(44,2,format_output(Broccoli_Web_Page[0]),REGULAR_FORMAT)
worksheet.write(44,3,format_output(Broccoli_Web_Page[1]),REGULAR_FORMAT)
worksheet.write(44,4,format_output(Broccoli_Web_Page[2]),REGULAR_FORMAT)
worksheet.write(44,5,format_output(Broccoli_Web_Page[3]),REGULAR_FORMAT)
worksheet.write(44,6,format_output(Broccoli_Web_Page[4]),REGULAR_FORMAT)
worksheet.write(44,7,"-",REGULAR_FORMAT)

worksheet.write(45,1,f"Dealerocassion/Basucco",REGULAR_FORMAT)
worksheet.write(45,2,format_output(Dealerocassion_Basucco[0]),REGULAR_FORMAT)
worksheet.write(45,3,format_output(Dealerocassion_Basucco[1]),REGULAR_FORMAT)
worksheet.write(45,4,format_output(Dealerocassion_Basucco[2]),REGULAR_FORMAT)
worksheet.write(45,5,format_output(Dealerocassion_Basucco[3]),REGULAR_FORMAT)
worksheet.write(45,6,format_output(Dealerocassion_Basucco[4]),REGULAR_FORMAT)
worksheet.write(45,7,"-",REGULAR_FORMAT)

worksheet.write(46,1,f"TD Global",REGULAR_FORMAT)
worksheet.write(46,2,format_output(TD_Global_hours[0]),REGULAR_FORMAT)
worksheet.write(46,3,format_output(TD_Global_hours[1]),REGULAR_FORMAT)
worksheet.write(46,4,format_output(TD_Global_hours[2]),REGULAR_FORMAT)
worksheet.write(46,5,format_output(TD_Global_hours[3]),REGULAR_FORMAT)
worksheet.write(46,6,format_output(TD_Global_hours[4]),REGULAR_FORMAT)
worksheet.write(46,7,"-",REGULAR_FORMAT)

worksheet.write(47,1,f"Broccoli Staffing",REGULAR_FORMAT)
worksheet.write(47,2,format_output(Broccoli_Staffing_hours[0]),REGULAR_FORMAT)
worksheet.write(47,3,format_output(Broccoli_Staffing_hours[1]),REGULAR_FORMAT)
worksheet.write(47,4,format_output(Broccoli_Staffing_hours[2]),REGULAR_FORMAT)
worksheet.write(47,5,format_output(Broccoli_Staffing_hours[3]),REGULAR_FORMAT)
worksheet.write(47,6,format_output(Broccoli_Staffing_hours[4]),REGULAR_FORMAT)
worksheet.write(47,7,"-",REGULAR_FORMAT)

worksheet.write(48,1,f"Mamography",REGULAR_FORMAT)
worksheet.write(48,2,format_output(mammography_hours[0]),REGULAR_FORMAT)
worksheet.write(48,3,format_output(mammography_hours[1]),REGULAR_FORMAT)
worksheet.write(48,4,format_output(mammography_hours[2]),REGULAR_FORMAT)
worksheet.write(48,5,format_output(mammography_hours[3]),REGULAR_FORMAT)
worksheet.write(48,6,format_output(mammography_hours[4]),REGULAR_FORMAT)
worksheet.write(48,7,"-",REGULAR_FORMAT)


worksheet.write(49,0,f"",GRAY_BG)
worksheet.write(49,1,f"total",GRAY_BG)
worksheet.write(49,2,format_output(daily_sum(0)),GRAY_BG)
worksheet.write(49,3,format_output(daily_sum(1)),GRAY_BG)
worksheet.write(49,4,format_output(daily_sum(2)),GRAY_BG)
worksheet.write(49,5,format_output(daily_sum(3)),GRAY_BG)
worksheet.write(49,6,format_output(daily_sum(4)),GRAY_BG)
worksheet.write(49,7,"-",GRAY_BG)


file.close()


sleep(10)
send_mail(arrival_email, pw, departure_email)