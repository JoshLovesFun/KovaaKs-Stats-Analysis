from os import listdir
from xlwt import Workbook, easyxf

BOLD_FONT_FORMAT = easyxf('font: bold 1')
VALORANT_SENSITIVITY_CONVERSION = 16.33

# Example path to files associated with Steam account
# path = (
#     "C:\\Program Files (x86)\\Steam\\steamapps\\common\\"
#     "FPSAimTrainer\\FPSAimTrainer\\stats"
# )

# Enter KovaaK's stats path
path = "C:\\Code\\Python\\GitHub\\KovaaKs-Stats-Analysis\\ExampleStats"

# This function creates an array of all the file names in the path
files = listdir(path)
files.sort()

# Create Excel workbook and worksheets
wb = Workbook()
sheet1 = wb.add_sheet('All Stats')
sheet2 = wb.add_sheet('Daily Stats')
sheet3 = wb.add_sheet('Monthly Stats')

# Create Excel table titles
sheet1.write(0, 0, 'Name', BOLD_FONT_FORMAT)
sheet1.write(0, 1, 'Date', BOLD_FONT_FORMAT)
sheet1.write(0, 2, 'Score', BOLD_FONT_FORMAT)
sheet1.write(0, 3, 'Sens', BOLD_FONT_FORMAT)
sheet2.write(0, 0, 'Name', BOLD_FONT_FORMAT)
sheet2.write(0, 1, 'Date', BOLD_FONT_FORMAT)
sheet2.write(0, 2, 'Daily Plays', BOLD_FONT_FORMAT)
sheet2.write(0, 3, 'Ave Score', BOLD_FONT_FORMAT)
sheet2.write(0, 4, 'Max Score', BOLD_FONT_FORMAT)
sheet2.write(0, 5, 'Ave Sens', BOLD_FONT_FORMAT)
sheet3.write(0, 0, 'Name', BOLD_FONT_FORMAT)
sheet3.write(0, 1, 'Date', BOLD_FONT_FORMAT)
sheet3.write(0, 2, 'Monthly Plays', BOLD_FONT_FORMAT)
sheet3.write(0, 3, 'Ave Score', BOLD_FONT_FORMAT)
sheet3.write(0, 4, 'Max Score', BOLD_FONT_FORMAT)
sheet3.write(0, 5, 'Ave Sens', BOLD_FONT_FORMAT)

# Iterate through all KovaaK's stats files ####################################
for i in range(0, len(files)):
    # Get task name from file name
    File_Name = files[i]
    Task_Name = File_Name[0:File_Name.find(" - Challenge - ")]

    # Get task date from file name
    Date = File_Name[
        File_Name.find(" - Challenge - ") + 15:
        File_Name.find(" Stats") - 9
    ]
    Date = Date[5:7] + "/" + Date[8:10] + "/" + Date[0:4]

    # Open stats files
    with open(f"{path}/{files[i]}", newline='\n') as csvfile:
        # Iterate through every line of each KovaaK's stats file
        for ii in csvfile:
            # If line has score in it
            if "Score" in ii:
                # Get score from stats file
                Score = ii[7:].strip()
            if "Horiz Sens" in ii:
                # Get sens from stats file
                Sens = ii[12:].strip()
        # Convert Valorant to cm/360
        if int(round(float(Sens), 2)) < 1:
            Sens = round(
                VALORANT_SENSITIVITY_CONVERSION / round(float(Sens), 2), 2
            )
            Sens = str(Sens)

        # Write results to text file
        sheet1.write(i+1, 0, Task_Name)
        sheet1.write(i+1, 1, Date)
        sheet1.write(i+1, 2, round(float(Score), 2))
        sheet1.write(i+1, 3, round(float(Sens), 2))

# Iterate through all KovaaK's stats files daily ##############################
Count = 1
Score_Sum = 0
Sens_Sum = 0
Max_Score = 0
row_index = 0
for i in range(0, len(files)):
    # Get task name from file name
    File_Name = files[i]
    Task_Name = File_Name[0:File_Name.find(" - Challenge - ")]

    # Get task date from file name
    Date = File_Name[
        File_Name.find(" - Challenge - ") + 15:
        File_Name.find(" Stats") - 9
    ]

    # Get day
    Day = Date[8:]

    Future_Task_Name = None
    Future_Day = None

    # Next stats
    if i < len(files) - 1:
        Future_File_Name = files[i + 1]
        challenge_start = Future_File_Name.find(" - Challenge - ")
        stats_start = Future_File_Name.find(" Stats")
        Future_Task_Name = Future_File_Name[:challenge_start]
        Future_Date = Future_File_Name[challenge_start + 15:stats_start - 9]
        Future_Day = Future_Date[8:]

    # Open stats files
    with open(f"{path}/{files[i]}", newline='\n') as csvfile:
        # Iterate through every line of each KovaaK's stats file
        for ii in csvfile:
            # If line has score in it
            if "Score" in ii:
                # Get score from stats file
                Score = ii[7:].strip()
            if "Horiz Sens" in ii:
                # Get sens from stats file
                Sens = ii[12:].strip()
        # Convert Valorant to cm/360
        if int(round(float(Sens), 2)) < 1:
            Sens = round(
                VALORANT_SENSITIVITY_CONVERSION / round(float(Sens), 2), 2
            )
        # Pull max value from range
        if int(round(float(Score), 2)) > Max_Score:
            Max_Score = int(round(float(Score), 2))
        # If future score is the same day
        if (
            Day == Future_Day
            and Task_Name == Future_Task_Name
            and i != len(files) - 1
        ):
            Count += 1
            Score_Sum = round(Score_Sum + int(round(float(Score), 2)), 2)
            Sens_Sum = round(Sens_Sum + int(round(float(Sens), 2)), 2)
        # If future score is not the same day
        else:
            # Be sure count is greater than 1
            if Count > 1:
                Score = round(
                    (Score_Sum + int(round(float(Score), 2))) / Count, 2
                )
                Sens = round(
                    (Sens_Sum + int(round(float(Sens), 2))) / Count, 2
                )
            # Write results to Excel file
            sheet2.write(row_index + 1, 0, Task_Name)
            sheet2.write(row_index + 1, 1, Date)
            sheet2.write(row_index + 1, 2, Count)
            sheet2.write(row_index + 1, 3, round(float(Score), 2))
            sheet2.write(row_index + 1, 4, round(float(Max_Score), 2))
            sheet2.write(row_index + 1, 5, round(float(Sens), 2))
            row_index = row_index + 1
            # Reset values
            Count = 1
            Score_Sum = 0
            Sens_Sum = 0
            Max_Score = 0

# Iterate through all KovaaK's stats files monthly ############################
Count = 1
Score_Sum = 0
Sens_Sum = 0
Max_Score = 0
row_index = 0
for i in range(0, len(files)):
    # Get task name from file name
    File_Name = files[i]
    Task_Name = File_Name[0:File_Name.find(" - Challenge - ")]

    # Get task date from file name
    challenge_start = File_Name.find(" - Challenge - ") + 15
    stats_end = File_Name.find(" Stats") - 12
    Date = File_Name[challenge_start:stats_end]

    # Get month
    Month = Date[5:7]

    Future_Task_Name = None
    Future_Month = None

    # Next stats
    if i < len(files) - 1:
        Future_File_Name = files[i + 1]
        challenge_start = Future_File_Name.find(" - Challenge - ")
        stats_start = Future_File_Name.find(" Stats")

        Future_Task_Name = Future_File_Name[:challenge_start]
        Future_Date = Future_File_Name[challenge_start + 15:stats_start - 12]
        Future_Month = Future_Date[5:7]

    # Open stats files
    with open(f"{path}/{files[i]}", newline='\n') as csvfile:
        # Iterate through every line of each KovaaK's stats file
        for ii in csvfile:
            # If line has score in it
            if "Score" in ii:
                # Get score from stats file
                Score = ii[7:].strip()
            if "Horiz Sens" in ii:
                # Get sens from stats file
                Sens = ii[12:].strip()
        # Convert Valorant to cm/360
        if int(round(float(Sens), 2)) < 1:
            Sens = round(
                VALORANT_SENSITIVITY_CONVERSION / round(float(Sens), 2), 2
            )
        # Pull max value from range
        if int(round(float(Score), 2)) > Max_Score:
            Max_Score = int(round(float(Score), 2))
        # If future score is the same month
        if (
            Month == Future_Month
            and Task_Name == Future_Task_Name
            and i != len(files) - 1
        ):
            Count += 1
            Score_Sum = round(Score_Sum + int(round(float(Score), 2)), 2)
            Sens_Sum = round(Sens_Sum + int(round(float(Sens), 2)), 2)
        # If future score is not the same month
        else:
            # Be sure count is greater than 1
            if Count > 1:
                Score = round(
                    (Score_Sum + int(round(float(Score), 2))) / Count, 2
                )
                Sens = round(
                    (Sens_Sum + int(round(float(Sens), 2))) / Count, 2
                )
            # Write results to Excel file
            sheet3.write(row_index + 1, 0, Task_Name)
            sheet3.write(row_index + 1, 1, Date)
            sheet3.write(row_index + 1, 2, Count)
            sheet3.write(row_index + 1, 3, round(float(Score), 2))
            sheet3.write(row_index + 1, 4, round(float(Max_Score), 2))
            sheet3.write(row_index + 1, 5, round(float(Sens), 2))
            row_index += 1
            # Reset values
            Count = 1
            Score_Sum = 0
            Sens_Sum = 0
            Max_Score = 0

# Close Excel file
wb.save('KovaaKs_Stats_Analysis.xls')
