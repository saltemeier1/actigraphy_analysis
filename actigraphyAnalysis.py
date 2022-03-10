import os
import re
import xlsxwriter
import shutil
from datetime import datetime
import csv
from statistics import mean
import statistics

# script starts at the bottom, scroll down to @main to understand functionality.

# FOLLOW THESE STEPS IF YOU WANT TO ADD AN ADDITIONAL VARIABLE:
    # Right now, this script grabs the entire row of data for the interval types SLEEP and ACTIVE. (occurs in @readFile)
    # This means that if you want to grab something additional from one of these rows, it should not
    # be too difficult. If you want to add a variable from the SLEEP interval type, you will need to edit @parseSleep
    # If you want to add a variable from the ACTIVE interval type, you will need to edit @parseActive.
    # See above @parseSleep to see how you would need to edit that function to add a variable.


# grabbing the time of day for naming the created csv file
now = datetime.now()
dayOfMonth = str(now.month) + "_" + str(now.day)
time = str(now.hour) + "_" + str(now.minute)
# all_participants will be a list of lists, holding each individual participant's data in its own list
# i.e. [["participant 1 ID","participant 1 days recorded",..], ["participant 2 ID", "part. 2 days recorded",..],..]
all_participants = []

# writeCSV creates the output excel file. It adds all the necessary headers,
# then loops through the all_participants list and adds each participant data
# to its own row.

# IF YOU WANT TO ADD AN ADDITIONAL VARIABLE:
    # you will need to add an additional header if you want your new
    # variable data to be included in the output file. Add a header by adding a line
    # that looks something like, "summarySheet.write("CellValue", "NewSleepVariable")"
def writeCSV(data, source, dest):
    fileName = dayOfMonth +"_Actigraphy_Summary_" + time +".xlsx"
    summaryStats = xlsxwriter.Workbook(fileName)
    summarySheet = summaryStats.add_worksheet()
    summarySheet.write("A1", "Participant ID")
    summarySheet.write("B1", "days_recorded")
    summarySheet.write("C1", "wknd_days_recorded")
    summarySheet.write("D1", "weekdays_recorded")
    summarySheet.write("E1", "mean_bedtime")
    summarySheet.write("F1", "std_bedtime")
    summarySheet.write("G1", "mean_waketime")
    summarySheet.write("H1", "std_waketime")
    summarySheet.write("I1", "mean_bedtime_wknd")
    summarySheet.write("J1", "mean_waketime_wknd")
    summarySheet.write("K1", "mean_bedtime_week")
    summarySheet.write("L1", "mean_waketime_week")
    summarySheet.write("M1", "mean_Avg_AC_min")
    summarySheet.write("N1", "std_Avg_AC_min")
    summarySheet.write("O1", "mean_sleep_latency")
    summarySheet.write("P1", "std_sleep_latency")
    summarySheet.write("Q1", "mean_efficiency")
    summarySheet.write("R1", "std_efficiency")
    summarySheet.write("S1", "mean_WASO")
    summarySheet.write("T1", "std_WASO")
    summarySheet.write("U1", "mean_TST")
    summarySheet.write("V1", "std_TST")
    summarySheet.write("W1", "mean_white")
    summarySheet.write("X1", "std_white")
    summarySheet.write("Y1", "mean_red")
    summarySheet.write("Z1", "std_red")
    summarySheet.write("AA1", "mean_green")
    summarySheet.write("AB1", "std_green")
    summarySheet.write("AC1", "mean_blue")
    summarySheet.write("AD1", "std_blue")
    row = 1
    column = 0
    for patient in data:
        for stat in patient:
            summarySheet.write(row, column, stat)
            column += 1
        row += 1
        column = 0
    summaryStats.close()
    # finally, the file is moved to the Summary folder
    shutil.move(source + "\\" + fileName, dest + "\\" + fileName)

# use a normal 24 time clock, where Midnight = 24:00/0:00, noon = 12:00 to calculate the average wake time
def parseWakeTime(time_list):
    totalTime = 0
    timeInSeconds = []
    for time in time_list:
        hour, min, end = time.split(':')
        sec, timeStamp = end.split()
        hour = int(hour)
        min = int(min)
        sec = int(sec)
        if (timeStamp == "PM") & (hour < 12):
            hour = hour + 12
        elif (timeStamp == "AM") & (hour == 12) & (min == 00) & (sec == 00):
            hour = hour+12
        elif (timeStamp == "AM") & (hour == 12):
            hour = hour - 12
        timeInSeconds.append(hour*3600 + min*60 + sec)
        totalTime = totalTime + hour*3600 + min*60 + sec
    totalTime_avg = totalTime / len(time_list)
    stdWakeTime = (statistics.stdev(timeInSeconds, xbar = totalTime_avg))/60
    avg_hour = int(totalTime_avg // 3600)
    totalTime_avg = totalTime_avg%3600
    avg_min = int(totalTime_avg // 60)
    avg_sec = int(round(totalTime_avg%60, 0))
    if (avg_hour < 12):
        timeStamp = "AM"
    else:
        timeStamp = "PM"
    if (avg_hour > 12):
        avg_hour = avg_hour - 12
    if (len(str(avg_min)) == 1):
        avg_min = "0" + str(avg_min)
    if (len(str(avg_sec)) == 1):
        avg_sec = "0" + str(avg_sec)
    return str(avg_hour) + ":" + str(avg_min) + ":" + str(avg_sec) + " " + timeStamp, stdWakeTime

# use an adjusted 24 time clock, where Midnight = 12:00, noon = 24:00/0:00,
# to calculate the average bed time
def parseBedTime(time_list):
    totalTime = 0
    timeInSeconds = []
    for time in time_list:
        hour, min, end = time.split(':')
        sec, timeStamp = end.split()
        hour = int(hour)
        min = int(min)
        sec = int(sec)
        if (timeStamp == "AM") & (hour < 12):
            hour = hour + 12
        elif (timeStamp == "PM") & (hour == 12) & (min == 00) & (sec == 00):
            hour = hour+12
        elif (timeStamp == "PM") & (hour == 12):
            hour = hour - 12
        timeInSeconds.append(hour*3600 + min*60 + sec)
        totalTime = totalTime + hour*3600 + min*60 + sec
    totalTime_avg = totalTime / len(time_list)
    stdSleepTime = (statistics.stdev(timeInSeconds, xbar =totalTime_avg ))/60
    avg_hour = int(totalTime_avg // 3600)
    totalTime_avg = totalTime_avg%3600
    avg_min = int(totalTime_avg // 60)
    totalTime_avg = int(totalTime_avg%60)
    avg_sec = totalTime_avg
    if (avg_hour >=12):
        timeStamp = "AM"
    else:
        timeStamp = "PM"
    if (avg_hour > 12):
        avg_hour = avg_hour - 12
    if (len(str(avg_min)) == 1):
        avg_min = "0" + str(avg_min)
    if (len(str(avg_sec)) == 1):
        avg_sec = "0" + str(avg_sec)
    return str(avg_hour) + ":" + str(avg_min) + ":" + str(avg_sec) + " ", stdSleepTime

# parseActive collects data needed on Active
def parseActive(active, weekends, weekdays):
    avgAcPerMin = []
    white_active = []
    red_active = []
    green_active = []
    blue_active = []
    NumDaysOfAct = len(active)
    numOfWeekdays = len(weekdays)
    numOfWeekends = len(weekends)
    for row in active:
        avgAcPerMin.append(float(row[12]))
        white_active.append(float(row[20]))
        red_active.append(float(row[24]))
        green_active.append(float(row[28]))
        blue_active.append(float(row[32]))
    avgAcPerMin_avg = mean(avgAcPerMin)
    stdAcPerMin = statistics.stdev(avgAcPerMin, xbar = avgAcPerMin_avg)
    avg_white_active = mean(white_active)
    avg_red_active = mean(red_active)
    avg_green_active = mean(green_active)
    avg_blue_active = mean(blue_active)
    std_white_active = statistics.stdev(white_active, xbar = avg_white_active)
    std_red_active = statistics.stdev(red_active, xbar = avg_red_active )
    std_green_active = statistics.stdev(green_active, xbar = avg_green_active )
    std_blue_active = statistics.stdev(blue_active, xbar =avg_blue_active )
    return [NumDaysOfAct, numOfWeekends, numOfWeekdays], [avgAcPerMin_avg, stdAcPerMin], [avg_white_active, std_white_active, avg_red_active, std_red_active, avg_green_active, std_green_active, avg_blue_active, std_blue_active]

# parseSleep collects data needed on Sleep, utilizes both @parseBedTime and @parseWakeTime
# IF YOU WANT TO ADD AN ADDITIONAL VARIABLE:
    # there is a list created for each variable that we keep track of. So first, you would add
    # a blank new list. (e.g. newSleepVariable = []) Then, you need to find, in the raw data csv,
    # what column that data is held in. Keep in mind, the first column is column 0 (header A in excel),
    # the next would be 1 (header b in excel), the next would be 2 (header c in excel)..... etc.
    # Then, you can add this data into your new list by adding a line under the 'for column in sleep' loop.
    # this would look something like: newSleepVariable.append(column[*column # you just determined]).
    # after this is done, the script should be holding your wanted data in the your newly created list,
    # and you can analyze this list however you see fit (e.g. get the average like many of the other variables).
    # After you analyzed it, make sure to add your calculated value to the returned list. (e.g. newSleepVariable_avg)
    # You will need to edit @writeCSV to add a column header for your variable.
    # Additionally, this same process could be carried out in @parseActive.
def parseSleep(sleep, weekends, weekdays):
    sleepStart = []
    sleepFinish = []
    sleepWeekendStart = []
    sleepWeekendFinish = []
    sleepWeekdayStart = []
    sleepWeekdayFinish = []
    onsetLat = []
    efficiency = []
    waso = []
    sleepTime = []
    for column in sleep:
        sleepStart.append(column[4])
        sleepFinish.append(column[7])
        onsetLat.append(float(column[14]))
        efficiency.append(float(column[15]))
        waso.append(float(column[16]))
        sleepTime.append(float(column[17]))
    startSleepTime_avg, sleepTime_std = parseBedTime(sleepStart)
    finishSleepTime_avg, wakeTime_std = parseWakeTime(sleepFinish)
    onsetLat_avg = mean(onsetLat)
    stdOnsetLat = statistics.stdev(onsetLat, xbar = onsetLat_avg)
    efficiency_avg = mean(efficiency)
    stdEfficiency = statistics.stdev(efficiency, xbar= efficiency_avg)
    waso_avg = mean(waso)
    stdWaso = statistics.stdev(waso, xbar =waso_avg )
    sleepTime_avg = mean(sleepTime)
    stdSleepTime = statistics.stdev(sleepTime, xbar = sleepTime_avg)
    for column in weekends:
        sleepWeekendStart.append(column[4])
        sleepWeekendFinish.append(column[7])
    for column in weekdays:
        sleepWeekdayStart.append(column[4])
        sleepWeekdayFinish.append(column[7])
    startSleepWeekend_avg, sleepWeekendTime_std = parseBedTime(sleepWeekendStart)
    finishSleepWeekend_avg, wakeWeekendTime_std = parseWakeTime(sleepWeekendFinish)
    startSleepWeekday_avg, sleepWeekdayTime_std = parseBedTime(sleepWeekdayStart)
    finishSleepWeekday_avg, wakeWeekdayTime_std = parseWakeTime(sleepWeekdayFinish)
    return [startSleepTime_avg, sleepTime_std,finishSleepTime_avg, wakeTime_std, startSleepWeekend_avg, finishSleepWeekend_avg, startSleepWeekday_avg, finishSleepWeekday_avg], [onsetLat_avg, stdOnsetLat, efficiency_avg, stdEfficiency, waso_avg, stdWaso, sleepTime_avg, stdSleepTime]



# readFile looks at each row in a file and organizes the needed data into lists
# Active - holds the Active data
# Sleep - holds the Sleep data
# there are also lists that hold the above categories divided into weekends and weekdays
# after these lists are created, they are parsed @parseSleep, @parseActive
def readFile(fileToRead, participantData):
    fileAsList = []
    Identity = []
    Active = []
    Sleep = []
    active_weekdays = []
    sleep_weekdays = []
    active_weekends = []
    sleep_weekends = []
    days_rec, ACdata, RGBdata, times, otherSleepInfo = [], [], [], [], []
    for row in fileToRead:
        fileAsList.append(row)
    index = 0
    while index < len(fileAsList):
        if (fileAsList[index] != []):
            if (fileAsList[index][0] == 'Identity:'):
                Identity.append(fileAsList[index][1])
            elif (fileAsList[index][0] == 'ACTIVE'):
                Active.append(fileAsList[index])
                if (fileAsList[index][3] == 'Fri' or fileAsList[index][3] == 'Sat'):
                    active_weekends.append(fileAsList[index])
                else:
                    active_weekdays.append(fileAsList[index])
            elif (fileAsList[index][0] == 'SLEEP'):
                Sleep.append(fileAsList[index])
                frontHalf, timeStamp = fileAsList[index][4].split()
                if (((fileAsList[index][3] == 'Fri') & (timeStamp == 'PM')) or ((fileAsList[index][3] == 'Sat') & (timeStamp == "AM")) or ((fileAsList[index][3] == 'Sat') & (timeStamp == "PM")) or (((fileAsList[index][3]) == 'Sun') & (timeStamp == "AM"))):
                    sleep_weekends.append(fileAsList[index])
                else:
                    sleep_weekdays.append(fileAsList[index])
        index += 1
    days_rec, ACdata, RGBdata = parseActive(Active, active_weekends, active_weekdays)
    times, otherSleepInfo = parseSleep(Sleep, sleep_weekends, sleep_weekdays)
    return Identity + days_rec + times + ACdata + otherSleepInfo + RGBdata

# main loops through the reports folder and reads each file (@readFile),
# adding gathered data to all_participants list.
# Once all report files have been parsed, create output csv file.
def main():
    rootDir = os.getcwd()
    os.chdir('Reports')
    reports = os.listdir(os.getcwd())
    for file in reports:
        participantData = []
        with open(file, 'r') as csv_file:
            reader = csv.reader(csv_file)
            temp_data = readFile(reader, participantData)
            participantData = participantData + temp_data
            all_participants.append(participantData)
    writeCSV(all_participants, rootDir+"\Reports", rootDir + "\Summary")

main()