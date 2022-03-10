actigraphyAnalysis.py takes raw data from Philips Spectrum Pro watch and creates csv file summarizing certain statistics.

The output includes:Participant ID, days_recorded, wknd_days_recorded, weekdays_recorded, mean_bedtime, std_bedtime, mean_waketime, std_waketime, mean_bedtime_wknd,
mean_waketime_wknd, mean_bedtime_week, mean_waketime_week, mean_Avg_AC_min, std_Avg_AC_min, mean_sleep_latency, std_sleep_latency, mean_efficiency, std_efficiency,
mean_WASO, std_WASO, mean_TST, std_TST, mean_white, std_white, mean_red, std_red, mean_green, std_green, mean_blue, std_blue

RUNNING actigraphyAnalysis.py:::::::::::::::::::::::::::::::::::::::::

Place all csv actigraphy files that you want to be parsed into J:\Actigraphy Data\ActigraphyScript\Reports.
Then, you can open up the command prompt. Once opened, navigate into the folder that has this script, J:\Actigraphy Data\ActigraphyScript.
Once in this folder, you can type this into the command prompt window:

python actigraphyAnalysis.py

After this runs to completion, close the command window. You can now navigate to J:\Actigraphy Data\ActigraphyScript\Summary using File Explorer.
Here, you should find a new Excel file titled, "Day_Year_Actigraphy_summary_Hour_Minute" where Day,Year, Hour, and Minute correspond to when you ran this script.
This file holds all the participant data. 