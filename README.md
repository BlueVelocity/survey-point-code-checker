# WF Survey Point Code Checker

### Purpose

The purpose of this program is to allow the user to quickly check for point  codes
that are inconsistent with Walterfedy point code standards.

### Prerequisites

1. Ensure python3 and its dependencies are installed as per the installation
   guide in the support folder.
2. Process survey file: Remove any odd or blank points that are blank in the fifth
   column, remove station points, etc. (try sorting the fifth column using the 
   data tab in excel)
3. Copy the survey file into the folder with files 'point_checker.py' and
   'codes.xlsx'.
4. Rename the survey file to 'survey.xlsx'.

### Run the Program

1. Navigate to the folder in the file explorer containing 'point_checker.py', 
   'codes.xlsx', and 'survey.xlsx'.
2. Click the address bar at the top of the file explorer, type 'cmd', and press Enter.
3. In the command prompt, copy and paste the following command and press Enter:
	python3 point_checker.py
4. Check the results; a 'Review Sheet' should be added to the 'survey.xlsx'
   workbook with point numbers requiring the user's attention, including vlookup
   columns referencing the corresponding data for review.

### Support Contact Information

Joseph Monighan