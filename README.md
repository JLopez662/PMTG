# Project Management Excel Tools Generator (PMXTG) 

Generates automatically Gantt Charts, Project Schedule and RACI Table sheets in Excel, from a simple input, mapping out Milestones that includes Tasks with Hours across Work Weeks, ideal for Project Management.

- ![Generated Gantt Chart](./Gantt_Chart_Gif.gif)
- ![Generated Project Schedule](./Project_Schedule.png)
- ![Generated Raci Table](./RACI_Table.png)

## Purpose

The Project Management Excel Tools Generator (PMXTG) serves as an Automated Solution for generating detailed Gantt Charts, Project Schedules and RACI Tables sheets to Excel files, ideal for Project Managers to use during Project Proposals, and presentations to Stakeholders or Teams..

## Usage Instructions

1. **Prepare the Scripts:**  
   Ensure that `chronogram.py` and `chronogram.sh` filed are both located in the same directory.

2. **Set Script Permissions:**  
   Open a terminal and navigate to the directory containing the scripts. Give executable permissions to the shell script using the command:
   - `chmod +x chronogram.sh`

3. **Execute the Script:**
   Run the script by typing the following command into the terminal:
   - `./chronogram.sh`

4. **Provide Input Data:**
   When prompted, provide the following inputs:

  4.1 **Year for the Gantt Chart:**
     Enter the year for the Gantt Chart. If you leave this empty, the current year will be used.
     
      - Prompt: Add the year for the Gantt Chart (leave empty if using current year):
      - Example: 2024`

  4.2 **Starting Week Date:**
     Enter the starting week in MM/DD format. If you leave this empty, instead of dates, it will show Month and Week counters.
      
      - Prompt: Add the starting week (MM/DD) (leave empty if not):
      - Example: `05/01`
      
   4.3 **Priorities for Tasks:**
      Use Yes or No option if you would like to enable Priorities for each Task.

      - Prompt: Do you want to add priorities for the tasks? (yes or no):
      - Example: `Yes`

   4.4 **List of Milestones:**
      Enter the list of milestones as comma-separated values.

      - Prompt: Enter the list of milestones (as comma-separated values), or leave empty:
      - Example: `Requirements Gathering, Design, Development, Testing, Deployment, Maintenance`

   4.5 **Tasks for Each Milestone:**
      For each milestone, enter the list of tasks as comma-separated values.

      - Prompt: Enter the list of tasks for Current Milestone (as comma-separated values):
      - Example for "Requirements Gathering": `Initial Meeting, Stakeholder Interviews, Requirements Documentation`

  4.6 **List of Hours for Tasks:**
     For each milestone, enter the hours for each task as comma-separated values, using integers or floats.

      - Prompt: Enter the hours for tasks under Current Milestone (as comma-separated values):
      - Example for "Requirements Gathering" tasks: `10, 15, 20`
      
   4.7 **List of Priorities for Tasks:**
      If you enabled Priorities, then for each milestone, enter priorities for each task as comma-separated values, using the options: Low, Medium, High.

      - Prompt: Enter the priority for tasks under Current Milestone (Low, Medium, High) (as comma-separated values):
      - Example for "Low, Medium, High `
  
5. **Access the Chronogram:**
   After providing the input, the script will generate two files in the same directory:
   
      - `chronogram.xlsx: An Excel file with the Gantt Chart, Project Proposal and RACI Tables sheets.`
      - `chronogram.csv: A CSV file with the data used to generate the Gantt Chart.`

6. **Open and view the Chronogram:**
  Open the chronogram.xlsx file in Excel to view your visual chronogram with three sheets, Gantt Chart Weekly, Gantt Chart Monthly,Project Schedule and RACI Table.
  - 
      ![Chronogram Weekly](./Gantt_Chart_Weeks.png)

      ![Chronogram Monthly](./Gantt_Chart_Months.png)

      ![Project Schedule](./Project_Schedule.png)

      ![Raci Table](./RACI_Table.png)

## Sample Input:

```bash
motreto@Jorge:/mnt/c/Users/execu/ScheduleChronogram$ 
./chronogram.sh
Add the year for the Gantt Chart (leave empty if using current year):
Input: 2024

Add the starting week (MM/DD) (leave empty if not):
Input: 05/10

Do you want to add priorities for the tasks? (yes or no): Yes

Enter the list of milestones (as comma-separated values), or leave empty:
Input: Requirements Gathering, Design, Development, Testing, Deployment, Maintenance 


Adding tasks for Milestone: Requirements Gathering
Enter the list of tasks for Requirements Gathering (as comma-separated values):
Input: Initial Meeting, Stakeholder Interviews, Requirements Documentation 

Enter the hours for tasks under Requirements Gathering (as comma-separated values):
Input: 10 15 20 

Enter the priority for tasks under Requirements Gathering (Low, Medium, High) (as comma-separated values):
Input: Medium, Low, Low 

Adding tasks for Milestone: Design
Enter the list of tasks for Design (as comma-separated values):
Input: System Architecture Design, Database Schema Design, UI/UX Design 

Enter the hours for tasks under Design (as comma-separated values):
Input: 30 20 25 

Enter the priority for tasks under Design (Low, Medium, High) (as comma-separated values):
Input: Low, Low, Low 

Adding tasks for Milestone: Development
Enter the list of tasks for Development (as comma-separated values):
Input: Frontend Development, Backend Development, Integration 

Enter the hours for tasks under Development (as comma-separated values):
Input: 40 50 30 

Enter the priority for tasks under Development (Low, Medium, High) (as comma-separated values):
Input: Medium, Medium, Medium 

Adding tasks for Milestone: Deployment
Enter the list of tasks for Deployment (as comma-separated values):
Input: Prepare Deployment Environment, Deployment, Post-Deployment Verification

Enter the hours for tasks under Deployment (as comma-separated values):
Input: 15 10 10

Enter the priority for tasks under Deployment (Low, Medium, High) (as comma-separated values):
Input: Medium, Low, Low

Adding tasks for Milestone: Maintenance
Enter the list of tasks for Maintenance (as comma-separated values):
Input: Bug Fixing, Performance Tuning, User Training

Enter the hours for tasks under Maintenance (as comma-separated values):
Input: 20 15 10

Enter the priority for tasks under Maintenance (Low, Medium, High) (as comma-separated values):
Input: High, High, Low


Do you want to add the names for the roles in the RACI Table? (yes or no): Yes
Enter the name for Product Owner: John Smith
Enter the name for Business Analyst: Jane Doe
Enter the name for Financial Lead: Emily Davis
Enter the name for Design Director: Michael Brown
Enter the name for CRM Lead: Olivia Johnson
Enter the name for Head of CRM: Liam Williams
Enter the name for Senior Stakeholder*: Sophia Martinez
Enter the name for Senior Stakeholder**: William Garcia
Enter the name for AGENCY: Creative Solutions Inc.
```

## Prerequisites
### Python:
Please note that Python 3 must be installed on your system to use this script. If you do not have Python 3, please install it from the [official Python website](https://www.python.org/) or use your system's package manager.

### Python Libraries:
Before running this script, you must have the following Python libraries installed:
If you do not have Python 3, please install it from the official Python website or use your system's package manager.

- `pandas`
- `openpyxl`
- `re`
- `os`

You can install these libraries using pip with the following command:

- `pip install pandas`
- `pip install openpyxl`
- `pip install re`
- `pip install os`
