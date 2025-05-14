# EPE-System-Excel-
Employee Performance Evaluation System Excel

Made by Atun, Jaspher G.

This system evaluates the performance of an employee based on four criteria; each with five subcriteria. The evaluation scale ranges from Excellent (5) to Extremely poor (1). There are four evaluators with different weighted percentage: Superior (40%), Colleague (30%), Subordinate (20%), Self-Assessment (10%).

There are three workbooks: “Employee List”, “Performance Evaluation” & “Employees”.

"Employee List" Workbook:

Sheet of list of all Employees that is hyperlinked to the “Employees” Workbook.

An “Add Employee” Sheet to add employees to the “Employees” Workbook.
 
    Two buttons: “Upload Picture” & “Save”
      1. Upload Picture – Inserts a picture in a certain range of cells and
        autofits it.
      2. Save:
        o Copies the Sheet to the Employees workbook,
        o Submits the values to the Employee List Sheet,
        o Hyperlinks the Employee ID on the saved sheet on the
          Employees workbook,
        o Change the sheet name based on the Employee ID, protects
          the sheet and removes the buttons.
        o Option to clear the current details in the fields.
        o Dependent Drop-Down List: The Position/Title is dependent on the dropdown list of the Department cell.

"Employees" Workbook:

Workbook for all employees. Shows the details and Evaluation Result.


    One button: Evaluate
    
    1. Opens the Performance Evaluation Workbook and
    2. Transfers the employee details to it.
    
"Performance Evaluation" Workbook:

Three Sheets; Employee, Overview & Dashboard

Employee Sheet:

    ▪ Evaluation sheet to evaluate the employee.
    ▪ Evaluation scales are made with VLOOKUP along with the Final Score &
      Remarks.
    ▪ SUMPRODUCT in the Combined Score along with some arithmetic
      operators and SUM function for the final score.
    ▪ Summary: Shows a dashboard of scores of criteria and per evaluator. 
    ▪ Two Buttons:
      1. Save to PDF – Saves the evaluation to a pdf format with the name
      of the employee on the pdf file.
      2. Submit:
          o Submits the employee details along with the scores to the
            Overview sheet.
          o Hyperlinks the employee ID to the Employee Workbook.
          o Copies the evaluation score to the Employee Workbook and
            formats it.
  Overview Sheet:
  
      o Shows all the evaluated employees with their scores.
      o Employee ID hyperlinked to the Employees Workbook.
      
  Dashboard Sheet:
  
      o Shows the top employees and department.
      o Refresh button.
