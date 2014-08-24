ApexExercise
============

Apex Interview Exercise

Description
===========
This is a bare-bones response to the following requirements. Bare-bones means not only did I do minimal additional work, but also that I did not remove un-used elements of the template project.

Requirements
============
Using the SpreadsheetBuilder class above and the AdventureWorks database (freely available from MS), please develop a web application with a single page that does the following:
+ Requests from the user start and end date for the report
+ Defaults (i.e., on the first page load) the start date to be the beginning of the previous month (e.g., if today is July 15, then the start date would be June 1)
+ Defaults the end date to be the end of the previous month (e.g., if today is July 15, then the start date would be June 30)
+ The table below will be filtered by due date to be between the start and end date provided by the user (i.e., by default, if today is July 15, the table would only show invoices dated between June 1 and June 30)
+ Upon clicking the Submit button, shows in the HTML the first 15 rows of the table below
+ Upon clicking the export button, generates an Excel Spreadsheet with all rows of the table below
 
The application can use any ASP.NET framework you wish (Web Forms, Web API, MVC), but must be written in C#.  It must use Linq2SQL, EF6, or some other freely available LINQ based ORM for data query.  The preferred deliverable would be a link to a GitHub repository where you have checked in your work.

Notes
=====
+ The database I used only goes to August of 2004.
+ The database I used did not include StoreID values, so I've commented out that join in the code-behind.
+ The exercise works correctly on my machine. :P
