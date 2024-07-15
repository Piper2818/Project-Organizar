# Project-Organizar 
This is a generalized version of a VBA code that I am working on to help Sort, Organize, and Prioritize my projects and the tasks associated with them on a daily basis at work.The corresponsing excel workbook contains 3 types of worksheets: 

Sheet1: Insert/search 
  This sheet is where the user can input new projects, update existing projects, and search for existing projects. 
  I am currently working on a way automate the insertion of new tasks and updates from co-workers (which typically come in the form of email) 
  
Sheet2: To do 
  This sheet is the to do sheet, it hold all of the projects which are currently in progress sorted first by date recived (oldest to newest), next by status (action required or no action required), and lastly by a Flag (has the project been flagged by the user or not). Once a project has been set to closed and the page has been refreshed the job will be removed from the To do list and stored for a certain period of time on one of the following sheets. 
  
Sheet3 - #: Records
  All sheets following the to do sheet are for the sake of records which are organized by job type. They will contain all jobs which are in progress and closed orgnaized by the date they where recived in chronological order. These record will be kept for any period of time that the user defines before deletion. 
    

This excel document/VBA code is just a proof of concept for a larger project which I hope to implement in a more sutaible system. This code is undergoing frequent changes/updates and there are still parts of it that are in progress. 

Tech issues relating to Excel: 

1) Ideally, multiple people will need to be working off of and makeing frequent changes to these lists throughout the day and likely at the same time which could cause excel to run poorly and/or crash do to the number of macros in use and the shear amount of data that will need to be held in this sheet.

2) Most of our information (new jobs, job updates, etc.) come in through email and there is no standard way of sending most of this information so this excel sheet will still require a lot of manual maintenance to keep it up to date. Once this is built into a system ideally we can start getting those updates throught the system rather than relying on email. 
