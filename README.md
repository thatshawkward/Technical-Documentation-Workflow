# Technical-Documentation-Workflow
Google Apps script for Sheets to assist maintaining a documentation workflow and provide a data source for a documentation dashboard. It breaks out a master spreadsheet tracking documentation data into Current Projects, Backlog, and On Hold, which are locked for editing and used only as reference. Items marked as complete are either deprecated or don’t require maintenance and subsequently sorted to an Archive sheet. A separate function is included to calculate the next review date based on the document completion date. 

Intended to solve the problem of being asked to create a dashboard to visualize workload and progress on documentation, but you’re a Technical Writer who doesn’t have access to visualization tools and at the same time it needs to be practical for tracking work, otherwise you’re creating even more work for yourself. 

Column references in the script are hardcoded. Master sheet columns include: ID, Title, Description, Category, Priority, Initial LoE, Due Date, Status, Custom Status, Dept Head/Approver, Start Date, Est. Completion Date, Published/Last Modified, Next Review, Lifecycle Stage, Maintenance Cadence, Team, Submission Date
  
