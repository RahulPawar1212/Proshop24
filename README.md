# proshop24
Conssignment report automation

Steps to run application
* 	Go into releases
*	Download ‘consignment automation’ folder from the latest release
*	Extract zip in a folder (You can place it anywhere)
*	Click on cli.exe (You can create shortcut of cli.exe to desktop)

Check new skus.
*	Click on new skus button.
*	A dialogue box will appear to select GRN File – select GRN file
*	A dialogue box will appear to select Reports folder -	 select Reports folder
*	A pop up box will appear to select dates – select start and date for the period you want check stocks 
*	Click ok
*	This step will generate a new skus file into reporting folder.

Add newly find skus.
*	Add newly find skus into reporting file (In required tabs as per your business logic. App will run even if it is added in just one tab.)


Run Stock update
*	Click on 'stock update' button
*	Dialogue  box will pop up to select CRN file - Select GRN file.
*	Dialogue box will pop up to select Report folder - Select folder which contains all the reports which you want to process for current week.
*	After that a popup box will appear to select dates – select date for those you wants to generate report
*	Click ok
*	Once reports gets generated application will show ‘Reports generated’ message.
*	This step will create a folder named ‘Stock Updates Output’ into the reports folder and it will put updated files into that folder.

 Run sales Update
*	Click on 'Sales update' button.
*	Dialogue  box will pop up to select sales data file - Select sales data file.
*	Dialogue box will pop up to select Report folder - Select ‘Stock Updates Output’ folder generated in above step.
*	After that a popup box will appear to select dates – select date for those you wants to generate report.
*	Click ok.
*	Once reports gets generated application will show ‘Reports generated’ message.
*	This step will create a folder named ‘Final Output’ into the reports folder and it will put updated files into that folder.



Notes...
* Make sure that PC delimation format should be ',' comma.



