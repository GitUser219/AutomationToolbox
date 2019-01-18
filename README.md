# Automation Toolbox

A collection of two tools: a preventative maintenance PDF creation tool and a preventative maintenance report creation tool

### Prerequisites

A minimum [Java Runtime Environment](http://www.oracle.com/technetwork/java/javase/downloads/jre8-downloads-2133155.html) version of 1.8.0 is required to run AutomationToolbox.exe

### Using the preventative maintenance PDF creation tool

1. Run AutomationToolbox.exe
2. Click the Preventative Maintenance PDF Creation button
3. Click Select spreadsheet and choose Spreadsheet.xlsx from Test Files.zip
4. Click Select images location and choose the location of the files from Test Files.zip
5. Click Select destination and choose where you would like to save the output
6. Type a password to protect the PDF with
7. Click Generate the file and the PDF will be created in the file destination that you chose

### Using the preventative maintenance report creation tool

1. Run Automation Toolbox.exe
2. Click the Preventative Maintenance Report Creation button
3. Click Select the detail report and choose Detail Report.csv from Test Files.zip
4. Click Select the summary report and choose Summary Report.csv from Test Files.zip
5. Click Select the file destination and choose where you would like to save the output
6. Click Generate the file and Report.xlsx will be created in the file destination that you chose

### Built With

* [Java](https://www.java.com/en/) - a general-purpose computer-programming language that is concurrent, class-based, object-oriented, and specifically designed to have as few implementation dependencies as possible
* [Eclipse](https://www.eclipse.org) - An integrated development environment (the most widely used Java IDE)
* [Apache POI](https://poi.apache.org/) - Java libraries for reading and writing files in Microsoft Office formats, such as Word, PowerPoint and Excel
* [iText PDF](https://itextpdf.com/) - A library for creating and manipulating PDF files in Java and .NET
* [openCSV](https://sourceforge.net/projects/opencsv/) - A Simple CSV Parser for Java under a commercial-friendly Apache 2.0 license
* [Launch4j](http://launch4j.sourceforge.net/) - Cross-platform Java executable wrapper for creating lightweight Windows native executable files

### Development

The following instructions will help you to set up an environment for development and testing:

1. Download [AutomationToolbox-master.zip](https://github.com/GitUser219/AutomationToolbox/archive/master.zip) and extract all of the files to some location on your machine
2. Extract all of the files from External Jar Files.zip and Images.zip to some location on your machine
3. Ensure that you have installed a minimum [Java Runtime Environment](http://www.oracle.com/technetwork/java/javase/downloads/jre8-downloads-2133155.html) version of 1.8.0
4. Download and install [Eclipse IDE for Java Developers](https://www.eclipse.org/downloads/)
5. Run Eclipse and click File, New, Java Project
6. Name the project whatever you like, click Next, click Create new source folder, name it resources, and click Finish twice
7. Right click on the project, click Build Path then Add External Archives, and add all of the files that you extracted from External Jar Files.zip
8. Right click src, click New then Package, name it automation.toolbox, click Finish
9. Drag AutomationToolbox.java to the automation.toolbox package and click OK
10. Right click resources, click New then Package, name it images, click Finish
11. Drag all of the files from Images.zip, except for black_icon.ico, to the images package and click OK
12. The environment is now set up; right click AutomationToolbox.java (in Eclipse) and click Run As then Java Application to run the program

### Deployment

If you make a change to the code and you would like to create a new executable file, follow these steps:

1. After running the program with no errors, right click on the Automation Toolbox project and click Export
2. Select Runnable JAR file in the Java folder and set the Launch configuration to AutomationToolbox - YOUR_PROJECT_NAME
3. Choose the destination for the JAR file and set the file name to AutomationToolbox.jar, click Finish twice
4. If there were no issues, you will have AutomationToolbox.jar saved on your computer
5. Install and run the [Launch4j Executable Wrapper](https://sourceforge.net/projects/launch4j/) tool
6. Make the following changes in the Basic tab of Launch4j
   - Output file: choose the file destination and set the file name to AutomationToolbox.exe
   - Jar: select the AutomationToolbox.jar file
   - Icon: select black_icon.ico from Images.zip
7. Make the following changes in the JRE tab of Launch4j
   - Min JRE version: 1.8.0
8. Click Build wrapper (gear icon in the upper left-hand corner)
9. Save a configuration file with any name and AutomationToolbox.exe will be created
10. You can delete AutomationToolbox.jar and the configuration file once the executable has been created
11. Run AutomationToolbox.exe to test that it is working properly

### Authors

* **Jonathan Nowak** - [GitUser219](https://github.com/GitUser219) - **Last Revised 1/17/19**
