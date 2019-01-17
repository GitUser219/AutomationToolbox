# Automation Toolbox

## A preventative maintenance PDF creation tool

**How to use this tool:**

1. Run Automation Toolbox.exe
2. Click the Preventative Maintenance PDF Creation button
3. Click Select spreadsheet
4. Navigate to Automation Toolbox\Assets\Test Files and select Spreadsheet.xlsx
5. Click Select images location
6. Navigate to Automation Toolbox\Assets\Test Files and click Open
7. Click Select destination and choose where you would like to save the output
8. Type a password to protect the PDF with
9. Click Generate the file and the PDF will be created in the file destination that you chose

## A preventative maintenance report creation tool

**How to use this tool:**

1. Run Automation Toolbox.exe
2. Click the Preventative Maintenance Report Creation button
3. Click Select the detail report
4. Navigate to Automation Toolbox\Assets\Test Files and select Detail Report.csv
5. Click Select the summary report
6. Navigate to Automation Toolbox\Assets\Test Files and select Summary Report.csv
7. Click Select the file destination and choose where you would like to save the output
8. Click Generate the file and Report.xlsx will be created in the file destination that you chose

### Built With

* [Eclipse](https://www.eclipse.org) - An integrated development environment (the most widely used Java IDE)
* [Apache POI](https://poi.apache.org/) - Java libraries for reading and writing files in Microsoft Office formats, such as Word, PowerPoint and Excel
* [iText PDF](https://itextpdf.com/) - A library for creating and manipulating PDF files in Java and .NET
* [openCSV](https://sourceforge.net/projects/opencsv/) - A Simple CSV Parser for Java under a commercial-friendly Apache 2.0 license
* [Launch4j](http://launch4j.sourceforge.net/) - Cross-platform Java executable wrapper for creating lightweight Windows native executable files

### Development Environment

The following instructions will help you to set up an environment for development and testing:

1. Download [Automation Toolbox.zip](https://github.com/GitUser219/AutomationToolbox) and extract all of the files to some location on your machine
2. Ensure that you have installed a minimum [Java Runtime Environment](http://www.oracle.com/technetwork/java/javase/downloads/jre8-downloads-2133155.html) version of 1.8.0
3. Download and install [Eclipse IDE for Java Developers](https://www.eclipse.org/downloads/)
4. Run Eclipse and click File, New, Java Project
5. Name the project Automation Toolbox, click Next, click Create new source folder, name it res, and click Finish
6. Right click on the Automation Toolbox project, click Build Path, Add External Archives, navigate to Automation Toolbox\Assets\External Jar Files and select all of the files in this location
7. Right click src, select New, then Package, name the package jn.automation, click Finish
8. Copy AutomationToolbox.java from Automation Toolbox\Assets to the jn.automation package
9. Right click res, select New, then Package, name the package imgs, click Finish
10. Copy all of the files in Automation Toolbox\Assets\Images except for black_icon.ico to the imgs package
11. Double click on the AutomationToolbox.java to open it in the IDE and click the green and white play button to run the program

### Publishing Changes

If you make a change to the code and you would like to create a new executable file, follow these steps:

1. After running the program with no errors, right click on the Automation Toolbox project and click Export
2. Select Runnable JAR file in the Java folder and set the Launch configuration to AutomationToolbox - Automation Toolbox
3. Choose the destination for the JAR file and set the file name to Automation Toolbox.jar, click Finish
4. If there were no issues, you will have Automation Toolbox.jar saved on your computer
5. Install and run the [Launch4j Executable Wrapper](https://sourceforge.net/projects/launch4j/) tool
6. Basic tab
   - Output file: choose the file destination and set the file name to Automation Toolbox.exe
   - Jar: select the Automation Toolbox.jar file
   - Icon: select Automation Toolbox\Assets\Images\black_icon.ico
7. JRE tab
   - Min JRE version: 1.8.0
8. Click Build wrapper (gear icon)
9. Save a configuration file and your executable will be created
10. You can delete the configuration file once the executable has been created

### Authors

* **Jonathan Nowak** - *Initial work*
