# BingR
Bing Rewards using powershell

BingR

BingR is an Bing Search automation tool to gather Bing Reward points on a regular (daily) basis.

It has several features including command line and ini argument options, debugging and logging and can be used as a random password generator.

Included Files:
BingR ScheduleTask.xml - Example schedule file to be imported (loaded) into the task scheduler automating the process to run daily at 6 am.
BingR.exe - the program
BingR.ini - a initialization file to control the program actions
BingR.log - if logging is turned on this file will be created and will capture the program activities.

How to use:
There are two ways you may run this program; from the command prompt or as a scheduled task.  Placing a shortcut on the desktop can simlify the run process to a click or double click to launch also.

Run BingR from a command line with arguments or place the arguments in an INI file.
optional arguments are:
/l - logging
/d - debuging
fn:bing - use the program in bing search mode
un:username - include a username to log into bing search as
pw:password - include a password to log into bing search with

Using the un & pw options in the ini file is risky as hackers could gain access to your PC and find this file and use it to hack your Bing Rewards account.  If you use the same password on multiple sites this could also compromise those sites.


Open the task scheduler and import the BingR ScheduleTask.xml file to automate the program execution.  This automation is setup to run with logging turned on.  You need to modify the Command, Arguments & WorkingDirectory entries to the proper file location and arguments settings before importing it to the task scheduler or you can modify the task after the fact.

      <Command>C:\Path\to\executable\file\BingR.exe</Command>
      <Arguments>/l</Arguments>
      <WorkingDirectory>C:\Path\to\executable\file</WorkingDirectory>

There are two ways to import the task schedule once these adjustments have been made.  You can open the task scheduler and import the file there or open a command prompt and enter the following command;

schtasks /Create /XML <xmlfile> /TN <taskname>
 - where xmlfile is BingR ScheduleTask.xml - you need to wrap the filename in quotes if there are spaces like I have in this filename.
 - where taskname is BingR

schtasks /Create /XML "BingR ScheduleTask.xml" /TN BingR


