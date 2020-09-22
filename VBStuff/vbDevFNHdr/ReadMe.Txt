;Updated. Version 1.1.1. Moved a line to allow Friend Property functions
'--------------------------------------------------------------------------------
' the Original incarnation of this program was grabbed from the web years and
' years and years ago. I have fleshed it out with a lot of new functionality,
' simplified a lot of procedures, and cleaned up the code. I wish I could remeber
' who had originally posted the original incarnation of useful utility.
'
' Anyway, it makes a great beginner project. I did this back in '98 when I was
' first learning VB (I came over from C++ and Fortran). Oh, since I'm a C++
' developer, let me tell you about my fellow bozos who stick their nose up at
' VB: Once I realized I could write an application in VB in a day that would take
' me close to a month to do in C++, I became an instant VB convert. --David Goben
'--------------------------------------------------------------------------------

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'NOTE: In the Connect.dsr file, in the AddinInstance_OnConnection() event,
'be sure to chage the default caption in the AddToAddInCommandBar()
'function to the name that you want to see in the Add-Ins menu in the IDE
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'****************************************************************************
'COMPILATION NOTES:
'
' 1. Compile the VBDevFNHdr.DLL either to the project folder, or to your
'    \Windows\System32 (Windows\System for Win95\98\ME) folder. If compiled
'    to the project folder, copy it to your System32 (or System) folder.
'
' 2. Exit VB, then re-enter it. Fram the "Add-Ins" menu,choose
'    "Add-In Manager...". Find the "VB Development Insert Routine Header" entry
'    and insure that the "Loaded/Unloaded" and "Load on Startup" are checked,
'    then hit OK.
'
' You should now see "VB Development Insert Routine Header" in the Add-Ins
' menu. Select it anytime you want to close all open forms and code modules.
'----------------------------------------------------------------------------
' IMPORTANT NOTE:
' If you are updating an Add-in, BE SURE to first unclock the Loaded/Unloaded
' open in the Add-In Manager (it doesn't hurt to also uncheck Load on Startup.
' This way you can write the new DLL without it yelling at you about access
' being denied because it is in use.
'
' Also, I've noticed that when you exit VB after compiling an Add-in, it
' suffers a small (but not harmful) conniption and issues a warning. Don't
' sweat it. You can cheat by opening up a different project and then exiting.
'****************************************************************************
