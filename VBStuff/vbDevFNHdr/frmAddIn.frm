VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "My Add In"
   ClientHeight    =   3195
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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

Public VBInstance As VBIDE.VBE
Public Connect As Connect

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Allow tooltips for entries that exceed the list width
'*******************************************************************************
Private Sub Form_Load()
  Dim W As Window
'
' scvan through each window defined in the VB project and find one that is:
'   1: Open
'   2: Active
'      Then find the cursor location, and from there the routine it is contained it
'
  On Error GoTo Oops
  For Each W In VBInstance.Windows
    If W.Visible = True And W.Type = vbext_wt_CodeWindow And W.Caption = VBInstance.ActiveCodePane.Window.Caption Then
      Dim S As String, StL As Long, StC As Long, EnL As Long, EnC As Long, Typ As String
      Dim STLP As Long, StLG As Long, StLL As Long, StLS As Long
      Dim Pname As String, Ptype As String, I As Integer, Idx As Integer, TS As String
     
      With VBInstance.ActiveCodePane
        .GetSelection StL, StC, EnL, EnC                    'get current cursor position
        StC = -1                                            'set column to beginning
        With .CodeModule
          Pname = .ProcOfLine(StL, vbext_pk_Proc)           'get procedure name from line
          If Len(Pname) Then                                'found one near it
            STLP = .CountOfLines                            'get number of lines in procedure
            StLG = STLP                                     'see if we are PROC/GET/SET/LET
            StLS = STLP
            StLL = STLP
            On Error Resume Next
            STLP = .ProcBodyLine(Pname, vbext_pk_Proc)      'get first line of a Procedure
            StLG = .ProcBodyLine(Pname, vbext_pk_Get)       'get first line of a GET
            StLS = .ProcBodyLine(Pname, vbext_pk_Set)       'get first line of a SET
            StLL = .ProcBodyLine(Pname, vbext_pk_Let)       'get first line of a LET
            On Error GoTo 0
            If STLP < StL And STLP > StC Then StC = STLP    'assume Procedure (Sub/Function)
            If StLG < StL And StLG > StC Then StC = StLG    'assume GET
            If StLS < StL And StLS > StC Then StC = StLS    'assume SET
            If StLL < StL And StLL > StC Then StC = StLL    'assume LET
          End If
'
' if we cannot find anything to use, then exit without doing anything
'
          If StC < 0 Then
            Unload frmDeclare
            Unload Me
            Exit Sub
          End If
'
' get first line of code block and find out its visiblility declaration, and strip them
'
          S = Trim$(LCase$(.Lines(StC, 1)))               'get first line of block
          
          If Left(S, 8) = "private " Then S = Mid(S, 9)   'skip Private
          If Left(S, 7) = "public " Then S = Mid(S, 8)    'skip Public
          If Left(S, 7) = "friend " Then S = Mid(S, 8)    'skip Friend
          If Left(S, 9) = "property " Then S = Mid(S, 10) 'skip Property
'
' now determine what type of routine we are dealing with
'
          Select Case Left$(S, 3)
            Case "fun"                      'set up for function
              Ptype = "Function "
              Typ = "Function Name"
            Case "sub"                      'set up for subroutine
              Ptype = "Subroutine "
              Typ = "Subroutine Name"
            Case "get"                      'set up for Property Get
              Ptype = "Get "
              Typ = "Property Get Name"
            Case "set"                      'set up for Property Set
              Ptype = "Set "
              Typ = "Property Set Name"
            Case "let"                      'set up for Property Let
              Ptype = "Let "
              Typ = "Property Let Name"
            Case Else                       'who knows what this puppy is....
              Unload frmDeclare
              Unload Me
              Exit Sub
          End Select
'
' build input dialog form, and set default text for each textbox on the Declare form
'
          With frmDeclare
            .lblVersion.Caption = "v" & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
            .Caption = "Declare " & Ptype & Pname
            .lblName.Caption = Pname
            .lblType.Caption = Typ
            .txtAuthor.Text = GetSetting(App.Title, "Settings", "Author", vbNullString)
            .txtPurpose.Text = GetSetting(App.Title, "Settings", "Purpose", vbNullString)
            .txtInputs.Text = GetSetting(App.Title, "Settings", "Inputs", vbNullString)
            .txtOutputs.Text = GetSetting(App.Title, "Settings", "Outputs", vbNullString)
            .txtAssumes.Text = GetSetting(App.Title, "Settings", "Assumes", vbNullString)
            .txtOther.Text = GetSetting(App.Title, "Settings", "Other", vbNullString)
            .txtSideEffects.Text = GetSetting(App.Title, "Settings", "SideEffects", vbNullString)
            frmDeclare.vCancel = False
            .Show vbModal
'
' if user cancelled...
'
            If .vCancel Then Exit For
'
' update Registry to textbox contents
'
            Call SaveSetting(App.Title, "Settings", "Author", .txtAuthor.Text)
            Call SaveSetting(App.Title, "Settings", "Purpose", .txtPurpose.Text)
            Call SaveSetting(App.Title, "Settings", "Inputs", .txtInputs.Text)
            Call SaveSetting(App.Title, "Settings", "Outputs", .txtOutputs.Text)
            Call SaveSetting(App.Title, "Settings", "Assumes", .txtAssumes.Text)
            Call SaveSetting(App.Title, "Settings", "Other", .txtOther.Text)
            Call SaveSetting(App.Title, "Settings", "SideEffects", .txtSideEffects.Text)
'
' build procedure header
'
            S = "'*******************************************************************************" & vbCrLf
            S = S & ExpandString(Pname, Ptype & "Name")
            S = S & ExpandString(.txtAuthor.Text, "Author")
            S = S & ExpandString(.txtPurpose.Text, "Purpose")
            S = S & ExpandString(.txtInputs.Text, "Inputs")
            S = S & ExpandString(.txtOutputs.Text, "Outputs")
            S = S & ExpandString(.txtAssumes.Text, "Assumes")
            S = S & ExpandString(.txtOther.Text, "Other Modules Used")
            S = S & "'*******************************************************************************"
          End With
          .InsertLines StC, S   'stuff this text block to the top of the routine
        End With
      End With
      Exit For                  'all done
    End If
  Next W
  
Oops:
  Unload frmDeclare
  Unload Me
End Sub

'*******************************************************************************
' Function Name     : ExpandString
' Purpose           : Format a string for the heading if it contains data
'*******************************************************************************
Private Function ExpandString(InLine As String, CmtType As String) As String
  Dim S As String, I As Integer
  
  S = vbNullString
  If Len(InLine) Then
    If Len(CmtType) < 18 Then
      S = S & "' " & CmtType + String(18 - Len(CmtType), " ") & ": " & InLine
    Else
      S = S & "' " & CmtType & ": " & InLine
    End If
    I = InStr(S, vbCrLf)
    Do While I
      S = Left(S, I + 1) & "'" & String(19, " ") & ": " & Mid(S, I + 2)
      I = InStr(I + 2, S, vbCrLf)
    Loop
    S = S & vbCrLf
  End If
  ExpandString = S
End Function
