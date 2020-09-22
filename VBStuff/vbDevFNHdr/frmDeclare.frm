VERSION 5.00
Begin VB.Form frmDeclare 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Declaration"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   2100
      TabIndex        =   7
      Top             =   7680
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5340
      TabIndex        =   9
      Top             =   7680
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   7680
      Width           =   1275
   End
   Begin VB.TextBox txtOther 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   5505
      Width           =   4575
   End
   Begin VB.TextBox txtPurpose 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1125
      Width           =   4575
   End
   Begin VB.TextBox txtSideEffects 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   6600
      Width           =   4575
   End
   Begin VB.TextBox txtAssumes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   4410
      Width           =   4575
   End
   Begin VB.TextBox txtOutputs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3315
      Width           =   4575
   End
   Begin VB.TextBox txtInputs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2220
      Width           =   4575
   End
   Begin VB.TextBox txtAuthor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " NOTE: Leave entries blank if you do not want them added"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   195
      Left            =   1620
      TabIndex        =   20
      Top             =   8100
      Width           =   5025
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   8100
      Width           =   90
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Side Effects:"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   6600
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Other Components used:"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   5505
      Width           =   1755
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Assumes:"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   4410
      Width           =   675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Outputs:"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   3315
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Inputs:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   2220
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Purpose:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1125
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Author:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   510
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      Caption         =   "Routine Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   1065
   End
End
Attribute VB_Name = "frmDeclare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Public vCancel As Boolean   'TRUE when the user cancels from frmDeclare.frm

'*******************************************************************************
' Subroutine Name   : cmdCancel_Click
' Purpose           : User does not want to use this form (on second thought)
'*******************************************************************************
Private Sub cmdCancel_Click()
  vCancel = True
  Me.Hide
End Sub

'*******************************************************************************
' Subroutine Name   : cmdOK_Click
' Purpose           : Accept changes and stuff it to the routine's heading
'*******************************************************************************
Private Sub cmdOK_Click()
  vCancel = False
  Me.Hide
End Sub

'*******************************************************************************
' Subroutine Name   : cmdClear_Click
' Author            : David Goben
' Purpose           : Clear all textboxes on the form
'*******************************************************************************
Private Sub cmdClear_Click()
  Me.txtAuthor.Text = vbNullString
  Me.txtPurpose.Text = vbNullString
  Me.txtInputs.Text = vbNullString
  Me.txtOutputs.Text = vbNullString
  Me.txtAssumes.Text = vbNullString
  Me.txtOther.Text = vbNullString
  Me.txtSideEffects.Text = vbNullString
  Me.txtAuthor.SetFocus
End Sub

'*******************************************************************************
' Subroutine Name   : txtAssumes_GotFocus
' Purpose           : Hightlight the contents of a textbox when it gets focus. This
'                   : allows the user to quickly delete iuseless data
'*******************************************************************************
Private Sub txtAssumes_GotFocus()
  HiLiteText Me.txtAssumes
End Sub

Private Sub txtAuthor_GotFocus()
  HiLiteText Me.txtAuthor
End Sub

Private Sub txtInputs_GotFocus()
  HiLiteText Me.txtInputs
End Sub

Private Sub txtOther_GotFocus()
  HiLiteText Me.txtOther
End Sub

Private Sub txtOutputs_GotFocus()
  HiLiteText Me.txtOutputs
End Sub

Private Sub txtPurpose_GotFocus()
  HiLiteText Me.txtPurpose
End Sub

Private Sub txtSideEffects_GotFocus()
  HiLiteText Me.txtSideEffects
End Sub
