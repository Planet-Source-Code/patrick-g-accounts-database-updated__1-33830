VERSION 5.00
Begin VB.Form FrmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "First Time Setup"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3465
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Remember my username"
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   3480
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   3255
   End
   Begin VB.TextBox Text5 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   2400
      Width           =   2055
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Female"
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   1680
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Male"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1440
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label7 
      Caption         =   "Please verify that all the above information is correct.  If all is correct please click ""Save"""
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2400
      Picture         =   "Form3.frx":65AA
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label6 
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Username:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Gender:"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Age:"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Lastname:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Firstname:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "FrmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim QuitNow As String
QuitNow = MsgBox("Are you sure you want to cancel the startup process?", vbYesNo Or vbQuestion, "Cancel?")

If QuitNow = vbYes Then
    End
Else
    Exit Sub
End If

End Sub

Private Sub Command2_Click()

Dim Gender As String

If Text1.Text = "" Then
    Text1.SetFocus
    Exit Sub
End If

If Text2.Text = "" Then
    Text2.SetFocus
    Exit Sub
End If

If Text3.Text = "" Then
    Text3.SetFocus
    Exit Sub
End If

If Text4.Text = "" Then
    Text4.SetFocus
    Exit Sub
End If

If Text5.Text = "" Then
    Text5.SetFocus
    Exit Sub
End If

SaveSetting "Accounts", "Profile", "Firstname", Text1.Text
SaveSetting "Accounts", "Profile", "Lastname", Text2.Text
SaveSetting "Accounts", "Profile", "Age", Text3.Text

If Option1.Value = True Then
    Gender = "Male"
Else
    Gender = "Female"
End If

SaveSetting "Accounts", "Profile", "Gender", Gender
SaveSetting "Accounts", "Profile", "Username", Text4.Text
SaveSetting "Accounts", "Profile", "Password", Text5.Text
SaveSetting "Accounts", "Profile", "Remember", Check1.Value

MsgBox "Your new profile has now been setup!" & vbCrLf & vbCrLf & "Welcome " & Text1.Text & " " & Text2.Text, vbInformation, "Profiles"
FrmAdmin.Show
FrmAdmin.Label25.Caption = "Automaticly Login: No"
Unload Me

End Sub


Private Sub Form_Load()

On Error Resume Next
Dim IfGender, IfRemember As String

Text1.Text = GetSetting("Accounts", "Profile", "Firstname")
Text2.Text = GetSetting("Accounts", "Profile", "Lastname")
Text3.Text = GetSetting("Accounts", "Profile", "Age")
IfGender = GetSetting("Accounts", "Profile", "Gender")

If GetSetting("Accounts", "Profile", "Gender") = "Male" Then
    Option1.Value = True
    Option2.Value = False
Else
    Option1.Value = False
    Option2.Value = True
End If
    
Check1.Value = GetSetting("Accounts", "Profile", "Remember")
Text4.Text = GetSetting("Accounts", "Profile", "Username")
Text5.Text = GetSetting("Accounts", "Profile", "Password")

End Sub

Private Sub Text3_Change()

If IsNumeric(Text3.Text) = False Then
    MsgBox "Must be numberic characters", vbExclamation, "Opps"
    Text3.Text = ""
End If
    
End Sub
