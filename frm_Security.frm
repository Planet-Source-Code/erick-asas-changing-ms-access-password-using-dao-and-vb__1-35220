VERSION 5.00
Begin VB.Form frm_Security 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change LogIn Password"
   ClientHeight    =   2295
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4215
   Icon            =   "frm_Security.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1355.962
   ScaleMode       =   0  'User
   ScaleWidth      =   3957.657
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1140
   End
   Begin VB.TextBox txtVerifyNewPass 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1740
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1170
      Width           =   2325
   End
   Begin VB.TextBox txtOldPass 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1740
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   390
      Width           =   2325
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2790
      TabIndex        =   3
      Top             =   1800
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1500
      TabIndex        =   5
      Top             =   1800
      Width           =   1140
   End
   Begin VB.TextBox txtNewPass 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1740
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   780
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   1740
      TabIndex        =   10
      Top             =   60
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Verify Password"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1230
      Width           =   1560
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   90
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   420
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   870
      Width           =   1080
   End
End
Attribute VB_Name = "frm_Security"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This sample program is designed by Erick Asas,
'SentrySoft2000 Ltd.

'The program is intended as an example of how to implement
'changing of MS password from VB.
'The program is free to use and comes as is with nowarranty
'
'clasas@mnd.philcom.com
'Initial UserName = Admin
'Initial Password = ""
Option Explicit
Private blnSuccess As Boolean

Private Sub cmdCancel_Click()
On Error GoTo ET
  Unload Me
  Exit Sub
ET:
    Select Case Err.Number
    Case Else
      MsgBox Err.Description, vbOKOnly + vbCritical, "Error # " & Err.Number
    End Select
    Err = False
    Exit Sub
End Sub

Private Sub cmdApply_Click()
On Error GoTo ET

If txtNewPass <> txtVerifyNewPass Then
  MsgBox "Verify the new password by retyping it in the Verify Box and clicking Apply button.", vbOKOnly + vbInformation, "Security Error!!!"
Else
      'Update password
      blnSuccess = UpdatePassword(strUID, txtOldPass, txtNewPass)
      MsgBox "The old password is successfully changed.", vbOKOnly + vbInformation, "Success!!!"
      txtOldPass = ""
      txtNewPass = ""
      txtVerifyNewPass = ""
      cmdApply.Enabled = False
End If
Exit Sub
ET:
    Select Case Err.Number
    Case Else
      MsgBox Err.Description, vbOKOnly + vbCritical, "Error # " & Err.Number
    End Select
    Err = False
    Exit Sub
End Sub

Private Sub cmdOk_Click()
On Error GoTo ET
    Unload Me
  Exit Sub
ET:
    Select Case Err.Number
    Case Else
      MsgBox Err.Description, vbOKOnly + vbCritical, "Error # " & Err.Number
    End Select
    Err = False
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo ET
  lblLabels(4).Caption = strUID
  lblLabels(4).FontBold = True
  Exit Sub
ET:
    Select Case Err.Number
    Case Else
      MsgBox Err.Description, vbOKOnly + vbCritical, "Error # " & Err.Number
    End Select
    Err = False
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ET
  Dim frm As Form
  
  For Each frm In Forms
    Unload frm
    Set frm = Nothing
  Next
  
  Exit Sub
ET:
    Select Case Err.Number
    Case Else
      MsgBox Err.Description, vbOKOnly + vbCritical, "Error # " & Err.Number
    End Select
    Err = False
    Exit Sub
End Sub

Private Sub txtNewPass_Change()
On Error GoTo ET
  cmdApply.Enabled = True
  Exit Sub
ET:
      Select Case Err.Number
      Case Else
        MsgBox Err.Description, vbOKOnly + vbCritical, "Error # " & Err.Number
      End Select
      Err = False
      Exit Sub
End Sub

Private Sub txtOldPass_Change()
On Error GoTo ET
  cmdApply.Enabled = True
  Exit Sub
ET:
      Select Case Err.Number
      Case Else
        MsgBox Err.Description, vbOKOnly + vbCritical, "Error # " & Err.Number
      End Select
      Err = False
      Exit Sub
End Sub

Private Sub txtVerifyNewPass_Change()
On Error GoTo ET
  cmdApply.Enabled = True
  Exit Sub
ET:
      Select Case Err.Number
      Case Else
        MsgBox Err.Description, vbOKOnly + vbCritical, "Error # " & Err.Number
      End Select
      Err = False
      Exit Sub
End Sub
