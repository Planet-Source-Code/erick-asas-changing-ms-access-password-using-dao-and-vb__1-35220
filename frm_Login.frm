VERSION 5.00
Begin VB.Form frm_Login 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2745
   ClientLeft      =   2835
   ClientTop       =   3195
   ClientWidth     =   4200
   ControlBox      =   0   'False
   Icon            =   "frm_Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1621.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3943.573
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   30
      TabIndex        =   4
      Top             =   30
      Width           =   4125
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1350
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1440
         Width           =   2325
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
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
         Left            =   2160
         TabIndex        =   3
         Top             =   1980
         Width           =   1140
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
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
         Left            =   750
         TabIndex        =   2
         Top             =   1980
         Width           =   1140
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1350
         TabIndex        =   0
         Top             =   1035
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
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
         Left            =   210
         TabIndex        =   7
         Top             =   1500
         Width           =   1080
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
         Index           =   0
         Left            =   165
         TabIndex        =   6
         Top             =   1050
         Width           =   1080
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   60
         Picture         =   "frm_Login.frx":0442
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1260
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Verification"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1530
         TabIndex        =   5
         Top             =   270
         UseMnemonic     =   0   'False
         Width           =   2115
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         X1              =   150
         X2              =   3720
         Y1              =   810
         Y2              =   810
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Index           =   1
         X1              =   210
         X2              =   3690
         Y1              =   780
         Y2              =   780
      End
   End
End
Attribute VB_Name = "frm_Login"
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo ET

  Dim db As Database
  Dim blnLogOnSuccess As Boolean

  strUID = txtUserName
  strPWD = txtPassword
  Set db = Open_Database(strUID, strPWD)
  blnLogOnSuccess = True
  Unload frm_Login
  frm_Security.Visible = True

  Exit Sub
  
ET:
    Select Case Err.Number
    Case Else
      MsgBox Err.Description, vbOKOnly + vbCritical, "Error # " & Err.Number
    End Select
    blnLogOnSuccess = False
    Err = False
    Exit Sub

End Sub
