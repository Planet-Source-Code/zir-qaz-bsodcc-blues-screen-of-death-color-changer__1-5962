VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form cd1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Blue Screen Of Death Color Changer"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd2 
      Left            =   1320
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Apply Colors"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Text"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Background"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Sample Of BSODC"
      Height          =   735
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "cd1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function ReadINI(AppName$, KeyName$, filename$) As String
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   ReadINI = Left(RetStr, GetPrivateProfileString(AppName$, _
   ByVal KeyName$, "", RetStr, Len(RetStr), filename$))
End Function

Sub WriteINI(AppName$, KeyName$, Entry$, filename$)
   Dim X As Integer
   X = WritePrivateProfileString(AppName$, KeyName$, Entry$, filename$)
End Sub

Public Function GetWinPath()
Dim strFolder As String
Dim lngResult As Long
strFolder = String(MAX_PATH, 0)
lngResult = GetWindowsDirectory(strFolder, MAX_PATH)
If lngResult <> 0 Then
    GetWinPath = Left(strFolder, InStr(strFolder, Chr(0)) - 1)
Else
    GetWinPath = ""
End If
End Function

Public Function GetWinSysPath()
Dim strFolder As String
Dim lngResult As Long
strFolder = String(MAX_PATH, 0)
lngResult = GetSystemDirectory(strFolder, MAX_PATH)
If lngResult <> 0 Then
    GetWinSysPath = Left(strFolder, InStr(strFolder, _
    Chr(0)) - 1)
Else
    GetSysWinPath = ""
End If
End Function
Private Sub Command1_Click()
cd1.ShowColor
Label1.BackColor = cd1.Color


End Sub

Private Sub Command2_Click()
cd2.ShowColor
Label1.ForeColor = cd2.Color
End Sub

Private Sub Command3_Click()
WriteINI "386Enh", "MessageBackcolor", cd1.Color, GetWinPath + "\system.ini"
WriteINI "386Enh", "MessageTextcolor", cd2.Color, GetWinPath + "\system.ini"
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()
xb = ReadINI("386Enh", "MessageBackcolor", GetWinPath + "\system.ini")
Label1.BackColor = xb
xb2 = ReadINI("386Enh", "MessageTextcolor", GetWinPath + "\system.ini")
Label1.ForeColor = xb2
End Sub



