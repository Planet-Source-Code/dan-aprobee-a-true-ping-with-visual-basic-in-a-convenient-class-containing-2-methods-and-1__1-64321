VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3540
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtConvert 
      Height          =   330
      Left            =   1350
      TabIndex        =   3
      Text            =   "http://yahoo.com"
      ToolTipText     =   $"Form1.frx":0000
      Top             =   765
      Width           =   1950
   End
   Begin VB.CommandButton cmdFormat 
      Caption         =   "&Convert www address to IP "
      Height          =   465
      Left            =   90
      TabIndex        =   2
      Top             =   720
      Width           =   1230
   End
   Begin VB.TextBox txtPing 
      Height          =   330
      Left            =   1350
      TabIndex        =   1
      ToolTipText     =   $"Form1.frx":0089
      Top             =   225
      Width           =   1950
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ping"
      Height          =   465
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   1230
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cping As cping
Attribute cping.VB_VarHelpID = -1


Private Sub cmdFormat_Click()
  'convert a string web address to an ip address
  txtPing = cping.funcGetIPFromHostName(txtConvert)
End Sub

Private Sub Command1_Click()
  cping.Ping txtPing
End Sub

Private Sub cping_PingReturn(ipAddress As String, successStatus As String, roundTripMilliseconds As Long)
  'display ping stats in debug
  Debug.Print ipAddress & vbTab & _
              successStatus & vbTab & _
              roundTripMilliseconds
End Sub

Private Sub Form_Load()
  Set cping = New cping
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set cping = Nothing
End Sub
