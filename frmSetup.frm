VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Server"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   2910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox cmbServer 
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Server:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' just config

Option Base 0
Option Explicit
Option Compare Text

Private Sub Form_Load()
    cmbServer.Text = G("Server")
    With cmbServer
        .AddItem "U.S. East"
        .AddItem "U.S. West"
        .AddItem "Europe"
        .AddItem "Asia"
    End With
    If cmbServer.Text = vbNullString Then cmbServer.Text = cmbServer.List(0)
End Sub

Private Sub cmdOK_Click()
    If cmbServer.Text = vbNullString Then
        MsgBox "Please select a server.", vbExclamation
        Exit Sub
    End If
    W "Server", cmbServer.Text
    Me.Hide
    frmMain.Form_Load
End Sub

Private Sub cmdCancel_Click()
    frmMain.Form_Load
End Sub
