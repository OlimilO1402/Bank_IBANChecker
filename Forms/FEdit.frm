VERSION 5.00
Begin VB.Form FEdit 
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   120
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   5880
      TabIndex        =   3
      Top             =   840
      Width           =   5895
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label LblIBAN 
      AutoSize        =   -1  'True
      Caption         =   "IBAN:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   450
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6000
      TabIndex        =   1
      Top             =   480
      Width           =   525
   End
End
Attribute VB_Name = "FEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Ret   As VbMsgBoxResult
Private m_IBANs As List

Public Function ShowDialog(OwnerForm As Form, IBANs As List) As VbMsgBoxResult
    
    Me.Show vbModal, OwnerForm
    
    
End Function

Private Sub BtnOK_Click()
    m_Ret = vbOK
    Unload Me
End Sub
Private Sub BtnCancel_Click()
    m_Ret = vbCancel
    Unload Me
End Sub

