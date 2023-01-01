VERSION 5.00
Begin VB.Form FEdit 
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
   Begin VB.ListBox List1 
      Height          =   5325
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label LblIBAN 
      AutoSize        =   -1  'True
      Caption         =   "IBAN:"
      Height          =   195
      Left            =   6000
      TabIndex        =   2
      Top             =   960
      Width           =   420
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "FEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

