VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "IBAN-Checker"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8895
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton btnBBbic 
      Caption         =   "^"
      Height          =   375
      Left            =   15480
      TabIndex        =   57
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton btnBBbank 
      Caption         =   "^"
      Height          =   375
      Left            =   15480
      TabIndex        =   56
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton btnBBort 
      Caption         =   "^"
      Height          =   375
      Left            =   15480
      TabIndex        =   55
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton btnBBplz 
      Caption         =   "^"
      Height          =   375
      Left            =   15480
      TabIndex        =   54
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton btnBBblz 
      Caption         =   "^"
      Height          =   375
      Left            =   15480
      TabIndex        =   53
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox TxBBbic 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   49
      Top             =   2520
      Width           =   5295
   End
   Begin VB.TextBox TxBBbank 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   46
      Top             =   2040
      Width           =   5295
   End
   Begin VB.TextBox TxBBort 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   45
      Top             =   1560
      Width           =   5295
   End
   Begin VB.TextBox TxBBplz 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   44
      Top             =   1080
      Width           =   5295
   End
   Begin VB.ComboBox CbBlzBic 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Form1.frx":179A
      Left            =   10200
      List            =   "Form1.frx":179C
      TabIndex        =   43
      Top             =   120
      Width           =   5655
   End
   Begin VB.TextBox TxBBblz 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   42
      Top             =   600
      Width           =   5295
   End
   Begin VB.PictureBox PnlKtrlZif2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   8535
      TabIndex        =   34
      Top             =   5760
      Width           =   8535
      Begin VB.TextBox TxKtrlZif2 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   35
         Top             =   0
         Width           =   5295
      End
      Begin VB.Label LbZKtrlZif2 
         BackStyle       =   0  'Transparent
         Caption         =   "max. 10 Ziffern"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   37
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label LbKtrlZif2 
         BackStyle       =   0  'Transparent
         Caption         =   "Kontrollziffer2"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.PictureBox PnlSFnkt 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   8535
      TabIndex        =   12
      Top             =   5280
      Width           =   8535
      Begin VB.TextBox TxSFnkt 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   9
         Top             =   0
         Width           =   5295
      End
      Begin VB.Label LbSFnkt 
         BackStyle       =   0  'Transparent
         Caption         =   "sonst. Funkt."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label LbZSFnkt 
         BackStyle       =   0  'Transparent
         Caption         =   "max. 10 Ziffern"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   14
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox PnlFilNr 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   8535
      TabIndex        =   13
      Top             =   4800
      Width           =   8535
      Begin VB.TextBox TxFilNr 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   8
         Top             =   0
         Width           =   5295
      End
      Begin VB.Label LbFilNr 
         BackStyle       =   0  'Transparent
         Caption         =   "Filialnummer"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label LbZFilNr 
         BackStyle       =   0  'Transparent
         Caption         =   "max. 10 Ziffern"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox PnlRegC 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   8535
      TabIndex        =   18
      Top             =   4320
      Width           =   8535
      Begin VB.TextBox TxRegC 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   0
         Width           =   5295
      End
      Begin VB.Label LbRegC 
         BackStyle       =   0  'Transparent
         Caption         =   "Regionalcode"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label LbZRegC 
         BackStyle       =   0  'Transparent
         Caption         =   "max. 10 Ziffern"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   23
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox PnlKtrlZif 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   8535
      TabIndex        =   19
      Top             =   3840
      Width           =   8535
      Begin VB.TextBox TxKtrlZif 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         Top             =   0
         Width           =   5295
      End
      Begin VB.Label LbKtrlZif 
         BackStyle       =   0  'Transparent
         Caption         =   "Kontrollziffer"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label LbZKtrlZif 
         BackStyle       =   0  'Transparent
         Caption         =   "max. 10 Ziffern"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   25
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox PnlKtoNr 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   8535
      TabIndex        =   20
      Top             =   3360
      Width           =   8535
      Begin VB.TextBox TxKtoNr 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         Top             =   0
         Width           =   5295
      End
      Begin VB.Label LbKtoNr 
         BackStyle       =   0  'Transparent
         Caption         =   "Kontonummer"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label LbZKtoNr 
         BackStyle       =   0  'Transparent
         Caption         =   "max. 10 Ziffern"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   27
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox PnlKTyp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   8535
      TabIndex        =   21
      Top             =   2880
      Width           =   8535
      Begin VB.TextBox TxKTyp 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   4
         Top             =   0
         Width           =   5295
      End
      Begin VB.Label LbKTyp 
         BackStyle       =   0  'Transparent
         Caption         =   "Kontotyp"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label LbZKTyp 
         BackStyle       =   0  'Transparent
         Caption         =   "max. 10 Ziffern"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   29
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox PnlBLZ 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   8775
      TabIndex        =   22
      Top             =   2400
      Width           =   8775
      Begin VB.CommandButton BtnOpenBlzBic 
         Caption         =   ">"
         Height          =   375
         Left            =   8400
         TabIndex        =   58
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox TxBLZ 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   0
         Width           =   5295
      End
      Begin VB.Label LbBLZ 
         BackStyle       =   0  'Transparent
         Caption         =   "Bankleitzahl"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label LbZBLZ 
         BackStyle       =   0  'Transparent
         Caption         =   "max. 10 Ziffern"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   31
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.CommandButton btnCheckIBAN 
      Caption         =   "Check IBAN v"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   39
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox TxIBAN 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   38
      Top             =   120
      Width           =   6975
   End
   Begin VB.TextBox TxBBANInfoW 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   33
      Top             =   1920
      Width           =   6975
   End
   Begin VB.CommandButton btnCalcIBAN 
      Caption         =   "Calc IBAN ^"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   600
      Width           =   2175
   End
   Begin VB.CheckBox CkGroup4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4er Gruppen"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.ComboBox CmbLC 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Form1.frx":179E
      Left            =   1560
      List            =   "Form1.frx":17A0
      TabIndex        =   1
      Top             =   1080
      Width           =   6975
   End
   Begin VB.TextBox TxBBANInfoR 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1560
      Width           =   6975
   End
   Begin VB.Label LbBlzBics 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   59
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "BIC:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   52
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   51
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Ort:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   50
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "PLZ:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   48
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "BLZ:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   47
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Land:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   41
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "BBAN-Format:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   40
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "IBAN:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'79.228.162.514.264.337.593.543.950.335
Private m_IBANInfo  As IBANInfo
Private m_iis       As IBANInfos
Private m_BBANInfoR As String
Private m_BlzBics   As BlzBics
Private m_col       As Collection 'Of BlzBic

Private Sub Form_Load()
    Set m_iis = New IBANInfos
    m_iis.ReadFromFile App.Path & "\Data\ibancodes.txt"
    Set m_BlzBics = MNew.BlzBics(App.Path & "\Data\blzBIC3_2015.txt")
    m_iis.FillComboBox CmbLC
    CmbLC.ListIndex = 18
    'Me.ScaleWidth = 8895
End Sub

Private Sub btnBBblz_Click()
    Set m_col = m_BlzBics.BLZcol(TxBBblz)
    FillCbBlzBic
End Sub
Private Sub btnBBplz_Click()
    Set m_col = m_BlzBics.PLZcol(TxBBplz)
    FillCbBlzBic
End Sub
Private Sub btnBBort_Click()
    Set m_col = m_BlzBics.ORTcol(TxBBort)
    FillCbBlzBic
End Sub
Private Sub btnBBbank_Click()
    Set m_col = m_BlzBics.BANKcol(TxBBbank)
    FillCbBlzBic
End Sub
Private Sub btnBBbic_Click()
    Set m_col = m_BlzBics.BICcol(TxBBbic)
    FillCbBlzBic
End Sub
Sub FillCbBlzBic()
    If m_col Is Nothing Then Exit Sub
    Dim v, bb As BlzBic
    CbBlzBic.Clear
    For Each v In m_col
        Set bb = v
        If Not bb Is Nothing Then
            CbBlzBic.AddItem bb.ToStr
        End If
    Next
    If CbBlzBic.ListCount > 0 Then CbBlzBic.ListIndex = 0
    LbBlzBics = m_col.Count
End Sub

Private Sub BtnOpenBlzBic_Click()
    If BtnOpenBlzBic.Caption = ">" Then
        Me.Width = Me.Width - Me.ScaleWidth + 15975
        BtnOpenBlzBic.Caption = "<"
        If Len(TxBLZ.Text) > 0 Then
            Set m_col = m_BlzBics.BLZcol(TxBLZ.Text)
            Dim bb As BlzBic
            For Each bb In m_col
                CbBlzBic.AddItem bb.BLZ
            Next
            CbBlzBic.ListIndex = 0
        End If
    Else
        Me.Width = Me.Width - Me.ScaleWidth + 8895
        BtnOpenBlzBic.Caption = ">"
        If Len(TxBBblz.Text) > 0 Then
            TxBLZ.Text = TxBBblz.Text
        End If
    End If
End Sub

Private Sub CbBlzBic_Click()
    Dim bb As BlzBic: Set bb = m_col.Item(CbBlzBic.ListIndex + 1)
    With bb
        TxBBblz = .BLZ
        TxBBplz = .PLZ
        TxBBort = .Ort
        TxBBbank = .BanknameLok
        TxBBbic = .BIC
    End With
End Sub

Private Sub Form_Resize()
    Resize
End Sub
Private Sub Resize()
    Dim brdr: brdr = 8 * Screen.TwipsPerPixelX
    Dim l As Single: l = 0
    Dim T As Single: T = 0
    Dim H As Single: H = 495
    l = brdr: T = 2400 '1200
    
    Dim sArr: sArr = Split(m_BBANInfoR, "; ")
    Dim i As Long, k1 As Boolean
    For i = 0 To UBound(sArr)
        If Len(sArr(i)) > 0 Then
            Dim elms: elms = Split(sArr(i), ": ")
            Select Case elms(0)
            Case "b":  PnlBLZ.Move l, T: T = T + H
            Case "d":  PnlKTyp.Move l, T: T = T + H
            Case "k":  PnlKtoNr.Move l, T: T = T + H
            Case "K":
                If Not k1 Then
                    PnlKtrlZif.Move l, T: T = T + H
                    k1 = True
                Else
                    PnlKtrlZif2.Move l, T: T = T + H
                End If
            Case "r":  PnlRegC.Move l, T: T = T + H
            Case "s":  PnlFilNr.Move l, T: T = T + H
            Case "X":  PnlSFnkt.Move l, T: T = T + H
            End Select
        End If
    Next
    'ab hier wird resize nochmal ausgeführt falls Height-neu anders als Height-alt
    Me.Height = (Me.Height - Me.ScaleHeight) + T + brdr '+ H
'    If PnlBLZ.Visible Then PnlBLZ.Move l, T: T = T + H
'    If PnlKTyp.Visible Then PnlKTyp.Move l, T: T = T + H
'    If PnlKtoNr.Visible Then PnlKtoNr.Move l, T: T = T + H
'    If PnlKtrlZif.Visible Then PnlKtrlZif.Move l, T: T = T + H
'    If PnlRegC.Visible Then PnlRegC.Move l, T: T = T + H
'    If PnlFilNr.Visible Then PnlFilNr.Move l, T: T = T + H
'    If PnlSFnkt.Visible Then PnlSFnkt.Move l, T: T = T + H
End Sub

Private Sub CkGroup4_Click()
    If Len(TxIBAN.Text) Then
        If (CkGroup4.Value = vbChecked) Then
            TxIBAN.Text = Trim(Group4(TxIBAN.Text))
        Else
            TxIBAN.Text = Trim(StringClean(TxIBAN.Text))
        End If
    End If
End Sub

Private Sub CmbLC_Click()
    Dim li As Integer: li = CmbLC.ListIndex
    Set m_IBANInfo = m_iis.Item(li)
    m_BBANInfoR = m_IBANInfo.BBANInfo.ToStr(True)
    TxBBANInfoR.Text = m_BBANInfoR
    TxBBANInfoW.Text = m_IBANInfo.BBANInfo.ToStr
    PnlBLZ.Visible = False
    PnlKTyp.Visible = False
    PnlKtoNr.Visible = False
    PnlKtrlZif.Visible = False
    PnlKtrlZif2.Visible = False
    PnlRegC.Visible = False
    PnlFilNr.Visible = False
    PnlSFnkt.Visible = False
    Dim sArr() As String: sArr = Split(m_BBANInfoR, "; ")
    Dim i As Long
    Dim k1 As Boolean
    For i = 0 To UBound(sArr)
        If Len(sArr(i)) > 0 Then
            Dim elms: elms = Split(sArr(i), ": ")
            Select Case elms(0)
            Case "b":  Enable PnlBLZ, LbZBLZ, TxBLZ, elms(1)
            Case "d":  Enable PnlKTyp, LbZKTyp, TxKTyp, elms(1)
            Case "k":  Enable PnlKtoNr, LbZKtoNr, TxKtoNr, elms(1)
            Case "K"
                If Not k1 Then
                    Enable PnlKtrlZif, LbZKtrlZif, TxKtrlZif, elms(1)
                    k1 = True
                Else
                    Enable PnlKtrlZif2, LbZKtrlZif2, TxKtrlZif2, elms(1)
                End If
            Case "r":  Enable PnlRegC, LbZRegC, TxRegC, elms(1)
            Case "s":  Enable PnlFilNr, LbZFilNr, TxFilNr, elms(1)
            Case "X":  Enable PnlSFnkt, LbZSFnkt, TxSFnkt, elms(1)
            End Select
        End If
    Next
    Resize
End Sub
Private Sub Enable(pnl As PictureBox, LbZ As Label, Tb As TextBox, ByVal z As Long)
    pnl.Visible = True: LbZ.Caption = "max. " & z & " Ziffern": Tb.Tag = z
End Sub

Private Sub btnCheckIBAN_Click()
    Dim IBAN As IBAN: Set IBAN = MNew.IBAN(m_iis, TxIBAN.Text)
    Dim s As String: s = IBAN.IBANInfo.Key
    CmbLC.ListIndex = m_iis.Index(IBAN.IBANInfo.CountryID)
    s = s & vbCrLf
    Dim BBAN As BBAN: Set BBAN = IBAN.BBAN
    Dim i As Long: Dim bv As BBANValue
    'jetzt die Textboxen mit den Bestandteilen der IBAN befüllen
    For i = 0 To BBAN.CountParts - 1
        Set bv = BBAN.Prop(i)
        With bv
            s = s & .BBANPart.Name & " = " & bv.Value & vbCrLf
            Dim e As EBBANPart: e = .BBANPart.EBBANPart
            Select Case e
            Case Bankleitzahl:    TxBLZ = .Value       '"b" 'bank identifier
            Case Kontotyp:        TxKTyp = .Value      '"d" 'type of account
            Case Kontonummer:     TxKtoNr = .Value     '"k" 'bank account number
            Case Kontrollziffer:  TxKtrlZif = .Value   '"K" 'control code
            Case Regionalcode:    TxRegC = .Value      '"r" 'region code
            Case Filialnummer:    TxFilNr = .Value     '"s" 'branch identifier
            Case SonstFunktion:   TxSFnkt = .Value     '"X" 'other functions
            Case Kontrollziffer2: TxKtrlZif2 = .Value
            End Select
        End With
    Next
End Sub

Private Sub TxBLZ_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        'zum nächsten Feld springen
        'schauen ob alle gefüllt
    End If
End Sub
Private Sub btnCalcIBAN_Click()
    FetchIBAN
    Clipboard.SetText TxIBAN.Text
End Sub

Private Sub FetchIBAN()
    Dim u As Long
    ReDim sArr(0 To m_IBANInfo.BBANInfo.CountBBANParts - 1) As String
    Dim s As String
    If PnlBLZ.Visible Then
        If Not TryGetStr(LbBLZ, TxBLZ, s) Then Exit Sub
        ReDim sArr(u): sArr(u) = s: u = u + 1
    End If
    If PnlKTyp.Visible Then
        If Not TryGetStr(LbKTyp, TxKTyp, s) Then Exit Sub
        ReDim Preserve sArr(u): sArr(u) = s: u = u + 1
    End If
    If PnlKtoNr.Visible Then
        If Not TryGetStr(LbKtoNr, TxKtoNr, s) Then Exit Sub
        ReDim Preserve sArr(u): sArr(u) = s: u = u + 1
    End If
    If PnlKtrlZif.Visible Then
        If Not TryGetStr(LbKtrlZif, TxKtrlZif, s) Then Exit Sub
        ReDim Preserve sArr(u): sArr(u) = s: u = u + 1
    End If
    If PnlRegC.Visible Then
        If Not TryGetStr(LbRegC, TxRegC, s) Then Exit Sub
        ReDim Preserve sArr(u): sArr(u) = s: u = u + 1
    End If
    If PnlFilNr.Visible Then
        If Not TryGetStr(LbFilNr, TxFilNr, s) Then Exit Sub
        ReDim Preserve sArr(u): sArr(u) = s: u = u + 1
    End If
    If PnlSFnkt.Visible Then
        If Not TryGetStr(LbSFnkt, TxSFnkt, s) Then Exit Sub
        ReDim Preserve sArr(u): sArr(u) = s: u = u + 1
    End If
    Dim li As Integer: li = CmbLC.ListIndex
    Dim IC As IBANCreator: Set IC = MNew.IBANCreator(m_iis, m_iis.Item(li), sArr)
    TxIBAN.Text = Trim(IC.IBAN.ToStr)
    CkGroup4_Click
End Sub
Function TryGetStr(Lb As Label, Tb As TextBox, strout As String) As Boolean
    Dim s As String: s = StringClean(Tb.Text)
    If Len(s) = 0 Or Len(s) > CLng(Tb.Tag) Then
        MsgBox "Bitte geben Sie im Feld " & Lb & " einen gültigen Wert ein."
        Exit Function
    End If
    strout = s
    TryGetStr = True
End Function
