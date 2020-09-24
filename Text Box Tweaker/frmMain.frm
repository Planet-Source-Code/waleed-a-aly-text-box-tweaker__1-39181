VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Text Boxes"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   Icon            =   "frmMain.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Tag             =   "By: Waleed A. Aly"
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   13
      Left            =   1980
      TabIndex        =   26
      Top             =   4860
      Width           =   2595
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   12
      Left            =   1980
      TabIndex        =   24
      Top             =   4500
      Width           =   2595
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   11
      Left            =   1980
      TabIndex        =   22
      Top             =   4140
      Width           =   2595
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   10
      Left            =   1980
      TabIndex        =   20
      Top             =   3780
      Width           =   2595
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   9
      Left            =   1980
      TabIndex        =   18
      Top             =   3420
      Width           =   2595
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   8
      Left            =   1980
      TabIndex        =   16
      Top             =   3060
      Width           =   2595
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   7
      Left            =   1980
      TabIndex        =   14
      Top             =   2700
      Width           =   2595
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   1980
      TabIndex        =   12
      Top             =   2340
      Width           =   2595
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   1980
      TabIndex        =   10
      Top             =   1980
      Width           =   2595
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   1980
      TabIndex        =   8
      Top             =   1620
      Width           =   2595
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   1980
      TabIndex        =   6
      Top             =   1260
      Width           =   2595
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   1980
      TabIndex        =   4
      Top             =   900
      Width           =   2595
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1980
      TabIndex        =   2
      Top             =   540
      Width           =   2595
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   1980
      TabIndex        =   0
      Top             =   180
      Width           =   2595
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "PhoneNumber"
      Height          =   195
      Index           =   13
      Left            =   240
      TabIndex        =   27
      Top             =   4920
      Width           =   1020
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "CashAllowNegative"
      Height          =   195
      Index           =   12
      Left            =   240
      TabIndex        =   25
      Top             =   4560
      Width           =   1380
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "CashPositive"
      Height          =   195
      Index           =   11
      Left            =   240
      TabIndex        =   23
      Top             =   4200
      Width           =   915
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "DecimalAllowNegative"
      Height          =   195
      Index           =   10
      Left            =   240
      TabIndex        =   21
      Top             =   3840
      Width           =   1590
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "DecimalPositive"
      Height          =   195
      Index           =   9
      Left            =   240
      TabIndex        =   19
      Top             =   3480
      Width           =   1125
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "IntegerAllowNegative"
      Height          =   195
      Index           =   8
      Left            =   240
      TabIndex        =   17
      Top             =   3120
      Width           =   1515
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "IntegerPositive"
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   15
      Top             =   2760
      Width           =   1050
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "AlphaNumericAllSmall"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   13
      Top             =   2400
      Width           =   1530
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "AlphaNumericAllCaps"
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   1515
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "AlphaNumeric"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "AllLettersAllSmall"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1185
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "AllLettersAllCaps"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1170
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "AllLetters"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   645
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Normal"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0
            Tweak Text(Index), KeyAscii, Normal
        Case 1
            Tweak Text(Index), KeyAscii, AllLetters
        Case 2
            Tweak Text(Index), KeyAscii, AllLettersAllCaps
        Case 3
            Tweak Text(Index), KeyAscii, AllLettersAllSmall
        Case 4
            Tweak Text(Index), KeyAscii, AlphaNumeric
        Case 5
            Tweak Text(Index), KeyAscii, AlphaNumericAllCaps
        Case 6
            Tweak Text(Index), KeyAscii, AlphaNumericAllSmall
        Case 7
            Tweak Text(Index), KeyAscii, IntegerPositive
        Case 8
            Tweak Text(Index), KeyAscii, IntegerAllowNegative
        Case 9
            Tweak Text(Index), KeyAscii, DecimalPositive
        Case 10
            Tweak Text(Index), KeyAscii, DecimalAllowNegative
        Case 11
            Tweak Text(Index), KeyAscii, CashPositive
        Case 12
            Tweak Text(Index), KeyAscii, CashAllowNegative
        Case 13
            Tweak Text(Index), KeyAscii, PhoneNumber
    End Select
End Sub
