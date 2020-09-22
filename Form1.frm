VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process RTF"
      Height          =   555
      Left            =   225
      TabIndex        =   2
      Top             =   6705
      Width           =   2550
   End
   Begin VB.TextBox html 
      Height          =   3645
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   2535
      Width           =   8805
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   4154
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0006
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdProcess_Click()
'-----------------------------------------------------------------------------------
' call the RTF to HTML function
'-----------------------------------------------------------------------------------

    html.Text = RTFtoHTML(rtf.TextRTF)
End Sub

Private Sub Form_Load()

    rtf.LoadFile (App.Path & "\" & "sample.rtf")
End Sub
