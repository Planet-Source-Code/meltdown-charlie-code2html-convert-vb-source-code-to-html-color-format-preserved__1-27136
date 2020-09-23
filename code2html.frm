VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form fCode2html 
   Caption         =   "Form1"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   6285
      Left            =   105
      TabIndex        =   1
      Top             =   60
      Width           =   9285
      ExtentX         =   16378
      ExtentY         =   11086
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Paste from Clipboard"
      Height          =   720
      Left            =   8115
      TabIndex        =   0
      Top             =   6465
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Instructions -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   4
      Left            =   165
      TabIndex        =   6
      Top             =   6525
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "4 : Wait for a moment - a ""viola"" your text is displayed in HTML"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   525
      TabIndex        =   5
      Top             =   7995
      Width           =   7455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "3 : Press the ""Past from Clipboard"" button on this form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   525
      TabIndex        =   4
      Top             =   7635
      Width           =   6540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "2 : Copy it to the clipboard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   525
      TabIndex        =   3
      Top             =   7275
      Width           =   3165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "1 : Select the text you want to convert into HTML"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   525
      TabIndex        =   2
      Top             =   6930
      Width           =   5865
   End
End
Attribute VB_Name = "fCode2html"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command4_Click()
    Dim s As String
    Dim i As Integer
    Dim lns() As String
    Dim f As Long
    
    s = Clipboard.GetText
    s = ProcessBlock(s)
    f = FreeFile
    Open App.Path & "\smp.html" For Output As f
    Print #f, s
    Close f
    Web.Navigate "file://" & App.Path & "\smp.html"
End Sub

