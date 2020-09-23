VERSION 5.00
Begin VB.Form frmGpaper1 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hawk's Graph Paper"
   ClientHeight    =   3375
   ClientLeft      =   2985
   ClientTop       =   2925
   ClientWidth     =   5160
   Icon            =   "gpaper2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5160
   Begin VB.Frame fraPrint 
      BackColor       =   &H0000FFFF&
      Caption         =   "Printer Stuff"
      Height          =   975
      Left            =   2640
      TabIndex        =   20
      Top             =   1200
      Width           =   2415
      Begin VB.ListBox lstCopies 
         Height          =   255
         Left            =   1200
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Copies"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraScale 
      BackColor       =   &H0000FFFF&
      Caption         =   "Scale"
      Height          =   975
      Left            =   2640
      TabIndex        =   17
      Top             =   120
      Width           =   2415
      Begin VB.OptionButton optCentimeter 
         BackColor       =   &H0000FFFF&
         Caption         =   "Centimeters"
         Height          =   255
         Left            =   360
         MaskColor       =   &H0080FFFF&
         TabIndex        =   19
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton optInch 
         BackColor       =   &H0000FFFF&
         Caption         =   "Inches"
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Frame fraLines 
      BackColor       =   &H0000FFFF&
      Caption         =   "Look of Lines"
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   2415
      Begin VB.ListBox lstColor 
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.ListBox lstThick 
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblColor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblThickness 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Thickness"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Frame fraThickLines 
      BackColor       =   &H0000FFFF&
      Caption         =   "Lines per Wide Line"
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   2415
      Begin VB.ListBox lstWideV 
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.ListBox lstWideH 
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblThickV 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vertical"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblThivkH 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Horizonal"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraNumLines 
      BackColor       =   &H0000FFFF&
      Caption         =   "Number of Lines"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.ListBox lstLinesV 
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.ListBox lstLinesH 
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblLinesV 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vertical"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblLinesH 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Horizonal"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmGpaper1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ScaleType$
Private Sub LoadListboxes()
    ' -- Load Values into listboxes and set default values
    Dim Counter%
    For Counter = 0 To 50
        lstLinesH.AddItem Counter
        lstLinesV.AddItem Counter
        lstWideH.AddItem Counter
        lstWideV.AddItem Counter
        lstThick.AddItem Counter
        lstCopies.AddItem Counter
    Next
    lstColor.AddItem "Black"
    'vbBlack &H0 Black
    lstColor.AddItem "Red"
    'vbRed   &HFF    Red
    lstColor.AddItem "Green"
    'vbGreen &HFF00  Green
    lstColor.AddItem "Yellow"
    'vbYellow    &HFFFF  Yellow
    lstColor.AddItem "Blue"
    'vbBlue  &HFF0000    Blue
    lstColor.AddItem "Magenta"
    'vbMagenta   &HFF00FF    Magenta
    lstColor.AddItem "Cyan"
    'vbCyan  &HFFFF00    Cyan
    'vbWhite &HFFFFFF    White
    lstLinesH.ListIndex = 10
    lstLinesV.ListIndex = 10
    lstWideH.ListIndex = 10
    lstWideV.ListIndex = 10
    lstThick.ListIndex = 1
    lstColor.ListIndex = 0
    lstCopies.ListIndex = 1
    optInch_Click   'select inch scale

End Sub
Private Sub cmdExit_Click()
    ' --- Exit Program
    End
End Sub

Private Sub cmdPrint_Click()
    ' --- Build and print Graph Paper
    On Error GoTo PrintFailed
    cmdPrint.Enabled = False
    Dim ColorToPrint As Long
    Dim Thickness%, WideThickness%, LineH%, LineV%
    Dim WideH%, WideV%
    Dim PrnMaxH%, PrnMaxW%, PrnMinH%, PrnMinW%
    Dim Counter%, WideCount%, CopyCounter%, NumOfCopies%
    If Printer.ColorMode = 1 Then 'Black and white only
        ColorToPrint = vbBlack
    Else
        Select Case lstColor.ListIndex
            Case 0
                ColorToPrint = vbBlack
                'vbBlack &H0 Black
            Case 1
                ColorToPrint = vbRed
                'vbRed   &HFF    Red
            Case 2
                ColorToPrint = vbGreen
                'vbGreen &HFF00  Green
            Case 3
                ColorToPrint = vbYellow
                'vbYellow    &HFFFF  Yellow
            Case 4
                ColorToPrint = vbBlue
                'vbBlue  &HFF0000    Blue
            Case 5
                ColorToPrint = vbMagenta
                'vbMagenta   &HFF00FF    Magenta
            Case 6
                ColorToPrint = vbCyan
                'vbCyan  &HFFFF00    Cyan
            Case Else
                ColorToPrint = vbBlack
        End Select
    End If ' - finished with color
    ' --- set Line Thickness
    Thickness = lstThick.ListIndex
    If Thickness = 0 Then Thickness = 1
    ' -- Set the Wide Line Thickness
    WideThickness = (Thickness * 1.5) + 2
    ' - max size of drawing area of printer
    ' I don't know whether to use scaleheight or height for this
    PrnMaxH = Printer.ScaleHeight
    PrnMaxW = Printer.ScaleWidth
    PrnMinH = 0 + WideThickness
    PrnMinW = 0 + WideThickness
    'LineH%, LineV%, WideH%, WideV%
    Select Case ScaleType
        Case "Inch"
            LineH = 1440 / lstLinesH.ListIndex
            LineV = 1440 / lstLinesV.ListIndex
        Case "Centimeter"
            LineH = 567 / lstLinesH.ListIndex
            LineV = 567 / lstLinesV.ListIndex
        Case Else
            LineH = 1440 / lstLinesH.ListIndex
            LineV = 1440 / lstLinesV.ListIndex
    End Select

    WideH = lstWideH.ListIndex
    WideV = lstWideV.ListIndex
    NumOfCopies = lstCopies.ListIndex
    If NumOfCopies = 0 Then NumOfCopies = 1
    For CopyCounter = 1 To NumOfCopies
            ' -- Horizonal Lines
            WideCount = 0
            For Counter = PrnMinH To PrnMaxH Step LineH
                If WideCount = 0 Then
                    WideCount = WideH - 1
                    Printer.DrawWidth = WideThickness
                Else
                    WideCount = WideCount - 1
                    Printer.DrawWidth = Thickness
                End If
                Printer.Line (PrnMinW, Counter)-(PrnMaxW, Counter), ColorToPrint
            Next
            ' -- Vertical Lines
            WideCount = 0
            For Counter = PrnMinW To PrnMaxW Step LineH
                If WideCount = 0 Then
                    WideCount = WideV - 1
                    Printer.DrawWidth = WideThickness
                Else
                    WideCount = WideCount - 1
                    Printer.DrawWidth = Thickness
                End If
                Printer.Line (Counter, PrnMinH)-(Counter, PrnMaxH), ColorToPrint
            Next
            'Printer.NewPage
            Printer.EndDoc  'finished printing
    Next 'copies
    'Printer.EndDoc  'finished printing
    cmdPrint.Enabled = True
    Exit Sub
PrintFailed:
    MsgBox "There was a problem printing!"
    cmdPrint.Enabled = True
    Exit Sub
End Sub

Private Sub Form_Click()
    ' --- Give quick program info if you click on a
    '     blank section of the form.
    MsgBox "Hawk's Graph Paper Printer"
End Sub

Private Sub Form_Load()
    ' -- Build listbox values
    LoadListboxes
End Sub

Private Sub mnuExit_Click()
    ' --- Menu item, exit program
    cmdExit_Click
End Sub

Private Sub mnuPrint_Click()
    ' --- Menu item, print graph paper
    cmdPrint_Click
End Sub

Private Sub optCentimeter_Click()
    ' --- Select number of lines per centimeter
    fraNumLines.Caption = "Number of lines per Centimeter"
    ScaleType = "Centimeter"
End Sub

Private Sub optInch_Click()
    ' --- Select number of lines per inch
    fraNumLines.Caption = "Number of lines -- Inch"
    ScaleType = "Inch"
End Sub
