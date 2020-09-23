VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDemoQSort 
   Caption         =   "QuickSort tricks"
   ClientHeight    =   7770
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleMode       =   0  'User
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkCaseInSensitive 
      Caption         =   "Case InSensitive"
      Height          =   255
      Left            =   6720
      TabIndex        =   16
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CheckBox chkPunctuation 
      Caption         =   "Expand Punctation"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   15
      ToolTipText     =   "seperate punctuation from text"
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CheckBox chkPunctuation 
      Caption         =   "Ignore Punctuation"
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   14
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CheckBox chkIgnoreFormat 
      Caption         =   "Ignore Format and dbl spaces"
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Frame fraFind 
      Caption         =   "InQSortArray & QSortArrayPos "
      Height          =   1215
      Left            =   360
      TabIndex        =   7
      Top             =   6480
      Width           =   8295
      Begin VB.PictureBox picCFXPBugFixForm1 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   100
         ScaleHeight     =   885
         ScaleWidth      =   7980
         TabIndex        =   8
         Top             =   276
         Width           =   7980
         Begin VB.CommandButton cmdFind 
            Caption         =   "Find Unique"
            Height          =   315
            Index           =   1
            Left            =   2060
            TabIndex        =   11
            Top             =   297
            Width           =   1575
         End
         Begin VB.TextBox txtFind 
            Height          =   375
            Left            =   20
            TabIndex        =   10
            Top             =   -18
            Width           =   1695
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "Find"
            Height          =   315
            Index           =   0
            Left            =   2060
            TabIndex        =   9
            Top             =   -18
            Width           =   1575
         End
         Begin VB.Label lblFind 
            BorderStyle     =   1  'Fixed Single
            Height          =   855
            Left            =   3960
            TabIndex        =   12
            Top             =   -15
            Width           =   2775
         End
      End
   End
   Begin VB.CheckBox chkAsendingSort 
      Caption         =   "Asending Sort"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5880
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuickSort 
      Caption         =   "QSUniqueFrequency"
      Height          =   255
      Index           =   3
      Left            =   6600
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cdlDemo 
      Left            =   120
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdQuickSort 
      Caption         =   "QuickSortUniqueCount"
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   2055
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   3625
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      FileName        =   "C:\downloads2004\04-08\10\EvanSort--_745_word_sorted_in_well_under_2_seconds_PSC\Module1.bas"
      TextRTF         =   $"frmDemoQSort.frx":0000
   End
   Begin VB.ListBox lstDisplay 
      Columns         =   5
      Height          =   2985
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   8655
   End
   Begin VB.CommandButton cmdQuickSort 
      Caption         =   "QuickSortUnique"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuickSort 
      Caption         =   "QuickSort"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblReport 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   5520
      Width           =   8055
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufileOpt 
         Caption         =   "&Load"
         Index           =   0
      End
      Begin VB.Menu mnufileOpt 
         Caption         =   "E&xit"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmDemoQSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bMsgBox     As Boolean

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub chkPunctuation_Click(Index As Integer)


  Select Case Index
   Case 0
    If chkPunctuation(0).Value = vbChecked Then
      chkPunctuation(1).Value = vbUnchecked
    End If
   Case 1
    If chkPunctuation(1).Value = vbChecked Then
      chkPunctuation(0).Value = vbUnchecked
    End If
  End Select

End Sub

Private Sub cmdFind_Click(Index As Integer)

  
  Dim arr           As Variant

  Dim lng_startTime As Single
  Dim workString    As String
  Dim lng_sortTime  As Single
  'Dim TestMode As Long
  workString = GetWorkString(rtb1, chkCaseInSensitive.Value = vbChecked, chkIgnoreFormat.Value = vbChecked, chkPunctuation(0).Value = vbChecked, chkPunctuation(1).Value = vbChecked)
  lng_startTime = GetTickCount()
  Select Case Index
   Case 0
    arr = QuickSortArray(Split(workString), chkAsendingSort.Value = vbChecked)
   Case 1
    arr = QuickSortUniqueArray(Split(workString), chkAsendingSort.Value = vbChecked)
   Case 2
    arr = QuickSortArray(Split(workString), chkAsendingSort.Value = vbChecked)
  End Select
  lblFind = "Exists: " & InQSortArray(arr, txtFind.Text) & vbNewLine & _
            "Position: " & QSortArrayPos(arr, txtFind.Text) & vbNewLine & _
            "Occurs: " & QArrayCount(QuickSortArray(Split(workString), chkAsendingSort.Value = vbChecked), txtFind.Text)
  lng_sortTime = GetTickCount()
  lblFind = lblFind & vbNewLine & _
   "Time Taken: " & ((lng_sortTime - lng_startTime) / 1000) & " seconds "

End Sub

Private Sub cmdQuickSort_Click(Index As Integer)

  
  Dim workString    As String

  Dim lng_startTime As Single
  Dim lng_sortTime  As Single
  Dim arr           As Variant
  Dim I             As Long
  lstDisplay.Clear
  workString = GetWorkString(rtb1, chkCaseInSensitive.Value = vbChecked, chkIgnoreFormat.Value = vbChecked, chkPunctuation(0).Value = vbChecked, chkPunctuation(1).Value = vbChecked)
  lng_startTime = GetTickCount()
  Select Case Index
   Case 0
    'make an array out of the RTB text then quicksort it
    arr = QuickSortArray(Split(workString), chkAsendingSort.Value = vbChecked)
   Case 1
    arr = QuickSortUniqueArray(Split(workString), chkAsendingSort.Value = vbChecked)
   Case 2
    arr = QuickSortUniqueCountArray(Split(workString), chkAsendingSort.Value = vbChecked)
   Case 3
    arr = QuickSortUniqueFrequencyArray(Split(workString), chkAsendingSort.Value = vbChecked)
  End Select
  lng_sortTime = GetTickCount()
  SendMessage lstDisplay.hWnd, WM_SETREDRAW, False, 0
  For I = LBound(arr) To UBound(arr)
    lstDisplay.AddItem arr(I)
  Next I
  SendMessage lstDisplay.hWnd, WM_SETREDRAW, True, 0
  lblReport = UBound(arr) & " words sorted in  " & ((lng_sortTime - lng_startTime) / 1000) & " seconds"
  If Not bMsgBox Then
    MsgBox "Note 1 Generating array time only, does not include preformatting or filling the list" & vbNewLine & _
       "Note 2 Because of the number of double spaces in the source 'QuickSort' may not appear" & vbNewLine & _
       " to display anything as the visible list will be all blanks, scroll across list to see visible words."
    bMsgBox = True
  End If

End Sub

Private Sub mnufileOpt_Click(Index As Integer)


  Select Case Index
   Case 0
    With cdlDemo
      .ShowOpen
      If LenB(.FileName) Then
        rtb1.LoadFile .FileName
      End If
    End With
   Case 1
    Unload Me
  End Select

End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)


  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdFind_Click 0
  End If

End Sub

':)Code Fixer V2.4.6 (13/08/2004 3:27:39 PM) 7 + 144 = 151 Lines Thanks Ulli for inspiration and lots of code.

