Attribute VB_Name = "basWordQSort"
Option Explicit
Public Const WM_SETREDRAW      As Long = &HB
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
                                                                       ByVal wMsg As Long, _
                                                                       ByVal wParam As Long, _
                                                                       LParam As Any) As Long

Public Function AccumulatorString(ByVal StrAccum As String, _
                                  VarAdd As Variant, _
                                  Optional Delimiter As String = ",", _
                                  Optional ByVal NoRepeats As Boolean = True) As String

  'Allows you to build up a delimited string with no duplicate members or excess delimiters
  'Call:
  '           SomeString= AccumulatorString (SomeString, SomeOtherData)
  '
  'VarAdd allows you to add array members or strings
  'Optional Delimiter allows you to do further formatting if needed
  'Optional NoRepeats default True exclude duplicates set to false if you want it
  'NOTE if you want to add blanks make sure that VarAdd is at least a single space (" ")

  If LenB(VarAdd) Then
    If LenB(StrAccum) Then
      'Not already collected
      If (NoRepeats And StrAccum <> VarAdd And Left$(StrAccum, Len(VarAdd & Delimiter)) <> VarAdd & Delimiter And InStr(StrAccum, Delimiter & VarAdd & Delimiter) = 0 And Not Right$(StrAccum, Len(Delimiter & VarAdd)) = Delimiter & VarAdd) Or (Not NoRepeats) Then
        AccumulatorString = StrAccum & Delimiter & VarAdd
       Else
        AccumulatorString = StrAccum
      End If
     Else
      AccumulatorString = VarAdd
    End If
   Else
    AccumulatorString = StrAccum
  End If

End Function

Public Function AppendArray(ByVal VarArray As Variant, _
                            ByVal strAdd As String) As Variant

  '
  
  Dim strT   As String

  Dim strTmp As String
  Dim StrDiv As String
  If Not IsEmpty(VarArray) Then
    strTmp = Join(VarArray)
    Do
      StrDiv = RandomString(48, 122, 3, 6)
    Loop While InStr(StrDiv, strTmp)
    strT = Join(VarArray, StrDiv)
  End If
  strT = AccumulatorString(strT, strAdd, StrDiv, False)
  AppendArray = Split(strT, StrDiv)

End Function

Private Function ArrayNoBlanks(arr As Variant) As Variant

  'eliminate blank members of an array
  
  Dim I      As Long

  Dim strTmp As String
  Dim StrDiv As String
  strTmp = Join(arr)
  Do
    StrDiv = RandomString(48, 122, 3, 6)
  Loop While InStr(StrDiv, strTmp)
  For I = LBound(arr) To UBound(arr)
    If Len(arr(I)) Then
      strTmp = strTmp & StrDiv & arr(I)
    End If
  Next I
  strTmp = Mid$(strTmp, Len(StrDiv))
  ArrayNoBlanks = Split(strTmp, StrDiv)

End Function

Public Function GetWorkString(ctrl As Control, _
                              Optional ByVal bCaseInsensitive As Boolean = False, _
                              Optional ByVal bFormatChar As Boolean = False, _
                              Optional ByVal bNoPunct As Boolean = False, _
                              Optional ByVal bPadPunct As Boolean = False) As String

'massage a string to ignore some characters
  GetWorkString = ctrl.Text
  If bCaseInsensitive Then
    GetWorkString = LCase$(GetWorkString)
  End If
  If bFormatChar Then
    GetWorkString = StringStrip(GetWorkString)
  End If
  If bNoPunct Then
    GetWorkString = PunctuationStrip(GetWorkString)
  End If
  If bPadPunct Then
    GetWorkString = PunctuationPad(GetWorkString)
  End If

End Function

Public Function InQSortArray(ByVal SortedArray As Variant, _
                             ByVal FindMe As String) As Boolean

  'binary search to find a member of a quicksorted array
  
  Dim Low    As Long

  Dim Middle As Long
  Dim High   As Long
  Dim Trap   As Boolean
  Dim TestMe As Variant
  If Not IsEmpty(SortedArray) Then
    If Not IsMissing(SortedArray) Then
      If UBound(SortedArray) > -1 Then
        'Binary search module very fast but requires array to be sorted
        Low = LBound(SortedArray)
        High = UBound(SortedArray)
        If High >= Low Then
          ' invert for Descending sorted Arrays
          If SortedArray(Low) > SortedArray(High) Then
            SwapAnyThing Low, High
          End If
          High = High + 1
          Do Until High - Low = 0
            Middle = (Low + High) \ 2
            ' see note below*
            If Trap Then
              Middle = Low
              High = Low
            End If
            TestMe = SortedArray(Middle) ' assign once to test twice
            ' Only tests half the time
            If TestMe >= FindMe Then
              If TestMe = FindMe Then
                InQSortArray = True
                Exit Function
              End If
              High = Middle
             Else
              Low = Middle
            End If
            Trap = (Low = High - 1)
          Loop
         ElseIf High = Low Then
          'single member test
          InQSortArray = SortedArray(Low) = FindMe
        End If
      End If
    End If
  End If

End Function

Public Function PunctuationPad(TString As String) As String

  ' Strip non ascii codes and extra spaces
  
  Dim I        As Long

  Dim arrPunct As Variant
  arrPunct = Array("~", "`", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "+", "-", _
                   "=", "{", "}", "|", "[", "]", "\", ":", ";", "'", "<", ">", "?", ",", ".", "/", Chr$(34))
  For I = LBound(arrPunct) To UBound(arrPunct)
    If InStr(TString, arrPunct(I)) Then
      TString = Replace(TString, arrPunct(I), " " & arrPunct(I) & " ")
    End If
  Next I
  PunctuationPad = Trim$(TString)

End Function

Public Function PunctuationStrip(TString As String) As String

  ' Strip non ascii codes and extra spaces
  
  Dim I        As Long

  Dim arrPunct As Variant
  arrPunct = Array("~", "`", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "+", "-", _
                   "=", "{", "}", "|", "[", "]", "\", ":", ";", "'", "<", ">", "?", ",", ".", "/", Chr$(34))
  For I = LBound(arrPunct) To UBound(arrPunct)
    If InStr(TString, arrPunct(I)) Then
      TString = Join(Split(TString, arrPunct(I)))
    End If
  Next I
  PunctuationStrip = Trim$(TString)

End Function

Public Function QArrayCount(arr As Variant, _
                            strFind As String) As Long

  
  
  Dim lngTestMember As Long

  'set temp array to maximum possible size
  lngTestMember = QSortArrayPos(arr, strFind)
  If lngTestMember > -1 Then
    Do While arr(lngTestMember) = strFind
      QArrayCount = QArrayCount + 1
      'skip over duplicates
      lngTestMember = lngTestMember + 1
      If lngTestMember > UBound(arr) Then
        ' trap to escape if the last member is a duplicate
        Exit Do
      End If
    Loop
  End If

End Function


Public Function QSortArrayPos(ByVal SortedArray As Variant, _
                              ByVal FindMe As String) As Long

  'find a word in a quicksorted array
  
  Dim Low    As Long

  Dim Middle As Long
  Dim High   As Long
  Dim Trap   As Boolean
  Dim TestMe As Variant
  QSortArrayPos = -1 ' default missing
  'Binary search module very fast but requires array to be sorted
  Low = LBound(SortedArray)
  High = UBound(SortedArray)
  If High >= Low Then
    ' invert for Descending sorted Arrays
    If SortedArray(Low) > SortedArray(High) Then
      SwapAnyThing Low, High
    End If
    High = High + 1
    Do Until High - Low = 0
      Middle = (Low + High) \ 2
      ' see note below*
      If Trap Then
        Middle = Low
        High = Low
      End If
      TestMe = SortedArray(Middle) ' assign once to test twice
      If TestMe >= FindMe Then
        ' Only tests half the time
        If TestMe = FindMe Then
          QSortArrayPos = Middle
          Exit Function
        End If
        High = Middle
       Else
        Low = Middle
      End If
      Trap = (Low = High - 1)
    Loop
  End If

End Function

Private Sub QuickSort(AnArray As Variant, _
                      Lo As Long, _
                      Hi As Long, _
                      Optional Ascending As Boolean = True)

  Dim NewHi      As Long

  Dim CurElement As Variant
  Dim NewLo      As Long
  NewLo = Lo
  NewHi = Hi
  CurElement = AnArray((Lo + Hi) / 2)
  Do While (NewLo <= NewHi)
    If Ascending Then
      Do While AnArray(NewLo) < CurElement And NewLo < Hi 'Ascending Core
        NewLo = NewLo + 1
      Loop
      Do While CurElement < AnArray(NewHi) And NewHi > Lo
        NewHi = NewHi - 1
      Loop
     Else
      Do While AnArray(NewLo) > CurElement And NewLo < Hi 'Descending Core
        NewLo = NewLo + 1
      Loop
      Do While CurElement > AnArray(NewHi) And NewHi > Lo
        NewHi = NewHi - 1
      Loop
    End If
    If NewLo <= NewHi Then
      SwapAnyThing AnArray(NewLo), AnArray(NewHi)
      NewLo = NewLo + 1
      NewHi = NewHi - 1
    End If
  Loop
  If Lo < NewHi Then
    QuickSort AnArray, Lo, NewHi, Ascending
  End If
  If NewLo < Hi Then
    QuickSort AnArray, NewLo, Hi, Ascending
  End If

End Sub

Public Function QuickSortArray(ByVal A As Variant, _
                               Optional Ascending As Boolean = True) As Variant

  'reutrn a quick orted array

  On Error GoTo Not_AnArray
  QuickSort A, LBound(A), UBound(A), Ascending
  QuickSortArray = A

Exit Function

Not_AnArray:
  QuickSortArray = Split("")

End Function

Public Function QuickSortUniqueArray(arr As Variant, _
                                     Optional Ascending As Boolean = True) As Variant

  'return a quicksort array or unique members

  QuickSortUniqueArray = StripDuplicateQArray(QuickSortArray(arr, Ascending))

End Function

Public Function QuickSortUniqueCountArray(arr As Variant, _
                                          Optional Ascending As Boolean = True) As Variant

  'return the array sorted alphabetically with frequency attached in brackets

  QuickSortUniqueCountArray = StripAndCountDuplicateQArray(QuickSortArray(arr, Ascending))

End Function

Public Function QuickSortUniqueFrequencyArray(arr As Variant, _
                                              Optional Ascending As Boolean = True) As Variant

  'return the array sorted by frequency of use for each member

  QuickSortUniqueFrequencyArray = QuickSortArray(StripAndCountDuplicateQArray(QuickSortArray(arr), True), Ascending)

End Function

Public Function RandomString(ByVal iLowerBoundAscii As Long, _
                             ByVal iUpperBoundAscii As Long, _
                             ByVal lLowerBoundLength As Long, _
                             ByVal lUpperBoundLength As Long) As String

  'generate a random string to use as a temporary delimiter which
  'cannot be mistaken for genuine part of an array
  '      --Eric Lynn, Ballwin, Missouri
  '        VBPJ TechTips 7th Edition
  
  Dim sHoldString As String

  Dim LCount      As Long
  'Verify boundaries
  If iLowerBoundAscii < 0 Then
    iLowerBoundAscii = 0
  End If
  If iLowerBoundAscii > 255 Then
    iLowerBoundAscii = 255
  End If
  If iUpperBoundAscii < 0 Then
    iUpperBoundAscii = 0
  End If
  If iUpperBoundAscii > 255 Then
    iUpperBoundAscii = 255
  End If
  If lLowerBoundLength < 0 Then
    lLowerBoundLength = 0
  End If
  'Set a random length
  'Create the random string
  For LCount = 1 To Int((CDbl(lUpperBoundLength) - CDbl(lLowerBoundLength) + 1) * Rnd + lLowerBoundLength)
    sHoldString = sHoldString & Chr$(Int((iUpperBoundAscii - iLowerBoundAscii + 1) * Rnd + iLowerBoundAscii))
  Next LCount
  RandomString = sHoldString

End Function

Public Function StringStrip(TString As String) As String ' Strip non ascii codes and extra spaces

  
  Dim TempName As String   ' Working storage

  Dim I        As Long
  For I = 1 To 31
    If InStr(TString, Chr$(I)) Then
      TString = Join(Split(TString, Chr$(I)))
    End If
  Next I
  Do While InStr(TString, "  ") > 0
    TempName = Replace(TString, "  ", " ")
    TString = TempName
  Loop
  StringStrip = Trim$(TString)

End Function

Public Function StripAndCountDuplicateQArray(ByVal arr As Variant, _
                                             Optional ByVal bCountFirst As Boolean = False) As Variant

  'This only works on QuickSorted arrays
  'For unsorted arrays you need 2 nested For Structures and it's a lot slower
  
  Dim lngNewIndex   As Long

  Dim lngDupCount   As Long
  Dim lngCount      As Long
  Dim lngTestMember As Long
  Dim strFormat     As String
  'set temp array to maximum possible size
  ReDim arrTmp(UBound(arr)) As Variant
  If bCountFirst Then
    strFormat = String$(Len(CStr(UBound(arr))), "0")
  End If
  Do
    lngTestMember = lngCount
    'add 1st member (it can't be duplicate;))
    arrTmp(lngNewIndex) = arr(lngTestMember)
    lngDupCount = 0
    Do While arr(lngTestMember) = arr(lngCount)
      lngDupCount = lngDupCount + 1
      'skip over duplicates
      lngCount = lngCount + 1
      If lngCount > UBound(arr) Then
        ' trap to escape if the last member is a duplicate
        Exit Do
      End If
    Loop
    If bCountFirst Then
      arrTmp(lngNewIndex) = "(" & Format$(lngDupCount, strFormat) & ") " & arrTmp(lngNewIndex)
     Else
      arrTmp(lngNewIndex) = arrTmp(lngNewIndex) & " (" & lngDupCount & ")"
    End If
    'increment the temp array counter
    lngNewIndex = lngNewIndex + 1
  Loop Until lngCount > UBound(arr)
  'delete the unused members of the temp array
  'including the empty one generated by last pass over 'lngNewIndex = lngNewIndex + 1'
  ReDim Preserve arrTmp(lngNewIndex - 1) As Variant
  StripAndCountDuplicateQArray = arrTmp

End Function

Public Function StripDuplicateQArray(ByVal arr As Variant) As Variant

  'This only works on QuickSorted arrays
  'For unsorted arrays you need 2 For Structures and it's a lot slower
  
  Dim lngNewIndex   As Long

  Dim lngCount      As Long
  Dim lngTestMember As Long
  'set temp array to maximum possible size
  ReDim arrTmp(UBound(arr)) As Variant
  Do
    lngTestMember = lngCount
    'add 1st member (it can't be duplicate;))
    arrTmp(lngNewIndex) = arr(lngTestMember)
    Do While arr(lngTestMember) = arr(lngCount)
      'skip over duplicates
      lngCount = lngCount + 1
      If lngCount > UBound(arr) Then
        ' trap to escape if the last member is a duplicate
        Exit Do
      End If
    Loop
    'increment the temp array counter
    lngNewIndex = lngNewIndex + 1
  Loop Until lngCount > UBound(arr)
  'delete the unused members of the temp array
  'including the empty one generated by last pass over 'lngNewIndex = lngNewIndex + 1'
  ReDim Preserve arrTmp(lngNewIndex - 1) As Variant
  StripDuplicateQArray = arrTmp

End Function

Private Sub SwapAnyThing(Var1 As Variant, _
                         Var2 As Variant)

  Dim Var3 As Variant

  Var3 = Var1
  Var1 = Var2
  Var2 = Var3

End Sub

''
Public Function QuickSortAppend(ByVal arr As Variant, _
                                varAppend As Variant, _
                                Optional ByVal bAscending As Boolean = True) As Variant

  
  'append a word to a quicksort array only if it is new

  If IsEmpty(arr) Then
    QuickSortAppend = Split(varAppend)
    Exit Function
  End If
  If InQSortArray(arr, varAppend) Then
    QuickSortAppend = arr
   Else
    QuickSortAppend = QuickSortArray(AppendArray(arr, varAppend), bAscending)
  End If

End Function

Public Function QuickSortRemove(ByVal arr As Variant, _
                                varRemove As Variant, _
                                Optional ByVal bAscending As Boolean = True) As Variant

  
  'delete a member of a quicksorted array

  If InQSortArray(arr, varRemove) Then
    arr(QSortArrayPos(arr, varRemove)) = vbNullString
    QuickSortRemove = QuickSortArray(ArrayNoBlanks(arr), bAscending)
   Else
    QuickSortRemove = arr
  End If

End Function

':)Code Fixer V2.4.6 (13/08/2004 3:27:45 PM) 8 + 629 = 637 Lines Thanks Ulli for inspiration and lots of code.

