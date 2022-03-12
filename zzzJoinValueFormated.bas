Attribute VB_Name = "zzzJoinValueFormated"
Option Explicit
Option Compare Text
Public Const project_name = "S_joinF"
Public Const project_Version = "1.0"
Public Const urlGithub = ""



#If Mac Then
''
#Else
  #If VBA7 Then
    Private Declare PtrSafe Function SetTimer Lib "User32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As LongPtr, ByVal lpTimerFunc As LongPtr) As Long
    Private Declare PtrSafe Function KillTimer Lib "User32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
  #Else
    Private Declare Function SetTimer Lib "user32" (ByVal HWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal HWnd As Long, ByVal nIDEvent As Long) As Long
  #End If
  #If VBA7 Then
    Private Declare PtrSafe Function GetClipboardData Lib "User32" (ByVal wFormat As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
    Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function EmptyClipboard Lib "User32" () As Long
    Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
    Private Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal length As LongPtr)
  #Else
    Private Declare Function GetClipboardData Lib "User32" (ByVal wFormat As  Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
  #End If

Private Const MaxH = 1450, MaxV = 409

Private Type FontFormatArguments
  action As Long
  target As Range
  caller As Range
  callerAddress As String
  Formula As String
  Cells As Variant
  sentenceSpace As String
End Type


''''///////////////////////////////////////////////////////
Public ContainCells As New VBA.Collection, FitDisable As Boolean
Private Works() As FontFormatArguments
Private Sub s_join_test()
  Debug.Print UBound(Works)
End Sub
Function S_joinF(ByVal toCell As Range, _
                  ByVal sentenceSpace As String, _
                  ParamArray Cells())
  On Error Resume Next

  Dim r As Object, s$, k%, i%
  Set r = Application.caller
  s = r.Address(0, 0, external:=1)
  S_joinF = "S_joinF: " & ChrW(272) & "ang g" & ChrW(7897) & "p"
  k = UBound(Works)
  For i = 1 To k
    With Works(i)
      If .callerAddress = s Then
        Select Case .action
        Case 0, 1: Exit Function
        Case 2: .action = 3
            S_joinF = "S_joinF: Ho" & ChrW(224) & "n th" & ChrW(224) & "nh"
            GoTo n
        Case Else: .action = 0: GoTo r
        End Select
      End If
    End With
  Next
  k = k + 1
r:
  ReDim Preserve Works(1 To k)
  With Works(k)
    Set .caller = r
    .Cells = Cells
    Set .target = toCell
    .callerAddress = s
    .sentenceSpace = sentenceSpace
    .Formula = r.Formula
    .action = 0
  End With
  
n:
  Call SetTimer(0&, 0&, 0, AddressOf S_joinF_callback)
  On Error GoTo 0
End Function
#If VBA7 And Win64 Then
Public Sub S_joinF_callback(ByVal hwnd As LongPtr, ByVal wMsg^, ByVal idEvent As LongPtr, ByVal dwTime^)
#Else
Public Sub S_joinF_callback(ByVal hwnd&, ByVal wMsg&, ByVal idEvent&, ByVal dwTime&)
#End If
  On Error Resume Next
  KillTimer 0&, idEvent
  S_joinF_working
  On Error GoTo 0
End Sub

Sub S_joinF_working()
  On Error Resume Next
  Dim A As Application, b As FontFormatArguments, i&, k&
  Dim u%, su As Boolean, Ac As Boolean, ec As Boolean, c As Object
  u = UBound(Works)
  For i = 1 To u
    b = Works(i)
    Select Case b.action
    Case 0
      If b.caller.Formula = b.Formula Then
        If A Is Nothing Then
          Set c = Application.ActiveSheet
          'Call savedClipboardText
          Set A = b.caller.Parent.Parent.Parent
          su = A.ScreenUpdating
          Ac = A.Calculation
          ec = A.EnableEvents
          If su Then A.ScreenUpdating = False
          If Ac = xlCalculationAutomatic Then A.Calculation = xlCalculationManual
          If ec Then A.EnableEvents = False
        End If
        Works(i).action = 1
        AddCellHasFormatByHtml b.target, b.sentenceSpace, b.Cells
        Works(i).action = 2
        b.caller.value = b.Formula
      End If
n:
    Case 1, 2:
    Case Else:
      k = k + 1
    End Select
  Next
  If k >= u Then
    Erase Works
  End If
  If Not A Is Nothing Then
    If Not c Is Nothing And Not c Is Application.ActiveSheet Then
      c.Activate
    End If
    'Call savedClipboardText
    If su And A.ScreenUpdating <> su Then
      A.ScreenUpdating = su
    End If
    If Ac = xlCalculationAutomatic And Ac <> A.Calculation Then
      A.Calculation = Ac
    End If
    If ec And A.EnableEvents <> ec Then
      A.EnableEvents = ec
    End If
  End If
  On Error GoTo 0
End Sub
Private Sub AddCellHasFormatByHtml_test()
  ''AddCellHasFormatByHtml [B1], " ", Array([C1], [c2], [C3], [C4], [C5])
  AddCellHasFormatByHtml [B1:B7], " ", [C1:C5]
End Sub
Private Sub AddCellHasFormatByHtml(ByVal toCell As Range, ByVal sentenceSpace$, ParamArray Cells())
  Const n_ = vbNullString
  On Error Resume Next
  Dim target, ft As Range, Cell, bCell, cCell, FileName$, s$, s1$, s2$, s3$, s4$, ss4$, s5$, s6$, s7$
  Dim temp$, Addr$, Class&, nClass&, Font&, nFont&, u%
  Dim r, p, p1, p2, i%:

  Dim re, mre, i2, ims
  Set re = glbRegex

  re.Pattern = "(<font[^>]*?>(\s*<(?:[BIU]|(?:SUP)) ?.*?>)*)" & "([^<]+?)(" & _
             "(?:(?:\s*</(?:[BIU]|(?:SUP)).*?>\s*)+|(?:<BR>\s*)|)+</font>)"
  
  temp = IIf(Environ("tmp") <> "", Environ("tmp"), Environ("temp")) & "\VBE\"

  u = UBound(Cells)
  
  For Each cCell In Cells
    Select Case TypeName(cCell)
    Case "Variant()": bCell = cCell
    Case "Range":  bCell = Array(cCell)
    End Select
    For Each Cell In bCell
      If TypeName(Cell) = "Range" Then
        If u = 0 Then
          For Each target In Cell
            Addr = target.Address(0, 0)
            FileName = temp & Addr & "_" & VBA.Timer & ".html"
            GoSub Cell
          Next
        Else
          Set target = Cell
          Addr = target.Address(0, 0)
          FileName = temp & Addr & "_" & VBA.Timer & ".html"
          GoSub Cell
        End If
      End If
    Next
  Next


  Application.DisplayAlerts = False
  Application.Goto toCell, 0
  TextToClipBoard s1 & s2 & s3 & s4 & s5

  Dim rs, cs
  
  rs = toCell.rows.Count
  cs = toCell.Columns.Count
  
  toCell.MergeCells = False
  toCell.Worksheet.Paste
  toCell.Resize(rs, cs).merge
  With toCell
    .HorizontalAlignment = ft.HorizontalAlignment
    .VerticalAlignment = ft.VerticalAlignment
    .WrapText = True
  End With
  If ft.Columns.Address(external:=1) <> ft.Columns.Address(external:=1) Then
    SetNewWidthArea toCell, ft
  End If
  Application.DisplayAlerts = True
e:
Exit Sub
Cell:
  If target.value = Empty Then
    Return
  End If
  If ft Is Nothing Then
    Set ft = target(1, 1)
  End If
  Application.CutCopyMode = False
  With target.Worksheet.Parent.PublishObjects.Add(4, FileName, target.Parent.name, Addr, 0, "cell", "")
    .Publish (True)
    .AutoRepublish = False
    s = readHTMLFile2(FileName)
    .Delete
  End With
  VBA.Kill FileName
  GoSub read
Return
read:
  p = Split(s, """;}", 2, 1)
  If s1 = "" Then
    s1 = p(0) & """;}"
  End If
  p = Split(p(1), "-->", 2, 1)
  s6 = p(0)
  p = Split("-->" & p(1), "<font ", 2, 1)
  If s3 = n_ Then
    s3 = p(0)
  End If
  
  Class = Split(Split(p(0), "class=xl", 2, 1)(1), " ", 2, 1)(0)
  nClass = Class
  
  p = Split("<font " & p(1), "</font>", , 1)
  If s5 = n_ Then
    s5 = p(UBound(p))
  End If
  p(UBound(p)) = ""

l:
  If s2 Like "*xl" & nClass & "*" Then
    nClass = nClass + 1: GoTo l
  Else
    s6 = Replace(s6, "xl" & Class, "xl" & nClass)
    Class = nClass
  End If
  p(0) = "<font class=""xl" & Class & """" & p(0)
  s7 = Join(p, "</font>")
  
  If s2 <> n_ Then
    p1 = Split(s6, ".font")
    For i = 1 To UBound(p1)
      Font = Split(p1(i), vbNewLine)(0)
      nFont = Font
r:
      If s2 Like "*.font" & nFont & "*" Then
        nFont = nFont + 1: GoTo r
      Else
        s6 = Replace(s6, ".font" & Font, ".font" & nFont)
        s7 = Replace(s7, "font" & Font, "font" & nFont)
      End If
    Next
  End If
  s2 = s2 & s6

  Set mre = re.Execute(s7)

  If mre.Count Then
    ss4 = sentenceSpace
    For i2 = 0 To mre.Count - 1
      Set ims = mre(i2).submatches
      s4 = s4 & ims(0) & ss4 & ims(2) & ims(3)
      ss4 = n_
    Next
  End If
Return

End Sub

Sub FileFastSave(Text$, Optional FileName$, Optional ByVal deleteAfterOpen As Boolean)
  If FileName = vbNullString Then
    FileName = IIf(Environ("tmp") <> "", Environ("tmp"), Environ("temp")) & "\VBE\text" & VBA.Timer & ".html"
  End If
  With VBA.CreateObject("ADODB.Stream")
    .Type = 2 'Stream type
    .Charset = "utf-8" 'or utf-16 etc
    .Open
    .WriteText Text
    .SaveToFile FileName, 2 'Save binary data To disk
  End With
  On Error Resume Next
  openFileWithEditor FileName
  If deleteAfterOpen Then
    VBA.Kill FileName
  End If
End Sub
Sub openFileWithEditor(ByVal FileName$)
  On Error Resume Next
  Shell "D:\Program Files\Microsoft VS Code\Code.exe -r """ & FileName & """", vbNormalFocus
  If VBA.Err Then
    VBA.Err.Clear
    Shell "D:\Program Files\Notepad++\notepad++.exe """ & FileName & """", vbNormalFocus
    If VBA.Err Then
      VBA.Err.Clear
      Shell "notepad """ & FileName & """", vbNormalFocus
    End If
  End If
End Sub

Private Sub SetNewHeightArea_test()
  SetNewHeightArea [A26], [d3]
End Sub

Private Function SetNewHeightArea(ByVal NewCell As Range, ByVal CellMerge As Range) As Boolean
  Const MaxV = 409
  Dim h1!, h2!, k&
  h2 = CellMerge.MergeArea.height
  If h2 > MaxV Then
    Exit Function
  End If
  h1 = h2 / 6.05
  NewCell.EntireRow.RowHeight = h1
  If NewCell.height >= h2 Then
    Do
      h1 = h1 - 0.3
      NewCell.EntireRow.RowHeight = h1
      k = k + 1
    Loop Until NewCell.height <= h2
  End If
  Do Until NewCell.height >= h2
    h1 = h1 + 0.1
    k = k + 1
    NewCell.EntireRow.RowHeight = h1
  Loop
  SetNewHeightArea = True
End Function


Function S_Cells(ParamArray Cells()) As String
  Dim s$, p As Object
  On Error Resume Next
  Set ContainCells = New VBA.Collection
  s = "S_Cells:" & Application.caller.Address(0, 0, external:=1)
  S_Cells = s
  Set p = Nothing
  Set p = ContainCells(s)
  If Not p Is Nothing Then
    ContainCells.Remove s
  End If
  ContainCells.Add Cells, s
End Function
Private Function cellsIntersect(cells1 As Range, ByVal cells2 As Range, Optional refcells As Range) As Range
  If cells1 Is Nothing Then
    Set cells1 = cells2
    Exit Function
  ElseIf cells2 Is Nothing Then
    Exit Function
  End If
  If Not cells1.Worksheet Is cells2.Worksheet Then
    Exit Function
  End If
  Set cellsIntersect = Application.Intersect(cells1, cells2)
  Set cells1 = Application.Union(cells1, cells2)
  If refcells Is Nothing Then
    Set refcells = cells1.Worksheet.Range(cells1, cells2)
  Else
    Set refcells = cells1.Worksheet.Range(cells1, refcells)
    Set refcells = cells1.Worksheet.Range(cells2, refcells)
  End If
End Function

Private Function newUnion(cells1 As Range, ByVal cells2 As Range) As Boolean
  If cells1 Is Nothing Then
    Set cells1 = cells2
    Exit Function
  ElseIf cells2 Is Nothing Then
    Exit Function
  End If
  If Not cells1.Worksheet Is cells2.Worksheet Then
    Exit Function
  End If
  newUnion = Not Application.Intersect(cells1, cells2) Is Nothing
  Set cells1 = Application.Union(cells1, cells2)
End Function

Private Function NewHeightArea(ByVal MergeCells As Range, ByVal height!) As Boolean
  Const MaxV = 409
  Set MergeCells = MergeCells.MergeArea
  Dim h1!, h2!, k&, i&, r&, e As Boolean
  i = MergeCells.rows.Count
  If height > MaxV * i Then
    Exit Function
  End If

  Dim t As Single: t = Timer
  h1 = height / i
  GoSub r
  If h2 > height Then
    Do
      h1 = h1 - 0.1
      GoSub r
    Loop Until h2 <= height
  End If
  Do Until h2 >= height
    h1 = h1 + 0.1
    GoSub r
  Loop
e:
  Debug.Print "NewHeightArea-Timer: "; Round(Timer - t, 2)
  NewHeightArea = True

Exit Function
r:
  k = k + 1
  For r = 1 To i
    MergeCells(r, 1).EntireRow.RowHeight = h1
    h2 = MergeCells.EntireRow.height
    If h2 > height - 1 And h2 < height + 1 Then
      GoTo e
    End If
  Next
Return
End Function


Public Function readHTMLFile2(strFile As String) As String
  Dim f As Long, s$: f = FreeFile
  Open strFile For Input As #f
  s = input$(LOF(f), #f)
  Close #f
  ''s = Join(Split(s, vbNewLine & "  "), vbNullString)
  ''s = Join(Split(s, vbNewLine), " ")
  readHTMLFile2 = s
End Function


Function savedClipboardText() As Boolean
  Static ClipboardText$
  If ClipboardText = vbNullString Then
    ClipboardText = ClipBoard
    savedClipboardText = ClipboardText <> vbNullString
  Else
    TextToClipBoard ClipboardText
    ClipboardText = vbNullString
  End If
End Function

Function TextToClipBoard(ByVal Text As String) As String
  #If Mac Then
    With New MSForms.DataObject
      .SetText Text: .PutInClipboard
    End With
  #Else
    #If VBA7 Then
      Dim hGlobalMemory     As LongPtr
      Dim hClipMemory       As LongPtr
      Dim lpGlobalMemory    As LongPtr
    #Else
      Dim hGlobalMemory     As Long
      Dim hClipMemory       As Long
      Dim lpGlobalMemory    As Long
    #End If
    Dim x                     As Long
    hGlobalMemory = GlobalAlloc(&H42, Len(Text) + 1)
    lpGlobalMemory = GlobalLock(hGlobalMemory)
    lpGlobalMemory = lstrcpy(lpGlobalMemory, Text)
    If GlobalUnlock(hGlobalMemory) <> 0 Then
      TextToClipBoard = "Could not unlock memory location. Copy aborted."
      GoTo PrepareToClose
    End If
    If OpenClipboard(0&) = 0 Then
      TextToClipBoard = "Could not open the Clipboard. Copy aborted."
      Exit Function
    End If
    x = EmptyClipboard()
    hClipMemory = SetClipboardData(1, hGlobalMemory)
PrepareToClose:
    If CloseClipboard() = 0 Then
      TextToClipBoard = "Could not close Clipboard."
    End If
  #End If
End Function


Function ClipBoard()
  On Error GoTo OutOfHere
  Const GHND = &H42
  Const CF_TEXT = 1
  Const MAXSIZE = 4096

    #If VBA7 Then
      Dim hGlobalMemory     As LongPtr
      Dim hClipMemory       As LongPtr
      Dim lpGlobalMemory    As LongPtr
      Dim lpClipMemory  As LongPtr
      Dim RetVal As LongPtr
    #Else
      Dim hGlobalMemory     As Long
      Dim hClipMemory       As Long
      Dim lpGlobalMemory    As Long
      Dim lpClipMemory  As Long
   Dim RetVal As Long
    #End If
   
   Dim MyString As String

   If OpenClipboard(0&) = 0 Then
      ''MsgBox "Cannot open Clipboard. Another app. may have it open"
      Exit Function
   End If

   '' Obtain the handle to the global memory
   '' block that is referencing the text.
   hClipMemory = GetClipboardData(CF_TEXT)
   If IsNull(hClipMemory) Then
      MsgBox "Could not allocate memory"
      GoTo OutOfHere
   End If

   '' Lock Clipboard memory so we can reference
   '' the actual data string.
   lpClipMemory = GlobalLock(hClipMemory)

   If Not IsNull(lpClipMemory) Then
      MyString = Space$(MAXSIZE)
      RetVal = lstrcpy(MyString, lpClipMemory)
      RetVal = GlobalUnlock(hClipMemory)

      '' Peel off the null terminating character.
      MyString = Mid(MyString, 1, InStr(1, MyString, Chr$(0), 0) - 1)
   Else
      ''MsgBox "Could not lock memory to copy string from."
   End If

OutOfHere:

   RetVal = CloseClipboard()
   ClipBoard = MyString

End Function

Function glbHTMLFile(Optional staticObject As Boolean, Optional release As Boolean) As Object
  Const s$ = "HTMLFile"
  If staticObject Or release Then
    Static o As Object
    If o Is Nothing And Not release Then Set o = VBA.CreateObject(s)
    Set glbHTMLFile = o
    If release Then Set o = Nothing
  Else
    Set glbHTMLFile = VBA.CreateObject(s)
  End If
End Function


Function glbRegex(Optional staticObject As Boolean, Optional release As Boolean) As Object
  Const s$ = "VBScript.RegExp"
  If staticObject Or release Then
    Static o As Object
    If o Is Nothing And Not release Then
      Set o = VBA.CreateObject(s)
      Set glbRegex = o: GoTo r
    Else
      Set glbRegex = o
    End If
    If release Then Set o = Nothing
  Else
    Set glbRegex = VBA.CreateObject(s): GoTo r
  End If
Exit Function
r:
  With glbRegex: .Global = True: .IgnoreCase = True: .MultiLine = True: End With
End Function

Private Function SetNewWidthArea(ByVal NewCell As Range, ByVal CellMerge As Range) As Boolean
  Dim w!, W2!, k&
  W2 = CellMerge.MergeArea.Columns.Width
  If W2 > MaxH Then
    Exit Function
  End If
  w = W2 / 6.05
  NewCell.EntireColumn.ColumnWidth = w
  If NewCell.Width >= W2 Then
    Do
      w = w - 0.3
      NewCell.EntireColumn.ColumnWidth = w
      k = k + 1
    Loop Until NewCell.Width <= W2
  End If
  Do Until NewCell.Width >= W2
    w = w + 0.1
    k = k + 1
    NewCell.EntireColumn.ColumnWidth = w
  Loop
  SetNewWidthArea = True
End Function


#End If
