Attribute VB_Name = "Module1"
Private prevRow As Long ' 이전에 클릭한 행을 저장하기 위한 변수
Private targetColumn As String ' 변경하려는 대상의 열을 저장하기 위한 변수

Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim valueToCopy As Variant
    Dim clickedRow As Long
    
    Set ws = ActiveSheet ' 현재 시트 가져오기
    targetColumn = "I" ' 변경하려는 대상의 열 (예: I열)
    
    ' 클릭한 셀의 행 정보 가져오기
    clickedRow = Range("K2").Value + 1
    
    ' 이전에 클릭한 행의 대상 열
    Set targetCell = ws.Cells(clickedRow, targetColumn)
    
    ' I열의 정보를 다른 셀에 적은 정보로 치환
    valueToCopy = ws.Range("L10").Value
    
    ' I열의 값을 다른 셀에 적은 정보로 변경
    targetCell.Value = valueToCopy
    ws.Range("L10").Value = Empty
End Sub

Private Sub CommandButton2_Click()
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim valueToCopy As Variant
    Dim clickedRow As Long
    
    Set ws = ActiveSheet ' 현재 시트 가져오기
    targetColumn = "I" ' 변경하려는 대상의 열 (예: I열)
    
    ' 클릭한 셀의 행 정보 가져오기
    clickedRow = Range("K2").Value + 1
    
    ' 이전에 클릭한 행의 대상 열
    Set targetCell = ws.Cells(clickedRow, targetColumn)
    
    ' I열의 정보를 다른 셀에 적은 정보로 치환
    valueToCopy = "O"
    
    ' I열의 값을 다른 셀에 적은 정보로 변경
    targetCell.Value = valueToCopy
    ws.Range("L10").Value = Empty
End Sub

Private Sub CommandButton3_Click()
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim valueToCopy As Variant
    Dim clickedRow As Long
    
    Set ws = ActiveSheet ' 현재 시트 가져오기
    targetColumn = "I" ' 변경하려는 대상의 열 (예: I열)
    
    ' 클릭한 셀의 행 정보 가져오기
    clickedRow = Range("K2").Value + 1
    
    ' 이전에 클릭한 행의 대상 열
    Set targetCell = ws.Cells(clickedRow, targetColumn)
    
    ' I열의 정보를 다른 셀에 적은 정보로 치환
    valueToCopy = Empty
    
    ' I열의 값을 다른 셀에 적은 정보로 변경
    targetCell.Value = valueToCopy
    ws.Range("L10").Value = Empty
End Sub

Private Sub CommandButton4_Click()
    Dim destRange As Range
    Set destRange = Range("D33")
    destRange.Select
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim inputRange As Range
    Dim count As Long
    
    Set inputRange = Range("L10") ' 글자 수를 측정할 대상 셀
    
    ' 변경된 셀이 L13인 경우에만 실행
    If Not Intersect(Target, inputRange) Is Nothing Then
        count = Len(inputRange.Value) ' 셀의 글자 수 측정
        Range("K15").Value = count ' 결과를 K15에 표시
    End If
    
    
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim selectedRow As Long
    Dim targetRange As Range
    Dim destRange As Range
    
    Dim contentRow As Long
    
    Set targetRange = Range("A:I") ' 특정 열 범위로 설정해주세요
    prevRow = Target.Row ' 이전에 클릭한 행의 번호 저장
    Set destRange = Range("L10")
    
    ' 선택된 셀이 하나의 셀이고 특정 열 범위에 속하는 경우에만 실행
    If Target.Cells.count = 1 And Not Intersect(Target, targetRange) Is Nothing Then
        ' 선택된 셀의 행 값을 가져오기
        selectedRow = Target.Row - 1
        Range("K2").Value = selectedRow
        
        ' VLookup 함수 코드에 대입
        On Error Resume Next
        Dim lookupValue As Range
        Dim lookupRange As Range
        Dim result As Variant
    
        Set lookupValue = Range("K2")
        Set lookupRange = Range("A2:I10")
    
        result = Application.VLookup(lookupValue, lookupRange, lookupRange.Columns.count, False)
        On Error GoTo 0
        
        ' 결과를 출력할 위치 설정
        If Not IsError(result) Then
            Range("L10").Value = result
            Range("L10").Font.Color = RGB(192, 192, 192)
            Range("L10").Font.Size = 11
        Else
            Range("L10").Value = ""
        End If
        
    End If

    ' I열이 선택된 경우
    If Not Intersect(Target, Range("I:I")) Is Nothing Then
        MsgBox "수정사항 셀에서 수정하세요. (해당 셀로 자동이동)"
        
    ' 대상 셀로 이동
        destRange.Select
    End If

End Sub

