Attribute VB_Name = "Module1"
Private prevRow As Long ' ������ Ŭ���� ���� �����ϱ� ���� ����
Private targetColumn As String ' �����Ϸ��� ����� ���� �����ϱ� ���� ����

Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim valueToCopy As Variant
    Dim clickedRow As Long
    
    Set ws = ActiveSheet ' ���� ��Ʈ ��������
    targetColumn = "I" ' �����Ϸ��� ����� �� (��: I��)
    
    ' Ŭ���� ���� �� ���� ��������
    clickedRow = Range("K2").Value + 1
    
    ' ������ Ŭ���� ���� ��� ��
    Set targetCell = ws.Cells(clickedRow, targetColumn)
    
    ' I���� ������ �ٸ� ���� ���� ������ ġȯ
    valueToCopy = ws.Range("L10").Value
    
    ' I���� ���� �ٸ� ���� ���� ������ ����
    targetCell.Value = valueToCopy
    ws.Range("L10").Value = Empty
End Sub

Private Sub CommandButton2_Click()
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim valueToCopy As Variant
    Dim clickedRow As Long
    
    Set ws = ActiveSheet ' ���� ��Ʈ ��������
    targetColumn = "I" ' �����Ϸ��� ����� �� (��: I��)
    
    ' Ŭ���� ���� �� ���� ��������
    clickedRow = Range("K2").Value + 1
    
    ' ������ Ŭ���� ���� ��� ��
    Set targetCell = ws.Cells(clickedRow, targetColumn)
    
    ' I���� ������ �ٸ� ���� ���� ������ ġȯ
    valueToCopy = "O"
    
    ' I���� ���� �ٸ� ���� ���� ������ ����
    targetCell.Value = valueToCopy
    ws.Range("L10").Value = Empty
End Sub

Private Sub CommandButton3_Click()
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim valueToCopy As Variant
    Dim clickedRow As Long
    
    Set ws = ActiveSheet ' ���� ��Ʈ ��������
    targetColumn = "I" ' �����Ϸ��� ����� �� (��: I��)
    
    ' Ŭ���� ���� �� ���� ��������
    clickedRow = Range("K2").Value + 1
    
    ' ������ Ŭ���� ���� ��� ��
    Set targetCell = ws.Cells(clickedRow, targetColumn)
    
    ' I���� ������ �ٸ� ���� ���� ������ ġȯ
    valueToCopy = Empty
    
    ' I���� ���� �ٸ� ���� ���� ������ ����
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
    
    Set inputRange = Range("L10") ' ���� ���� ������ ��� ��
    
    ' ����� ���� L13�� ��쿡�� ����
    If Not Intersect(Target, inputRange) Is Nothing Then
        count = Len(inputRange.Value) ' ���� ���� �� ����
        Range("K15").Value = count ' ����� K15�� ǥ��
    End If
    
    
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim selectedRow As Long
    Dim targetRange As Range
    Dim destRange As Range
    
    Dim contentRow As Long
    
    Set targetRange = Range("A:I") ' Ư�� �� ������ �������ּ���
    prevRow = Target.Row ' ������ Ŭ���� ���� ��ȣ ����
    Set destRange = Range("L10")
    
    ' ���õ� ���� �ϳ��� ���̰� Ư�� �� ������ ���ϴ� ��쿡�� ����
    If Target.Cells.count = 1 And Not Intersect(Target, targetRange) Is Nothing Then
        ' ���õ� ���� �� ���� ��������
        selectedRow = Target.Row - 1
        Range("K2").Value = selectedRow
        
        ' VLookup �Լ� �ڵ忡 ����
        On Error Resume Next
        Dim lookupValue As Range
        Dim lookupRange As Range
        Dim result As Variant
    
        Set lookupValue = Range("K2")
        Set lookupRange = Range("A2:I10")
    
        result = Application.VLookup(lookupValue, lookupRange, lookupRange.Columns.count, False)
        On Error GoTo 0
        
        ' ����� ����� ��ġ ����
        If Not IsError(result) Then
            Range("L10").Value = result
            Range("L10").Font.Color = RGB(192, 192, 192)
            Range("L10").Font.Size = 11
        Else
            Range("L10").Value = ""
        End If
        
    End If

    ' I���� ���õ� ���
    If Not Intersect(Target, Range("I:I")) Is Nothing Then
        MsgBox "�������� ������ �����ϼ���. (�ش� ���� �ڵ��̵�)"
        
    ' ��� ���� �̵�
        destRange.Select
    End If

End Sub

