Attribute VB_Name = "Module1"
Sub GenerateWorkSchedule()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' �V�[�g�ƕϐ��̐ݒ�
    Dim wsInput As Worksheet, wsExample As Worksheet, wsHoliday As Worksheet
    Dim wsStaffList As Worksheet, wsVacation As Worksheet
    Dim wsOutput As Worksheet
    Dim cytologyLeaderSheet As Worksheet, cytologySubSheet As Worksheet
    Dim cytology2Sheet As Worksheet, cytology3Sheet As Worksheet
    Dim immunoStainerSheet As Worksheet, stainSubSheet As Worksheet, processorSheet As Worksheet
    Dim cuttingLeaderSheet As Worksheet, supportSheet As Worksheet, outsideWorkerSheet As Worksheet, outsideWorker2Sheet As Worksheet, thinSliceSheet As Worksheet
    Dim year As Integer, month As Integer, lastDay As Integer, dayIndex As Integer
    Dim dateToCheck As Date
    Dim assignedStaff As New Collection
    Dim foundCell As Range, firstFound As Range
    Dim allFound As New Collection
    Dim cell As Range
    
    ' ���[�N�V�[�g�̎Q�Ƃ�ݒ�
    Set wsInput = ThisWorkbook.Sheets("���[�U�[����")
    Set wsOutput = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    Set wsExample = ThisWorkbook.Sheets("�쐬��")
    Set wsHoliday = ThisWorkbook.Sheets("���{�̋x��")
    Set wsStaffList = ThisWorkbook.Sheets("�v�����X�g")
    Set wsVacation = ThisWorkbook.Sheets("�v���̋x��")

    Set cytologyLeaderSheet = ThisWorkbook.Sheets("�זE�f1")
    Set cytology2Sheet = ThisWorkbook.Sheets("�זE�f2")
    Set cytology3Sheet = ThisWorkbook.Sheets("�זE�f3")
    Set immunoStainerSheet = ThisWorkbook.Sheets("�Ɖu���F")
    Set processorSheet = ThisWorkbook.Sheets("���̏���")
    Set cuttingLeaderSheet = ThisWorkbook.Sheets("�؂�o��")
    Set supportSheet = ThisWorkbook.Sheets("�T�|�[�g")
    Set outsideWorkerSheet = ThisWorkbook.Sheets("�O���1")
    Set outsideWorker2Sheet = ThisWorkbook.Sheets("�O���2")
    Set thinSliceSheet = ThisWorkbook.Sheets("�����")
    Set slicerSheet = ThisWorkbook.Sheets("����1")
    Set slicer2Sheet = ThisWorkbook.Sheets("����2")
    Set slicer3Sheet = ThisWorkbook.Sheets("����3")
    
    ' �N���̓��͒l���擾
    year = wsInput.Range("B2").Value
    month = wsInput.Range("C2").Value

    ' ���̑��̕K�v�ȕϐ���R���N�V�����̏����ݒ肪����΂����ɒǉ�

    wsOutput.name = year & "�N" & month & "���Ζ��\"

    ' "�쐬��"�̃w�b�_�[���R�s�[
    ' �Z���̓��e�ƃ��C�A�E�g���R�s�[
    wsExample.Cells.Copy Destination:=wsOutput.Cells


    ' ���̍ŏI�����v�Z
    lastDay = Day(DateSerial(year, month + 1, 0))


    For dayIndex = 1 To lastDay
        dateToCheck = DateSerial(year, month, dayIndex)
        wsOutput.Cells(dayIndex + 1, 1).Value = dateToCheck
        offCol = 20

        Set assignedStaff = New Collection
        Set allFound = New Collection
        With wsVacation.Range("B2:Q40")
            Set foundCell = .Find(What:=dateToCheck, LookIn:=xlValues, LookAt:=xlWhole)
            If Not foundCell Is Nothing Then
                Set firstFound = foundCell
                Do
                    allFound.Add foundCell

                    Set foundCell = .FindNext(foundCell)
                Loop While Not foundCell Is Nothing And foundCell.Address <> firstFound.Address
            End If
        End With

        If allFound.Count > 0 Then
            offCol = 20 ' �x�݂̗�ԍ��AT���20
            For Each cell In allFound
                staffName = wsVacation.Cells(1, cell.column).Value
                wsOutput.Cells(dayIndex + 1, offCol).Value = staffName
                assignedStaff.Add staffName
                offCol = offCol + 1
            Next cell
        Else
        ' ��v����Z����������Ȃ��ꍇ�̏���
        End If
        
        If Not (Weekday(dateToCheck) = vbSaturday Or Weekday(dateToCheck) = vbSunday Or _
            Application.CountIf(wsHoliday.columns(1), dateToCheck) > 0) Then '�y�j���j���{�̏j���ȊO�Ȃ�A

            
            ' ���̏����ɓ������l���A�����̕���؃��[�_�[�ɔz�u����R�[�h�u���b�N
            Dim lastStaffRowProcessor As Integer
            lastStaffRowProcessor = wsOutput.Cells(wsOutput.Rows.Count, "F").End(xlUp).row
            ' F��ŉ��������������Ă���Z���́A��ԉ��͉��s�ڂ��擾���Ă����lastStaffRowProcessor�Ɩ��t����
            
            If lastStaffRowProcessor = 1 Then
                lastGStaffName = wsOutput.Cells(2, "G").Value
                assignedStaff.Add lastGStaffName
                
            Else
                Dim lastStaffName As String
                lastStaffName = wsOutput.Cells(lastStaffRowProcessor, "F").Value
                ' F���lastStaffRowProcessor�s�ڂɓ��͂��ꂽ���O��lastStaffName�Ɩ��t����
            
                If lastStaffRowProcessor > 1 Then ' 1���傫���ꍇ�̂ݎ��s
                    wsOutput.Cells(dayIndex + 1, "G").Value = lastStaffName
                    assignedStaff.Add lastStaffName
                End If
            End If
            
            Debug.Print "Assigned in G AssignStaff:"
    For Each staff In assignedStaff
        Debug.Print staff
    Next staff

            ' �Ɖu���F�̃X�^�b�t���蓖�Ă����s
    Call CheckBreakAndAssignStaff(dateToCheck, wsOutput, immunoStainerSheet, assignedStaff, dayIndex, "L")
    
    ' �זE�f1�̃X�^�b�t���蓖�Ă����s
    Call CheckBreakAndAssignStaff(dateToCheck, wsOutput, cytologyLeaderSheet, assignedStaff, dayIndex, "M")
    
    ' �זE�f2�̃X�^�b�t���蓖�Ă����s
    Call CheckBreakAndAssignStaff(dateToCheck, wsOutput, cytology2Sheet, assignedStaff, dayIndex, "N")


        End If

        ' �����V�[�g����X�^�b�t��z�u
        Dim staffOrder As Collection
        Set staffOrder = New Collection
        ' �D��x�ɏ]���Ė����V�[�g�Ɨ��ǉ�
        staffOrder.Add Array(processorSheet, "F")
        staffOrder.Add Array(cuttingLeaderSheet, "B")
        ' �ȉ��̏��Œǉ�
        staffOrder.Add Array(supportSheet, "C")
        staffOrder.Add Array(outsideWorkerSheet, "D")
        staffOrder.Add Array(outsideWorker2Sheet, "E")
        staffOrder.Add Array(thinSliceSheet, "H")
        staffOrder.Add Array(slicerSheet, "I")
        staffOrder.Add Array(slicer2Sheet, "J")
        staffOrder.Add Array(slicer3Sheet, "K")
        
        staffOrder.Add Array(cytology3Sheet, "O") ' �Œ�̗D��x
        

        ' �D��x�ɏ]���ăX�^�b�t��z�u
        Dim item As Variant

        Dim targetSheet As Worksheet
        Dim targetColumn As String
        For Each item In staffOrder
            Set targetSheet = item(0)
            targetColumn = item(1)
            AssignStaff wsOutput, assignedStaff, targetSheet, targetColumn, year, month, lastDay, dayIndex
        Next item
        
        ' ���̖����V�[�g�ɑ΂���AssignStaff�̌Ăяo��...
        Dim wsStaff As Range
        Dim staffNameList As Range
        Dim unassignedStaffCol As Integer
        unassignedStaffCol = 23 ' W��̗�ԍ�
        Set wsStaff = wsStaffList.columns(1) ' �X�^�b�t���X�g��������z��
        Set staffNameList = wsStaff.SpecialCells(xlCellTypeConstants)
        
        If Not (Weekday(dateToCheck) = vbSaturday Or Weekday(dateToCheck) = vbSunday Or _
            Application.CountIf(wsHoliday.columns(1), dateToCheck) > 0) Then
        
            ' ���t�Ɋ�Â��Ė����蓖�ẴX�^�b�t�������āAwsOutput�ɔz�u
            For Each cell In staffNameList
                staffName = cell.Value
                isAlreadyAssigned = False
            
                For Each staff In assignedStaff
                    If staff = staffName Then
                        isAlreadyAssigned = True
                        Exit For
                    End If
                Next staff
            
                ' �����蓖�ăX�^�b�t����������wsOutput�ɔz�u
                If Not isAlreadyAssigned Then
                    wsOutput.Cells(dayIndex + 1, unassignedStaffCol).Value = staffName
                    unassignedStaffCol = unassignedStaffCol + 1 ' ���̋󂫗�Ɉړ�
                End If
            Next cell
        End If
        
            Set assignedStaff = New Collection
    Next dayIndex
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

Sub AssignStaff(ByRef wsOutput As Worksheet, ByRef assignedStaff As Collection, ByRef staffSheet As Worksheet, ByVal column As String, ByVal year As Integer, ByVal month As Integer, ByVal lastDay As Integer, ByVal dayIndex As Integer)
    Dim dateToCheck As Date
    Dim staffName As String
    Dim staffRange As Range, cell As Range
    Dim nonAlreadyAssigned As New Collection
    Dim wsHoliday As Worksheet
    Dim isAlreadyAssigned As Boolean
    
    Set wsHoliday = ThisWorkbook.Sheets("���{�̋x��")
    dateToCheck = DateSerial(year, month, dayIndex)
    
    Set staffRange = staffSheet.columns(1).SpecialCells(xlCellTypeConstants)
    
    ' ���ׂẴX�^�b�t���m�F���A�����蓖�ẴX�^�b�t����肷��
    For Each cell In staffRange
        staffName = cell.Value
        isAlreadyAssigned = False
        
        For Each staff In assignedStaff
            If staff = staffName Then
                isAlreadyAssigned = True
                Exit For
            End If
        Next staff
        
        ' �����蓖�ẴX�^�b�t���R���N�V�����ɒǉ�
        If Not isAlreadyAssigned Then
            nonAlreadyAssigned.Add staffName
        End If
    Next cell
    
    
    If Not (Weekday(dateToCheck) = vbSaturday Or Weekday(dateToCheck) = vbSunday Or _
        Application.CountIf(wsHoliday.columns(1), dateToCheck) > 0) Then
        
        ' �����蓖�ẴX�^�b�t�����݂���΁A���̒����烉���_���ɑI�����Ċ��蓖�Ă�
        If nonAlreadyAssigned.Count > 0 Then
            Dim randomIndex As Integer
            randomIndex = Application.WorksheetFunction.RandBetween(1, nonAlreadyAssigned.Count)
            staffName = nonAlreadyAssigned(randomIndex) ' �����_���ɑI�΂ꂽ�X�^�b�t��
            wsOutput.Cells(dayIndex + 1, column).Value = staffName
            assignedStaff.Add staffName
        End If
    End If
    
    
End Sub

Function isHoliday(dateToCheck As Date) As Boolean
    ' �y����j�����x���Ƃ��锻��
    If Weekday(dateToCheck) = vbSaturday Or Weekday(dateToCheck) = vbSunday Or _
       Application.CountIf(ThisWorkbook.Sheets("���{�̋x��").columns(1), dateToCheck) > 0 Then
        isHoliday = True
    Else
        isHoliday = False
    End If
End Function

Sub CheckBreakAndAssignStaff(dateToCheck As Date, wsOutput As Worksheet, targetSheet As Worksheet, assignedStaff As Collection, dayIndex As Integer, columnToAssign As String)
    Dim lastStaff As String, nextStaff As String
    Dim lastStaffRow As Integer, foundRow As Integer
    Dim isAlreadyAssigned As Boolean
    
    ' �w�肳�ꂽ��̍ŏI���͍s���擾���A���̖��O��lastStaff�Ɋi�[
    lastStaffRow = wsOutput.Cells(wsOutput.Rows.Count, columnToAssign).End(xlUp).row
    lastStaff = wsOutput.Cells(lastStaffRow, columnToAssign).Value

    ' 2�A�x�������ǂ����̃`�F�b�N
    Dim isLongBreak As Boolean
    If isHoliday(dateToCheck - 1) And isHoliday(dateToCheck - 2) Then
        isLongBreak = True
    Else
        isLongBreak = False
    End If

    If isLongBreak Then
        ' targetSheet����lastStaff�ƈ�v����Z���������A���̍s���擾
        With targetSheet.columns(1)
            Set foundCell = .Find(What:=lastStaff, LookIn:=xlValues, LookAt:=xlWhole)
            If Not foundCell Is Nothing Then
                foundRow = foundCell.row
            End If
        End With

        ' targetSheet��foundRow+1�s�ڂ��󂩊m�F
        If IsEmpty(targetSheet.Cells(foundRow + 1, 1)) Then
            nextStaff = targetSheet.Cells(1, 1).Value
        Else
            nextStaff = targetSheet.Cells(foundRow + 1, 1).Value
        End If

        ' nextStaff������assignedStaff�Ɋi�[����Ă��邩�m�F
        isAlreadyAssigned = False
        For Each staff In assignedStaff
            If staff = nextStaff Then
                isAlreadyAssigned = True
                Exit For
            End If
        Next staff

        If isAlreadyAssigned Then
            ' targetSheet��A�񂩂烉���_����1�l�I��ŁA����assignedStaff�Ɋ܂܂�Ă��Ȃ����m�F
            Dim totalStaff As Integer, randomRow As Integer
            Dim isStaffAssigned As Boolean
            
            Set staffRange = targetSheet.columns(1).SpecialCells(xlCellTypeConstants)
            totalStaff = staffRange.Cells.Count
            
            Do
                randomRow = Application.WorksheetFunction.RandBetween(1, totalStaff)
                nextStaff = staffRange.Cells(randomRow).Value
                
                ' �I�񂾃X�^�b�t������assignedStaff�Ɋ܂܂�Ă��Ȃ����m�F
                isStaffAssigned = False
                For Each staff In assignedStaff
                    If staff = nextStaff Then
                        isStaffAssigned = True
                        Exit For
                    End If
                Next staff
                
            Loop While isStaffAssigned ' ����assignedStaff�Ɋ܂܂�Ă�����J��Ԃ�
            
            ' nextStaff��wsOutput�ɏo��
            wsOutput.Cells(dayIndex + 1, columnToAssign).Value = nextStaff
            assignedStaff.Add nextStaff
            
        Else
            ' nextStaff��wsOutput�ɏo��
            wsOutput.Cells(dayIndex + 1, columnToAssign).Value = nextStaff
            assignedStaff.Add nextStaff
        End If
    Else
        ' 2�A�x�����łȂ��ꍇ�AlastStaff������assignedStaff�Ɋi�[����Ă��邩�m�F
        isAlreadyAssigned = False
        For Each staff In assignedStaff
            If staff = lastStaff Then
                isAlreadyAssigned = True
                Exit For
            End If
        Next staff

        If isAlreadyAssigned Then
            ' targetSheet��A�񂩂烉���_����1�l�I��ŏo��
            Set staffRange = targetSheet.columns(1).SpecialCells(xlCellTypeConstants)
            totalStaff = staffRange.Cells.Count
            
            Do
                randomRow = Application.WorksheetFunction.RandBetween(1, totalStaff)
                nextStaff = staffRange.Cells(randomRow).Value
                
                ' �I�񂾃X�^�b�t������assignedStaff�Ɋ܂܂�Ă��Ȃ����m�F
                isStaffAssigned = False
                For Each staff In assignedStaff
                    If staff = nextStaff Then
                        isStaffAssigned = True
                        Exit For
                    End If
                Next staff
                
            Loop While isStaffAssigned ' ����assignedStaff�Ɋ܂܂�Ă�����J��Ԃ�
            
            ' nextStaff��wsOutput�ɏo��
            wsOutput.Cells(dayIndex + 1, columnToAssign).Value = nextStaff
            assignedStaff.Add nextStaff
            
        Else
            ' lastStaff�����̂܂�wsOutput�ɏo��
            wsOutput.Cells(dayIndex + 1, columnToAssign).Value = lastStaff
            assignedStaff.Add lastStaff
        End If
    End If
    
    Debug.Print "Assigned in CheckBreakAndAssignStaff (Column " & columnToAssign & "): " & nextStaff
    For Each staff In assignedStaff
        Debug.Print staff
    Next staff

    ' �N���A����
    nextStaff = ""
    lastStaff = ""
End Sub





