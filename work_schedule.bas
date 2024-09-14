Attribute VB_Name = "Module1"
Sub GenerateWorkSchedule()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' シートと変数の設定
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
    
    ' ワークシートの参照を設定
    Set wsInput = ThisWorkbook.Sheets("ユーザー入力")
    Set wsOutput = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    Set wsExample = ThisWorkbook.Sheets("作成例")
    Set wsHoliday = ThisWorkbook.Sheets("日本の休日")
    Set wsStaffList = ThisWorkbook.Sheets("要員リスト")
    Set wsVacation = ThisWorkbook.Sheets("要員の休み")

    Set cytologyLeaderSheet = ThisWorkbook.Sheets("細胞診1")
    Set cytology2Sheet = ThisWorkbook.Sheets("細胞診2")
    Set cytology3Sheet = ThisWorkbook.Sheets("細胞診3")
    Set immunoStainerSheet = ThisWorkbook.Sheets("免疫染色")
    Set processorSheet = ThisWorkbook.Sheets("検体処理")
    Set cuttingLeaderSheet = ThisWorkbook.Sheets("切り出し")
    Set supportSheet = ThisWorkbook.Sheets("サポート")
    Set outsideWorkerSheet = ThisWorkbook.Sheets("外回り1")
    Set outsideWorker2Sheet = ThisWorkbook.Sheets("外回り2")
    Set thinSliceSheet = ThisWorkbook.Sheets("包埋薄切")
    Set slicerSheet = ThisWorkbook.Sheets("薄切1")
    Set slicer2Sheet = ThisWorkbook.Sheets("薄切2")
    Set slicer3Sheet = ThisWorkbook.Sheets("薄切3")
    
    ' 年月の入力値を取得
    year = wsInput.Range("B2").Value
    month = wsInput.Range("C2").Value

    ' その他の必要な変数やコレクションの初期設定があればここに追加

    wsOutput.name = year & "年" & month & "月勤務表"

    ' "作成例"のヘッダーをコピー
    ' セルの内容とレイアウトをコピー
    wsExample.Cells.Copy Destination:=wsOutput.Cells


    ' 月の最終日を計算
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
            offCol = 20 ' 休みの列番号、T列は20
            For Each cell In allFound
                staffName = wsVacation.Cells(1, cell.column).Value
                wsOutput.Cells(dayIndex + 1, offCol).Value = staffName
                assignedStaff.Add staffName
                offCol = offCol + 1
            Next cell
        Else
        ' 一致するセルが見つからない場合の処理
        End If
        
        If Not (Weekday(dateToCheck) = vbSaturday Or Weekday(dateToCheck) = vbSunday Or _
            Application.CountIf(wsHoliday.columns(1), dateToCheck) > 0) Then '土曜日曜日本の祝日以外なら、

            
            ' 検体処理に入った人を、翌日の包埋薄切リーダーに配置するコードブロック
            Dim lastStaffRowProcessor As Integer
            lastStaffRowProcessor = wsOutput.Cells(wsOutput.Rows.Count, "F").End(xlUp).row
            ' F列で何か文字が入っているセルの、一番下は何行目か取得してそれをlastStaffRowProcessorと名付ける
            
            If lastStaffRowProcessor = 1 Then
                lastGStaffName = wsOutput.Cells(2, "G").Value
                assignedStaff.Add lastGStaffName
                
            Else
                Dim lastStaffName As String
                lastStaffName = wsOutput.Cells(lastStaffRowProcessor, "F").Value
                ' F列のlastStaffRowProcessor行目に入力された名前をlastStaffNameと名付ける
            
                If lastStaffRowProcessor > 1 Then ' 1より大きい場合のみ実行
                    wsOutput.Cells(dayIndex + 1, "G").Value = lastStaffName
                    assignedStaff.Add lastStaffName
                End If
            End If
            
            Debug.Print "Assigned in G AssignStaff:"
    For Each staff In assignedStaff
        Debug.Print staff
    Next staff

            ' 免疫染色のスタッフ割り当てを実行
    Call CheckBreakAndAssignStaff(dateToCheck, wsOutput, immunoStainerSheet, assignedStaff, dayIndex, "L")
    
    ' 細胞診1のスタッフ割り当てを実行
    Call CheckBreakAndAssignStaff(dateToCheck, wsOutput, cytologyLeaderSheet, assignedStaff, dayIndex, "M")
    
    ' 細胞診2のスタッフ割り当てを実行
    Call CheckBreakAndAssignStaff(dateToCheck, wsOutput, cytology2Sheet, assignedStaff, dayIndex, "N")


        End If

        ' 役割シートからスタッフを配置
        Dim staffOrder As Collection
        Set staffOrder = New Collection
        ' 優先度に従って役割シートと列を追加
        staffOrder.Add Array(processorSheet, "F")
        staffOrder.Add Array(cuttingLeaderSheet, "B")
        ' 以下の順で追加
        staffOrder.Add Array(supportSheet, "C")
        staffOrder.Add Array(outsideWorkerSheet, "D")
        staffOrder.Add Array(outsideWorker2Sheet, "E")
        staffOrder.Add Array(thinSliceSheet, "H")
        staffOrder.Add Array(slicerSheet, "I")
        staffOrder.Add Array(slicer2Sheet, "J")
        staffOrder.Add Array(slicer3Sheet, "K")
        
        staffOrder.Add Array(cytology3Sheet, "O") ' 最低の優先度
        

        ' 優先度に従ってスタッフを配置
        Dim item As Variant

        Dim targetSheet As Worksheet
        Dim targetColumn As String
        For Each item In staffOrder
            Set targetSheet = item(0)
            targetColumn = item(1)
            AssignStaff wsOutput, assignedStaff, targetSheet, targetColumn, year, month, lastDay, dayIndex
        Next item
        
        ' 他の役割シートに対するAssignStaffの呼び出し...
        Dim wsStaff As Range
        Dim staffNameList As Range
        Dim unassignedStaffCol As Integer
        unassignedStaffCol = 23 ' W列の列番号
        Set wsStaff = wsStaffList.columns(1) ' スタッフリストがある列を想定
        Set staffNameList = wsStaff.SpecialCells(xlCellTypeConstants)
        
        If Not (Weekday(dateToCheck) = vbSaturday Or Weekday(dateToCheck) = vbSunday Or _
            Application.CountIf(wsHoliday.columns(1), dateToCheck) > 0) Then
        
            ' 日付に基づいて未割り当てのスタッフを見つけて、wsOutputに配置
            For Each cell In staffNameList
                staffName = cell.Value
                isAlreadyAssigned = False
            
                For Each staff In assignedStaff
                    If staff = staffName Then
                        isAlreadyAssigned = True
                        Exit For
                    End If
                Next staff
            
                ' 未割り当てスタッフを見つけたらwsOutputに配置
                If Not isAlreadyAssigned Then
                    wsOutput.Cells(dayIndex + 1, unassignedStaffCol).Value = staffName
                    unassignedStaffCol = unassignedStaffCol + 1 ' 次の空き列に移動
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
    
    Set wsHoliday = ThisWorkbook.Sheets("日本の休日")
    dateToCheck = DateSerial(year, month, dayIndex)
    
    Set staffRange = staffSheet.columns(1).SpecialCells(xlCellTypeConstants)
    
    ' すべてのスタッフを確認し、未割り当てのスタッフを特定する
    For Each cell In staffRange
        staffName = cell.Value
        isAlreadyAssigned = False
        
        For Each staff In assignedStaff
            If staff = staffName Then
                isAlreadyAssigned = True
                Exit For
            End If
        Next staff
        
        ' 未割り当てのスタッフをコレクションに追加
        If Not isAlreadyAssigned Then
            nonAlreadyAssigned.Add staffName
        End If
    Next cell
    
    
    If Not (Weekday(dateToCheck) = vbSaturday Or Weekday(dateToCheck) = vbSunday Or _
        Application.CountIf(wsHoliday.columns(1), dateToCheck) > 0) Then
        
        ' 未割り当てのスタッフが存在すれば、その中からランダムに選択して割り当てる
        If nonAlreadyAssigned.Count > 0 Then
            Dim randomIndex As Integer
            randomIndex = Application.WorksheetFunction.RandBetween(1, nonAlreadyAssigned.Count)
            staffName = nonAlreadyAssigned(randomIndex) ' ランダムに選ばれたスタッフ名
            wsOutput.Cells(dayIndex + 1, column).Value = staffName
            assignedStaff.Add staffName
        End If
    End If
    
    
End Sub

Function isHoliday(dateToCheck As Date) As Boolean
    ' 土日や祝日を休日とする判定
    If Weekday(dateToCheck) = vbSaturday Or Weekday(dateToCheck) = vbSunday Or _
       Application.CountIf(ThisWorkbook.Sheets("日本の休日").columns(1), dateToCheck) > 0 Then
        isHoliday = True
    Else
        isHoliday = False
    End If
End Function

Sub CheckBreakAndAssignStaff(dateToCheck As Date, wsOutput As Worksheet, targetSheet As Worksheet, assignedStaff As Collection, dayIndex As Integer, columnToAssign As String)
    Dim lastStaff As String, nextStaff As String
    Dim lastStaffRow As Integer, foundRow As Integer
    Dim isAlreadyAssigned As Boolean
    
    ' 指定された列の最終入力行を取得し、その名前をlastStaffに格納
    lastStaffRow = wsOutput.Cells(wsOutput.Rows.Count, columnToAssign).End(xlUp).row
    lastStaff = wsOutput.Cells(lastStaffRow, columnToAssign).Value

    ' 2連休明けかどうかのチェック
    Dim isLongBreak As Boolean
    If isHoliday(dateToCheck - 1) And isHoliday(dateToCheck - 2) Then
        isLongBreak = True
    Else
        isLongBreak = False
    End If

    If isLongBreak Then
        ' targetSheet内でlastStaffと一致するセルを見つけ、その行を取得
        With targetSheet.columns(1)
            Set foundCell = .Find(What:=lastStaff, LookIn:=xlValues, LookAt:=xlWhole)
            If Not foundCell Is Nothing Then
                foundRow = foundCell.row
            End If
        End With

        ' targetSheetのfoundRow+1行目が空か確認
        If IsEmpty(targetSheet.Cells(foundRow + 1, 1)) Then
            nextStaff = targetSheet.Cells(1, 1).Value
        Else
            nextStaff = targetSheet.Cells(foundRow + 1, 1).Value
        End If

        ' nextStaffが既にassignedStaffに格納されているか確認
        isAlreadyAssigned = False
        For Each staff In assignedStaff
            If staff = nextStaff Then
                isAlreadyAssigned = True
                Exit For
            End If
        Next staff

        If isAlreadyAssigned Then
            ' targetSheetのA列からランダムに1人選んで、既にassignedStaffに含まれていないか確認
            Dim totalStaff As Integer, randomRow As Integer
            Dim isStaffAssigned As Boolean
            
            Set staffRange = targetSheet.columns(1).SpecialCells(xlCellTypeConstants)
            totalStaff = staffRange.Cells.Count
            
            Do
                randomRow = Application.WorksheetFunction.RandBetween(1, totalStaff)
                nextStaff = staffRange.Cells(randomRow).Value
                
                ' 選んだスタッフが既にassignedStaffに含まれていないか確認
                isStaffAssigned = False
                For Each staff In assignedStaff
                    If staff = nextStaff Then
                        isStaffAssigned = True
                        Exit For
                    End If
                Next staff
                
            Loop While isStaffAssigned ' 既にassignedStaffに含まれていたら繰り返す
            
            ' nextStaffをwsOutputに出力
            wsOutput.Cells(dayIndex + 1, columnToAssign).Value = nextStaff
            assignedStaff.Add nextStaff
            
        Else
            ' nextStaffをwsOutputに出力
            wsOutput.Cells(dayIndex + 1, columnToAssign).Value = nextStaff
            assignedStaff.Add nextStaff
        End If
    Else
        ' 2連休明けでない場合、lastStaffが既にassignedStaffに格納されているか確認
        isAlreadyAssigned = False
        For Each staff In assignedStaff
            If staff = lastStaff Then
                isAlreadyAssigned = True
                Exit For
            End If
        Next staff

        If isAlreadyAssigned Then
            ' targetSheetのA列からランダムに1人選んで出力
            Set staffRange = targetSheet.columns(1).SpecialCells(xlCellTypeConstants)
            totalStaff = staffRange.Cells.Count
            
            Do
                randomRow = Application.WorksheetFunction.RandBetween(1, totalStaff)
                nextStaff = staffRange.Cells(randomRow).Value
                
                ' 選んだスタッフが既にassignedStaffに含まれていないか確認
                isStaffAssigned = False
                For Each staff In assignedStaff
                    If staff = nextStaff Then
                        isStaffAssigned = True
                        Exit For
                    End If
                Next staff
                
            Loop While isStaffAssigned ' 既にassignedStaffに含まれていたら繰り返す
            
            ' nextStaffをwsOutputに出力
            wsOutput.Cells(dayIndex + 1, columnToAssign).Value = nextStaff
            assignedStaff.Add nextStaff
            
        Else
            ' lastStaffをそのままwsOutputに出力
            wsOutput.Cells(dayIndex + 1, columnToAssign).Value = lastStaff
            assignedStaff.Add lastStaff
        End If
    End If
    
    Debug.Print "Assigned in CheckBreakAndAssignStaff (Column " & columnToAssign & "): " & nextStaff
    For Each staff In assignedStaff
        Debug.Print staff
    Next staff

    ' クリア処理
    nextStaff = ""
    lastStaff = ""
End Sub





