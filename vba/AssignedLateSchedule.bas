Attribute VB_Name = "AssignedLateSchedule"
Option Explicit

' Assigned Late Schedule
' Import this module into the workbook's VBA project.  All addresses are
' resolved through workbook tables and names so the paper layout can move.

Private Const LOCATION_ONE As String = "1 Centre Street"
Private Const LOCATION_TWO As String = "100 Centre Street"
Private Const PAIR_SEPARATOR As String = " / "

Private Type ScheduleSlot
    ScheduleDate As Date
    Location As String
    TableRow As Long
    TableColumn As Long
    Mechanic As String
End Type

' Main supervisor action: build dates, assign every empty future slot, and
' permanently record the resulting assignments.
Public Sub GenerateSchedule()
    On Error GoTo HandleError
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    GenerateDatesInternal
    AssignFutureSchedule
    PrepareScheduleForPrinting

    MsgBox "The late schedule has been generated.", vbInformation, "Assigned Late Schedule"
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
HandleError:
    MsgBox Err.Description, vbExclamation, "Schedule not generated"
    Resume CleanExit
End Sub

' Generates only missing schedule dates from the three settings on Schedule.
Public Sub GenerateDates()
    On Error GoTo HandleError
    Application.ScreenUpdating = False
    GenerateDatesInternal
    MsgBox "Schedule dates have been generated.", vbInformation, "Assigned Late Schedule"
CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
HandleError:
    MsgBox Err.Description, vbExclamation, "Dates not generated"
    Resume CleanExit
End Sub

' Removes assignments only on dates that have not yet occurred.  History is
' intentionally untouched; it is the permanent fairness record.
Public Sub ClearFutureSchedule()
    Dim schedule As ListObject, scheduleRow As ListRow
    Dim dateIndex As Long, firstLocationIndex As Long, secondLocationIndex As Long
    Dim cleared As Long

    If MsgBox("Clear assignments for future dates only? History will be kept.", _
              vbQuestion + vbYesNo, "Clear Future Schedule") <> vbYes Then Exit Sub

    Set schedule = GetTable("ScheduleTable")
    dateIndex = GetColumnIndex(schedule, "Date")
    firstLocationIndex = GetColumnIndex(schedule, LOCATION_ONE)
    secondLocationIndex = GetColumnIndex(schedule, LOCATION_TWO)

    For Each scheduleRow In schedule.ListRows
        If IsFutureScheduleDate(scheduleRow.Range.Cells(1, dateIndex).Value) Then
            scheduleRow.Range.Cells(1, firstLocationIndex).ClearContents
            scheduleRow.Range.Cells(1, secondLocationIndex).ClearContents
            cleared = cleared + 1
        End If
    Next scheduleRow
    MsgBox cleared & " future schedule row(s) cleared.", vbInformation, "Clear Future Schedule"
End Sub

' Formats the familiar paper schedule as a single printable page where possible.
Public Sub PrintSchedule()
    On Error GoTo HandleError
    PrepareScheduleForPrinting
    GetSheet("Schedule").PrintOut
    Exit Sub
HandleError:
    MsgBox "The schedule could not be printed: " & Err.Description, vbExclamation, "Print Schedule"
End Sub

' Date generation ------------------------------------------------------------

Private Sub GenerateDatesInternal()
    Dim startDate As Date, requestedDay As Long, numberOfWeeks As Long
    Dim schedule As ListObject, scheduleRow As ListRow
    Dim dateIndex As Long, firstLocationIndex As Long, secondLocationIndex As Long
    Dim generatedDate As Date, weekNumber As Long
    Dim existingDates As Object

    startDate = GetNamedDate("ScheduleStartDate")
    numberOfWeeks = GetNamedPositiveLong("ScheduleWeekCount")
    requestedDay = WeekdayNumber(GetNamedText("ScheduleDayOfWeek"))
    Set schedule = GetTable("ScheduleTable")
    dateIndex = GetColumnIndex(schedule, "Date")
    firstLocationIndex = GetColumnIndex(schedule, LOCATION_ONE)
    secondLocationIndex = GetColumnIndex(schedule, LOCATION_TWO)

    RemoveUnusedFutureBlankDates schedule, dateIndex, firstLocationIndex, secondLocationIndex
    Set existingDates = ExistingScheduleDates(schedule, dateIndex)
    generatedDate = FirstMatchingDate(startDate, requestedDay)

    For weekNumber = 1 To numberOfWeeks
        If Not existingDates.Exists(DateKey(generatedDate)) Then
            Set scheduleRow = FirstEmptyScheduleRow(schedule, dateIndex)
            If scheduleRow Is Nothing Then Set scheduleRow = schedule.ListRows.Add
            scheduleRow.Range.Cells(1, dateIndex).Value = generatedDate
            scheduleRow.Range.Cells(1, dateIndex).NumberFormat = "dddd, m/d/yyyy"
            existingDates.Add DateKey(generatedDate), True
        End If
        generatedDate = DateAdd("d", 7, generatedDate)
    Next weekNumber

    SortScheduleByDate schedule, dateIndex
End Sub

' Clears obsolete, unassigned future dates. Any scheduled or historic date is
' preserved, which prevents changing settings from erasing completed work.
Private Sub RemoveUnusedFutureBlankDates(ByVal schedule As ListObject, ByVal dateIndex As Long, _
                                         ByVal firstLocationIndex As Long, ByVal secondLocationIndex As Long)
    Dim scheduleRow As ListRow
    For Each scheduleRow In schedule.ListRows
        If IsFutureScheduleDate(scheduleRow.Range.Cells(1, dateIndex).Value) Then
            If IsBlank(scheduleRow.Range.Cells(1, firstLocationIndex).Value) And _
               IsBlank(scheduleRow.Range.Cells(1, secondLocationIndex).Value) Then
                scheduleRow.Range.Cells(1, dateIndex).ClearContents
            End If
        End If
    Next scheduleRow
End Sub

' Assignment engine ----------------------------------------------------------

' Builds all future empty slots first, selects across the entire set, then
' performs pair/location/fairness swap passes until no swap improves the plan.
Private Sub AssignFutureSchedule()
    Dim eligible As Collection, slots() As ScheduleSlot
    Dim slotCount As Long
    Dim totalCounts As Object, recentCounts As Object, locationCounts As Object
    Dim lastAssigned As Object, pairCounts As Object

    Set eligible = ActiveMechanics()
    If eligible.Count < 2 Then Err.Raise vbObjectError + 101, , _
        "At least two active mechanics are required to generate a schedule."

    slotCount = CollectEmptyFutureSlots(slots)
    If slotCount = 0 Then Exit Sub

    Set totalCounts = LifetimeCounts(eligible)
    Set recentCounts = RecentCounts(eligible, DateAdd("m", -3, Date))
    Set locationCounts = LifetimeLocationCounts(eligible)
    Set lastAssigned = LastAssignmentDates(eligible)
    Set pairCounts = LifetimePairCounts

    AssignSlotsGlobally slots, slotCount, eligible, totalCounts, recentCounts, _
                        locationCounts, lastAssigned, pairCounts
    ImproveSchedule slots, slotCount, eligible, totalCounts, recentCounts, _
                    locationCounts, lastAssigned, pairCounts
    WriteSlotsToSchedule slots, slotCount
    AppendScheduleToHistory slots, slotCount
End Sub

' Selects the next assignment from every remaining date/location rather than
' treating each schedule row independently.
Private Sub AssignSlotsGlobally(ByRef slots() As ScheduleSlot, ByVal slotCount As Long, _
                                ByVal eligible As Collection, ByVal totalCounts As Object, _
                                ByVal recentCounts As Object, ByVal locationCounts As Object, _
                                ByVal lastAssigned As Object, ByVal pairCounts As Object)
    Dim assigned As Long, bestSlot As Long, bestMechanic As String
    Dim candidateSlot As Long, candidate As Variant, score As Double, bestScore As Double

    For assigned = 1 To slotCount
        bestScore = 1E+30
        bestSlot = 0
        bestMechanic = vbNullString
        For candidateSlot = 1 To slotCount
            If Len(slots(candidateSlot).Mechanic) = 0 Then
                For Each candidate In eligible
                    If IsEligibleForSlot(CStr(candidate), slots(candidateSlot), slots, slotCount) Then
                        score = AssignmentScore(CStr(candidate), slots(candidateSlot), slots, slotCount, _
                                                totalCounts, recentCounts, locationCounts, lastAssigned, pairCounts)
                        If score < bestScore Or (score = bestScore And CStr(candidate) < bestMechanic) Then
                            bestScore = score
                            bestSlot = candidateSlot
                            bestMechanic = CStr(candidate)
                        End If
                    End If
                Next candidate
            End If
        Next candidateSlot
        If bestSlot = 0 Then Err.Raise vbObjectError + 102, , _
            "No eligible mechanic is available for every future schedule date."

        slots(bestSlot).Mechanic = bestMechanic
        ApplyTemporaryAssignment bestMechanic, slots(bestSlot), totalCounts, recentCounts, _
                                 locationCounts, lastAssigned, pairCounts, slots, slotCount
    Next assigned
End Sub

' Scores a candidate lexicographically through heavily separated terms:
' lifetime total, time since last assignment, recent load, consecutive dates,
' repeated partnerships, location balance, and maximum-load spread.
Private Function AssignmentScore(ByVal mechanic As String, ByRef slot As ScheduleSlot, _
                                 ByRef slots() As ScheduleSlot, ByVal slotCount As Long, _
                                 ByVal totalCounts As Object, ByVal recentCounts As Object, _
                                 ByVal locationCounts As Object, ByVal lastAssigned As Object, _
                                 ByVal pairCounts As Object) As Double
    Dim lifetime As Long, recent As Long, daysSince As Long, locationGap As Long
    Dim maximumAfterAssignment As Long, minimumAfterAssignment As Long
    Dim consecutivePenalty As Long, pairPenalty As Long
    lifetime = CLng(totalCounts(mechanic))
    recent = CLng(recentCounts(mechanic))
    daysSince = DaysSinceLastAssignment(lastAssigned(mechanic), slot.ScheduleDate)
    locationGap = LocationImbalanceAfter(mechanic, slot.Location, locationCounts)
    GetCountSpreadAfter mechanic, totalCounts, maximumAfterAssignment, minimumAfterAssignment
    consecutivePenalty = ConsecutiveDatePenalty(mechanic, slot.ScheduleDate, slots, slotCount)
    pairPenalty = PairRepeatPenalty(mechanic, slot, slots, slotCount, pairCounts)

    AssignmentScore = (maximumAfterAssignment - minimumAfterAssignment) * 1000000000# _
                    + lifetime * 1000000# _
                    + recent * 10000# _
                    + consecutivePenalty * 1000# _
                    + pairPenalty * 100# _
                    + locationGap * 10# _
                    - WorksheetFunction.Min(daysSince, 9999) / 100000#
End Function

' Enforces availability and prevents a mechanic being assigned to both locations
' on the same date.
Private Function IsEligibleForSlot(ByVal mechanic As String, ByRef slot As ScheduleSlot, _
                                   ByRef slots() As ScheduleSlot, ByVal slotCount As Long) As Boolean
    Dim index As Long
    If IsUnavailable(mechanic, slot.ScheduleDate) Then Exit Function
    For index = 1 To slotCount
        If slots(index).ScheduleDate = slot.ScheduleDate And slots(index).Mechanic = mechanic Then Exit Function
    Next index
    IsEligibleForSlot = True
End Function

Private Sub ApplyTemporaryAssignment(ByVal mechanic As String, ByRef slot As ScheduleSlot, _
                                     ByVal totalCounts As Object, ByVal recentCounts As Object, _
                                     ByVal locationCounts As Object, ByVal lastAssigned As Object, _
                                     ByVal pairCounts As Object, ByRef slots() As ScheduleSlot, ByVal slotCount As Long)
    Dim partner As String
    totalCounts(mechanic) = CLng(totalCounts(mechanic)) + 1
    recentCounts(mechanic) = CLng(recentCounts(mechanic)) + 1
    locationCounts(mechanic & "|" & slot.Location) = CLng(locationCounts(mechanic & "|" & slot.Location)) + 1
    lastAssigned(mechanic) = slot.ScheduleDate
    partner = AssignedPartner(slot, slots, slotCount)
    If Len(partner) > 0 Then pairCounts(PairKey(mechanic, partner)) = CLng(pairCounts(PairKey(mechanic, partner))) + 1
End Sub

' Balancing pass: swaps two provisional assignments only when the full fairness
' objective decreases, repeatedly, until a complete pass finds no improvement.
Private Sub ImproveSchedule(ByRef slots() As ScheduleSlot, ByVal slotCount As Long, _
                            ByVal eligible As Collection, ByVal totalCounts As Object, _
                            ByVal recentCounts As Object, ByVal locationCounts As Object, _
                            ByVal lastAssigned As Object, ByVal pairCounts As Object)
    Dim changed As Boolean, firstSlot As Long, secondSlot As Long
    Dim beforeScore As Double, afterScore As Double, firstMechanic As String, secondMechanic As String

    Do
        changed = False
        For firstSlot = 1 To slotCount - 1
            For secondSlot = firstSlot + 1 To slotCount
                firstMechanic = slots(firstSlot).Mechanic
                secondMechanic = slots(secondSlot).Mechanic
                If firstMechanic <> secondMechanic Then
                    If IsUnavailable(firstMechanic, slots(secondSlot).ScheduleDate) = False And _
                       IsUnavailable(secondMechanic, slots(firstSlot).ScheduleDate) = False And _
                       Not CreatesSameDayDuplicate(firstMechanic, secondSlot, slots, slotCount) And _
                       Not CreatesSameDayDuplicate(secondMechanic, firstSlot, slots, slotCount) Then
                        beforeScore = LocalFairnessScore(slots, slotCount)
                        slots(firstSlot).Mechanic = secondMechanic
                        slots(secondSlot).Mechanic = firstMechanic
                        afterScore = LocalFairnessScore(slots, slotCount)
                        If afterScore < beforeScore Then
                            changed = True
                        Else
                            slots(firstSlot).Mechanic = firstMechanic
                            slots(secondSlot).Mechanic = secondMechanic
                        End If
                    End If
                End If
            Next secondSlot
        Next firstSlot
    Loop While changed
End Sub

' A compact plan score used by the local improvement pass. The constants retain
' the required priority: count spread, consecutive dates, pair repetition, then
' location imbalance.
Private Function LocalFairnessScore(ByRef slots() As ScheduleSlot, ByVal slotCount As Long) As Double
    Dim totals As Object, locations As Object, pairs As Object, index As Long
    Dim maxCount As Long, minCount As Long, mechanic As Variant, score As Double
    Set totals = CreateObject("Scripting.Dictionary")
    Set locations = CreateObject("Scripting.Dictionary")
    Set pairs = CreateObject("Scripting.Dictionary")
    For index = 1 To slotCount
        Increment totals, slots(index).Mechanic
        Increment locations, slots(index).Mechanic & "|" & slots(index).Location
        If Len(AssignedPartner(slots(index), slots, slotCount)) > 0 Then _
            Increment pairs, PairKey(slots(index).Mechanic, AssignedPartner(slots(index), slots, slotCount))
        score = score + ConsecutiveDatePenalty(slots(index).Mechanic, slots(index).ScheduleDate, slots, slotCount) * 1000#
    Next index
    minCount = 2147483647
    For Each mechanic In totals.Keys
        maxCount = WorksheetFunction.Max(maxCount, CLng(totals(mechanic)))
        minCount = WorksheetFunction.Min(minCount, CLng(totals(mechanic)))
    Next mechanic
    For Each mechanic In pairs.Keys
        score = score + CLng(pairs(mechanic)) * CLng(pairs(mechanic)) * 100#
    Next mechanic
    LocalFairnessScore = score + (maxCount - minCount) * 1000000000#
End Function

' Writes the two people for each location in the paper-compatible "Name / Name"
' format. It never changes populated schedule cells.
Private Sub WriteSlotsToSchedule(ByRef slots() As ScheduleSlot, ByVal slotCount As Long)
    Dim schedule As ListObject, index As Long, scheduleRow As ListRow
    Dim grouped As Object, key As String, cellValue As String
    Set schedule = GetTable("ScheduleTable")
    Set grouped = CreateObject("Scripting.Dictionary")
    For index = 1 To slotCount
        key = CStr(slots(index).TableRow) & "|" & CStr(slots(index).TableColumn)
        If grouped.Exists(key) Then
            grouped(key) = grouped(key) & PAIR_SEPARATOR & slots(index).Mechanic
        Else
            grouped.Add key, slots(index).Mechanic
        End If
    Next index
    For Each key In grouped.Keys
        Set scheduleRow = schedule.ListRows(CLng(Split(key, "|")(0)))
        cellValue = CStr(grouped(key))
        If IsBlank(scheduleRow.Range.Cells(1, CLng(Split(key, "|")(1))).Value) Then
            scheduleRow.Range.Cells(1, CLng(Split(key, "|")(1))).Value = cellValue
        End If
    Next key
End Sub

' History --------------------------------------------------------------------

' Adds every newly generated assignment to the permanent History table. The
' duplicate key makes rerunning Generate Schedule safe.
Private Sub AppendScheduleToHistory(ByRef slots() As ScheduleSlot, ByVal slotCount As Long)
    Dim history As ListObject, existing As Object, index As Long, historyRow As ListRow
    Dim key As String
    Set history = GetTable("HistoryTable")
    Set existing = HistoryKeys(history)
    For index = 1 To slotCount
        key = DateKey(slots(index).ScheduleDate) & "|" & slots(index).Location & "|" & slots(index).Mechanic
        If Not existing.Exists(key) Then
            Set historyRow = history.ListRows.Add
            historyRow.Range.Cells(1, GetColumnIndex(history, "Date")).Value = slots(index).ScheduleDate
            historyRow.Range.Cells(1, GetColumnIndex(history, "Location")).Value = slots(index).Location
            historyRow.Range.Cells(1, GetColumnIndex(history, "Mechanic")).Value = slots(index).Mechanic
            historyRow.Range.Cells(1, GetColumnIndex(history, "Assignment Created On")).Value = Now
            existing.Add key, True
        End If
    Next index
End Sub

' Statistics -----------------------------------------------------------------

Private Function ActiveMechanics() As Collection
    Dim mechanics As ListObject, row As ListRow, result As New Collection
    Set mechanics = GetTable("MechanicsTable")
    For Each row In mechanics.ListRows
        If LCase$(Trim$(CStr(row.Range.Cells(1, GetColumnIndex(mechanics, "Active (Yes/No)")).Value))) = "yes" _
           And Len(Trim$(CStr(row.Range.Cells(1, GetColumnIndex(mechanics, "Mechanic Name")).Value))) > 0 Then
            result.Add Trim$(CStr(row.Range.Cells(1, GetColumnIndex(mechanics, "Mechanic Name")).Value))
        End If
    Next row
    Set ActiveMechanics = result
End Function

Private Function LifetimeCounts(ByVal mechanics As Collection) As Object
    Set LifetimeCounts = HistoryStatistic(mechanics, "TOTAL", 0)
End Function

Private Function RecentCounts(ByVal mechanics As Collection, ByVal cutoffDate As Date) As Object
    Set RecentCounts = HistoryStatistic(mechanics, "RECENT", cutoffDate)
End Function

Private Function LifetimeLocationCounts(ByVal mechanics As Collection) As Object
    Dim result As Object, mechanic As Variant
    Set result = CreateObject("Scripting.Dictionary")
    For Each mechanic In mechanics
        result.Add CStr(mechanic) & "|" & LOCATION_ONE, 0
        result.Add CStr(mechanic) & "|" & LOCATION_TWO, 0
    Next mechanic
    AddHistoryToLocationCounts result
    Set LifetimeLocationCounts = result
End Function

Private Function HistoryStatistic(ByVal mechanics As Collection, ByVal statisticType As String, ByVal cutoffDate As Date) As Object
    Dim result As Object, mechanic As Variant, history As ListObject, row As ListRow, name As String
    Set result = CreateObject("Scripting.Dictionary")
    For Each mechanic In mechanics: result.Add CStr(mechanic), 0: Next mechanic
    Set history = GetTable("HistoryTable")
    For Each row In history.ListRows
        name = Trim$(CStr(row.Range.Cells(1, GetColumnIndex(history, "Mechanic")).Value))
        If result.Exists(name) Then
            If statisticType = "TOTAL" Or row.Range.Cells(1, GetColumnIndex(history, "Date")).Value >= cutoffDate Then
                result(name) = CLng(result(name)) + 1
            End If
        End If
    Next row
    Set HistoryStatistic = result
End Function

Private Function LastAssignmentDates(ByVal mechanics As Collection) As Object
    Dim result As Object, mechanic As Variant, history As ListObject, row As ListRow, name As String, scheduled As Date
    Set result = CreateObject("Scripting.Dictionary")
    For Each mechanic In mechanics: result.Add CStr(mechanic), DateSerial(1900, 1, 1): Next mechanic
    Set history = GetTable("HistoryTable")
    For Each row In history.ListRows
        name = Trim$(CStr(row.Range.Cells(1, GetColumnIndex(history, "Mechanic")).Value))
        If result.Exists(name) And IsDate(row.Range.Cells(1, GetColumnIndex(history, "Date")).Value) Then
            scheduled = CDate(row.Range.Cells(1, GetColumnIndex(history, "Date")).Value)
            If scheduled > CDate(result(name)) Then result(name) = scheduled
        End If
    Next row
    Set LastAssignmentDates = result
End Function

Private Function LifetimePairCounts() As Object
    Dim result As Object, grouped As Object, history As ListObject, row As ListRow
    Dim groupKey As String, mechanic As String, pair As Variant, people As Collection
    Set result = CreateObject("Scripting.Dictionary")
    Set grouped = CreateObject("Scripting.Dictionary")
    Set history = GetTable("HistoryTable")
    For Each row In history.ListRows
        If IsDate(row.Range.Cells(1, GetColumnIndex(history, "Date")).Value) Then
            mechanic = Trim$(CStr(row.Range.Cells(1, GetColumnIndex(history, "Mechanic")).Value))
            If Len(mechanic) > 0 Then
                groupKey = DateKey(CDate(row.Range.Cells(1, GetColumnIndex(history, "Date")).Value)) & "|" & _
                           CStr(row.Range.Cells(1, GetColumnIndex(history, "Location")).Value)
                If grouped.Exists(groupKey) Then
                    grouped(groupKey).Add mechanic
                Else
                    Set people = New Collection
                    people.Add mechanic
                    grouped.Add groupKey, people
                End If
            End If
        End If
    Next row
    For Each pair In grouped.Keys
        If grouped(pair).Count = 2 Then Increment result, PairKey(grouped(pair)(1), grouped(pair)(2))
    Next pair
    Set LifetimePairCounts = result
End Function

' Helpers --------------------------------------------------------------------

Private Function CollectEmptyFutureSlots(ByRef slots() As ScheduleSlot) As Long
    Dim schedule As ListObject, row As ListRow, rowNumber As Long, index As Long, count As Long
    Dim dateIndex As Long, locationIndex As Variant, mechanicsNeeded As Long
    Set schedule = GetTable("ScheduleTable")
    dateIndex = GetColumnIndex(schedule, "Date")
    For Each row In schedule.ListRows
        rowNumber = row.Index
        If IsFutureScheduleDate(row.Range.Cells(1, dateIndex).Value) Then
            For Each locationIndex In Array(GetColumnIndex(schedule, LOCATION_ONE), GetColumnIndex(schedule, LOCATION_TWO))
                If IsBlank(row.Range.Cells(1, CLng(locationIndex)).Value) Then
                    For mechanicsNeeded = 1 To 2
                        count = count + 1
                        ReDim Preserve slots(1 To count)
                        slots(count).ScheduleDate = CDate(row.Range.Cells(1, dateIndex).Value)
                        slots(count).Location = schedule.ListColumns(CLng(locationIndex)).Name
                        slots(count).TableRow = rowNumber
                        slots(count).TableColumn = CLng(locationIndex)
                    Next mechanicsNeeded
                End If
            Next locationIndex
        End If
    Next row
    CollectEmptyFutureSlots = count
End Function

Private Function IsUnavailable(ByVal mechanic As String, ByVal scheduleDate As Date) As Boolean
    Dim availability As ListObject, row As ListRow
    Set availability = GetTable("AvailabilityTable")
    For Each row In availability.ListRows
        If StrComp(Trim$(CStr(row.Range.Cells(1, GetColumnIndex(availability, "Mechanic")).Value)), mechanic, vbTextCompare) = 0 _
           And IsDate(row.Range.Cells(1, GetColumnIndex(availability, "Date")).Value) Then
            If CDate(row.Range.Cells(1, GetColumnIndex(availability, "Date")).Value) = scheduleDate Then
                IsUnavailable = True
                Exit Function
            End If
        End If
    Next row
End Function

Private Function AssignedPartner(ByRef slot As ScheduleSlot, ByRef slots() As ScheduleSlot, ByVal slotCount As Long) As String
    Dim index As Long
    For index = 1 To slotCount
        If slots(index).ScheduleDate = slot.ScheduleDate And slots(index).Location = slot.Location _
           And slots(index).Mechanic <> slot.Mechanic And Len(slots(index).Mechanic) > 0 Then
            AssignedPartner = slots(index).Mechanic
            Exit Function
        End If
    Next index
End Function

Private Function ConsecutiveDatePenalty(ByVal mechanic As String, ByVal scheduleDate As Date, _
                                        ByRef slots() As ScheduleSlot, ByVal slotCount As Long) As Long
    Dim index As Long
    For index = 1 To slotCount
        If slots(index).Mechanic = mechanic Then
            If Abs(DateDiff("d", slots(index).ScheduleDate, scheduleDate)) = 7 Then
                ConsecutiveDatePenalty = 1
                Exit Function
            End If
        End If
    Next index
End Function

Private Function PairRepeatPenalty(ByVal mechanic As String, ByRef slot As ScheduleSlot, _
                                   ByRef slots() As ScheduleSlot, ByVal slotCount As Long, ByVal pairCounts As Object) As Long
    Dim partner As String
    partner = AssignedPartner(slot, slots, slotCount)
    If Len(partner) > 0 And pairCounts.Exists(PairKey(mechanic, partner)) Then _
        PairRepeatPenalty = CLng(pairCounts(PairKey(mechanic, partner)))
End Function

Private Function LocationImbalanceAfter(ByVal mechanic As String, ByVal location As String, ByVal locationCounts As Object) As Long
    Dim oneCount As Long, twoCount As Long
    oneCount = CLng(locationCounts(mechanic & "|" & LOCATION_ONE))
    twoCount = CLng(locationCounts(mechanic & "|" & LOCATION_TWO))
    If location = LOCATION_ONE Then oneCount = oneCount + 1 Else twoCount = twoCount + 1
    LocationImbalanceAfter = Abs(oneCount - twoCount)
End Function

Private Sub GetCountSpreadAfter(ByVal selectedMechanic As String, ByVal totalCounts As Object, _
                                ByRef maximum As Long, ByRef minimum As Long)
    Dim mechanic As Variant, count As Long
    minimum = 2147483647
    For Each mechanic In totalCounts.Keys
        count = CLng(totalCounts(mechanic))
        If CStr(mechanic) = selectedMechanic Then count = count + 1
        maximum = WorksheetFunction.Max(maximum, count)
        minimum = WorksheetFunction.Min(minimum, count)
    Next mechanic
End Sub

Private Function DaysSinceLastAssignment(ByVal lastDate As Date, ByVal scheduleDate As Date) As Long
    DaysSinceLastAssignment = WorksheetFunction.Max(0, DateDiff("d", lastDate, scheduleDate))
End Function

Private Function CreatesSameDayDuplicate(ByVal mechanic As String, ByVal changingSlot As Long, _
                                         ByRef slots() As ScheduleSlot, ByVal slotCount As Long) As Boolean
    Dim index As Long
    For index = 1 To slotCount
        If index <> changingSlot And slots(index).ScheduleDate = slots(changingSlot).ScheduleDate _
           And slots(index).Mechanic = mechanic Then CreatesSameDayDuplicate = True: Exit Function
    Next index
End Function

Private Sub Increment(ByVal dictionary As Object, ByVal key As String)
    If dictionary.Exists(key) Then dictionary(key) = CLng(dictionary(key)) + 1 Else dictionary.Add key, 1
End Sub

Private Sub AddHistoryToLocationCounts(ByVal locationCounts As Object)
    Dim history As ListObject, row As ListRow, key As String
    Set history = GetTable("HistoryTable")
    For Each row In history.ListRows
        key = Trim$(CStr(row.Range.Cells(1, GetColumnIndex(history, "Mechanic")).Value)) & "|" & _
              Trim$(CStr(row.Range.Cells(1, GetColumnIndex(history, "Location")).Value))
        If locationCounts.Exists(key) Then locationCounts(key) = CLng(locationCounts(key)) + 1
    Next row
End Sub

Private Function ExistingScheduleDates(ByVal schedule As ListObject, ByVal dateIndex As Long) As Object
    Dim result As Object, row As ListRow
    Set result = CreateObject("Scripting.Dictionary")
    For Each row In schedule.ListRows
        If IsDate(row.Range.Cells(1, dateIndex).Value) Then result(DateKey(CDate(row.Range.Cells(1, dateIndex).Value))) = True
    Next row
    Set ExistingScheduleDates = result
End Function

Private Function HistoryKeys(ByVal history As ListObject) As Object
    Dim result As Object, row As ListRow, key As String
    Set result = CreateObject("Scripting.Dictionary")
    For Each row In history.ListRows
        If IsDate(row.Range.Cells(1, GetColumnIndex(history, "Date")).Value) Then
            key = DateKey(CDate(row.Range.Cells(1, GetColumnIndex(history, "Date")).Value)) & "|" & _
                  CStr(row.Range.Cells(1, GetColumnIndex(history, "Location")).Value) & "|" & _
                  CStr(row.Range.Cells(1, GetColumnIndex(history, "Mechanic")).Value)
            result(key) = True
        End If
    Next row
    Set HistoryKeys = result
End Function

Private Function FirstEmptyScheduleRow(ByVal schedule As ListObject, ByVal dateIndex As Long) As ListRow
    Dim row As ListRow
    For Each row In schedule.ListRows
        If IsBlank(row.Range.Cells(1, dateIndex).Value) Then Set FirstEmptyScheduleRow = row: Exit Function
    Next row
End Function

Private Sub SortScheduleByDate(ByVal schedule As ListObject, ByVal dateIndex As Long)
    With schedule.Sort
        .SortFields.Clear
        .SortFields.Add Key:=schedule.ListColumns(dateIndex).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
End Sub

Private Function FirstMatchingDate(ByVal startDate As Date, ByVal requestedDay As Long) As Date
    FirstMatchingDate = DateAdd("d", (requestedDay - Weekday(startDate, vbMonday) + 7) Mod 7, startDate)
End Function

Private Function WeekdayNumber(ByVal weekdayName As String) As Long
    Select Case LCase$(Trim$(weekdayName))
        Case "monday": WeekdayNumber = 1
        Case "tuesday": WeekdayNumber = 2
        Case "wednesday": WeekdayNumber = 3
        Case "thursday": WeekdayNumber = 4
        Case "friday": WeekdayNumber = 5
        Case "saturday": WeekdayNumber = 6
        Case "sunday": WeekdayNumber = 7
        Case Else: Err.Raise vbObjectError + 103, , "Choose a valid Day of Week."
    End Select
End Function

Private Function GetTable(ByVal tableName As String) As ListObject
    Dim sheet As Worksheet
    For Each sheet In ThisWorkbook.Worksheets
        On Error Resume Next
        Set GetTable = sheet.ListObjects(tableName)
        On Error GoTo 0
        If Not GetTable Is Nothing Then Exit Function
    Next sheet
    Err.Raise vbObjectError + 104, , "Workbook table '" & tableName & "' was not found."
End Function

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
End Function

Private Function GetColumnIndex(ByVal table As ListObject, ByVal columnName As String) As Long
    GetColumnIndex = table.ListColumns(columnName).Index
End Function

Private Function GetNamedDate(ByVal nameText As String) As Date
    If Not IsDate(ThisWorkbook.Names(nameText).RefersToRange.Value) Then _
        Err.Raise vbObjectError + 105, , "Enter a valid Start Date."
    GetNamedDate = CDate(ThisWorkbook.Names(nameText).RefersToRange.Value)
End Function

Private Function GetNamedPositiveLong(ByVal nameText As String) As Long
    If Not IsNumeric(ThisWorkbook.Names(nameText).RefersToRange.Value) Then _
        Err.Raise vbObjectError + 106, , "Enter a whole number of weeks."
    GetNamedPositiveLong = CLng(ThisWorkbook.Names(nameText).RefersToRange.Value)
    If GetNamedPositiveLong < 1 Then Err.Raise vbObjectError + 107, , "Number of Weeks must be at least 1."
End Function

Private Function GetNamedText(ByVal nameText As String) As String
    GetNamedText = CStr(ThisWorkbook.Names(nameText).RefersToRange.Value)
End Function

Private Function DateKey(ByVal value As Date) As String
    DateKey = Format$(value, "yyyymmdd")
End Function

Private Function PairKey(ByVal firstMechanic As String, ByVal secondMechanic As String) As String
    If StrComp(firstMechanic, secondMechanic, vbTextCompare) < 0 Then
        PairKey = firstMechanic & "|" & secondMechanic
    Else
        PairKey = secondMechanic & "|" & firstMechanic
    End If
End Function

Private Function IsFutureScheduleDate(ByVal value As Variant) As Boolean
    IsFutureScheduleDate = IsDate(value) And CDate(value) >= Date
End Function

Private Function IsBlank(ByVal value As Variant) As Boolean
    IsBlank = Len(Trim$(CStr(value))) = 0
End Function

' Printing -------------------------------------------------------------------

Private Sub PrepareScheduleForPrinting()
    Dim schedule As ListObject
    Set schedule = GetTable("ScheduleTable")
    With GetSheet("Schedule")
        .PageSetup.PrintArea = .Range(.Range("A1"), schedule.Range.Cells(schedule.Range.Rows.Count, 3)).Address
    End With
    With GetSheet("Schedule").PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .PrintTitleRows = "$1:$10"
        .LeftMargin = Application.InchesToPoints(0.35)
        .RightMargin = Application.InchesToPoints(0.35)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
    End With
End Sub
