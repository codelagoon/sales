Attribute VB_Name = "AssignedLateSchedule"
Option Explicit

' Assigned Late Schedule
' Import this module into the workbook's VBA project.  All addresses are
' resolved through workbook tables and names so the paper layout can move.
' Target platform: Microsoft Excel for Windows (.xlsm).  Mac Excel is not
' supported (Scripting.Dictionary and other COM features are required).

Private Const LOCATION_ONE As String = "1 Centre Street"
Private Const LOCATION_TWO As String = "100 Centre Street"
Private Const PAIR_SEPARATOR As String = " / "

' Tracks the current step so error messages can name exactly where a failure
' occurred (helps diagnose issues without a debugger).
Private mStage As String

Private Type ScheduleSlot
    ScheduleDate As Date
    Location As String
    TableRow As Long
    TableColumn As Long
    Mechanic As String
    Note As String
End Type

' Main supervisor action: build dates, assign every empty future slot, and
' permanently record the resulting assignments.
Public Sub GenerateSchedule()
    Dim verificationLog As String
    If Not RequireWindowsExcel() Then Exit Sub
    On Error GoTo HandleError
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    mStage = "generating schedule dates"
    GenerateDatesInternal
    mStage = "checking active staffing"
    CheckActiveStaffThreshold           ' operational safeguard (aborts if understaffed)
    AssignFutureSchedule
    mStage = "preparing the printout"
    PrepareScheduleForPrinting
    mStage = "updating the fairness summary"
    UpdateFairnessSummary               ' transparency: refresh statistics block
    mStage = "verifying the result"
    verificationLog = VerifyGeneratedSchedule()  ' post-generation work verification

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "The late schedule has been generated." & vbCrLf & vbCrLf & verificationLog, _
           vbInformation, "Assigned Late Schedule"
    Exit Sub
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
HandleError:
    MsgBox "Could not finish while " & mStage & "." & vbCrLf & vbCrLf & _
           Err.Description & " (error " & (Err.Number And &HFFFF&) & ")", _
           vbExclamation, "Schedule not generated"
    Resume CleanExit
End Sub

' Generates only missing schedule dates from the three settings on Schedule.
Public Sub GenerateDates()
    If Not RequireWindowsExcel() Then Exit Sub
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

    If Not RequireWindowsExcel() Then Exit Sub

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
    If Not RequireWindowsExcel() Then Exit Sub
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

' Operational safeguards -----------------------------------------------------

' Pre-execution staffing check. Every empty location cell needs two mechanics,
' so a fully empty date needs four distinct active mechanics who are available
' that day. Aborts with a clear, supervisor-friendly message naming the first
' date that cannot be staffed.
Private Sub CheckActiveStaffThreshold()
    Dim schedule As ListObject, row As ListRow, eligible As Collection
    Dim dateIndex As Long, firstLocationIndex As Long, secondLocationIndex As Long
    Dim scheduleDate As Date, needed As Long, available As Long, candidate As Variant

    Set eligible = ActiveMechanics()
    Set schedule = GetTable("ScheduleTable")
    dateIndex = GetColumnIndex(schedule, "Date")
    firstLocationIndex = GetColumnIndex(schedule, LOCATION_ONE)
    secondLocationIndex = GetColumnIndex(schedule, LOCATION_TWO)

    For Each row In schedule.ListRows
        If IsFutureScheduleDate(row.Range.Cells(1, dateIndex).Value) Then
            scheduleDate = CDate(row.Range.Cells(1, dateIndex).Value)
            needed = 0
            If IsBlank(row.Range.Cells(1, firstLocationIndex).Value) Then needed = needed + 2
            If IsBlank(row.Range.Cells(1, secondLocationIndex).Value) Then needed = needed + 2
            If needed > 0 Then
                available = 0
                For Each candidate In eligible
                    If Not IsUnavailable(CStr(candidate), scheduleDate) Then available = available + 1
                Next candidate
                If available < needed Then
                    Err.Raise vbObjectError + 110, , _
                        "Only " & available & " active, available mechanic(s) on " & _
                        Format$(scheduleDate, "m/d/yyyy") & "; " & needed & " are required. " & _
                        "Add active mechanics or remove availability blocks for that date, then try again."
                End If
            End If
        End If
    Next row
End Sub

' Assignment engine ----------------------------------------------------------

' Builds all future empty slots first, selects across the entire set, then
' performs pair/location/fairness swap passes until no swap improves the plan.
Private Sub AssignFutureSchedule()
    Dim eligible As Collection, slots() As ScheduleSlot
    Dim slotCount As Long
    Dim totalCounts As Object, recentCounts As Object, locationCounts As Object
    Dim lastAssigned As Object, pairCounts As Object
    Dim historySheet As Worksheet, historyWasHidden As Boolean
    Dim baseLocationCounts As Object

    mStage = "assigning: reading the mechanics list"
    Set eligible = ActiveMechanics()
    If eligible.Count < 4 Then Err.Raise vbObjectError + 101, , _
        "At least four active mechanics are required (two locations x two mechanics per schedule date)."

    mStage = "assigning: collecting open shifts"
    slotCount = CollectEmptyFutureSlots(slots)
    If slotCount = 0 Then Exit Sub

    mStage = "assigning: reading history statistics"
    Set historySheet = GetTable("HistoryTable").Parent
    historyWasHidden = (historySheet.Visible <> xlSheetVisible)
    If historyWasHidden Then historySheet.Visible = xlSheetVisible

    On Error GoTo HistoryReadFailed
    Call LoadHistoryStatistics(eligible, DateAdd("m", -3, Date), totalCounts, recentCounts, _
                               locationCounts, lastAssigned, pairCounts)
    On Error GoTo 0

    If historyWasHidden Then historySheet.Visible = xlSheetHidden

    Set baseLocationCounts = CopyDictionary(locationCounts)

    mStage = "assigning: selecting the fairest mechanics"
    AssignSlotsGlobally slots, slotCount, eligible, totalCounts, recentCounts, _
                        locationCounts, lastAssigned, pairCounts
    mStage = "assigning: balancing the plan"
    ImproveSchedule slots, slotCount, baseLocationCounts
    mStage = "assigning: writing the schedule"
    WriteSlotsToSchedule slots, slotCount
    mStage = "assigning: recording history"
    CommitAssignmentToHistory slots, slotCount
    Exit Sub

HistoryReadFailed:
    If historyWasHidden Then historySheet.Visible = xlSheetHidden
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' Selects the next assignment from every remaining date/location rather than
' treating each schedule row independently.
Private Sub AssignSlotsGlobally(ByRef slots() As ScheduleSlot, ByVal slotCount As Long, _
                                ByVal eligible As Collection, ByVal totalCounts As Object, _
                                ByVal recentCounts As Object, ByVal locationCounts As Object, _
                                ByVal lastAssigned As Object, ByVal pairCounts As Object)
    Dim assigned As Long, bestSlot As Long, bestMechanic As String
    Dim candidateSlot As Long, candidate As Variant, score As Double, bestScore As Double
    Dim targetDate As Date

    For assigned = 1 To slotCount
        bestScore = 1E+30
        bestSlot = 0
        bestMechanic = vbNullString
        targetDate = NextUnassignedDate(slots, slotCount)
        For candidateSlot = 1 To slotCount
            If Len(slots(candidateSlot).Mechanic) = 0 And slots(candidateSlot).ScheduleDate = targetDate Then
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

        ' Explicit duplicate-assignment guard: never place a mechanic at both
        ' buildings on the same date (scoring already avoids it; this enforces).
        If WouldDuplicateOnDate(bestMechanic, bestSlot, slots, slotCount) Then _
            Err.Raise vbObjectError + 111, , _
                bestMechanic & " would be assigned twice on " & _
                Format$(slots(bestSlot).ScheduleDate, "m/d/yyyy") & "."

        slots(bestSlot).Note = BuildSelectionNote(bestMechanic, slots(bestSlot), totalCounts, _
                                                  recentCounts, locationCounts, lastAssigned, _
                                                  pairCounts, slots, slotCount)
        slots(bestSlot).Mechanic = bestMechanic
        ApplyTemporaryAssignment bestMechanic, slots(bestSlot), totalCounts, recentCounts, _
                                 locationCounts, lastAssigned, pairCounts, slots, slotCount
    Next assigned
End Sub

' Explicit same-date duplicate guard used inside the assignment loop.
Private Function WouldDuplicateOnDate(ByVal mechanic As String, ByVal targetSlot As Long, _
                                      ByRef slots() As ScheduleSlot, ByVal slotCount As Long) As Boolean
    Dim index As Long
    For index = 1 To slotCount
        If index <> targetSlot And slots(index).ScheduleDate = slots(targetSlot).ScheduleDate _
           And slots(index).Mechanic = mechanic Then
            WouldDuplicateOnDate = True
            Exit Function
        End If
    Next index
End Function

' Builds the short audit tag stored with each history row explaining why this
' mechanic won the slot (lowest lifetime/recent load, idle time, pairing
' rotation status, and location balance).
Private Function BuildSelectionNote(ByVal mechanic As String, ByRef slot As ScheduleSlot, _
                                    ByVal totalCounts As Object, ByVal recentCounts As Object, _
                                    ByVal locationCounts As Object, ByVal lastAssigned As Object, _
                                    ByVal pairCounts As Object, ByRef slots() As ScheduleSlot, _
                                    ByVal slotCount As Long) As String
    Dim lifetime As Long, recent As Long, idle As String, pairText As String
    Dim partner As String, repeats As Long, locationGap As Long, daysSince As Long
    lifetime = CLng(totalCounts(mechanic))
    recent = CLng(recentCounts(mechanic))
    If CDate(lastAssigned(mechanic)) <= DateSerial(1900, 1, 1) Then
        idle = "new"
    Else
        daysSince = WorksheetFunction.Max(0, DateDiff("d", CDate(lastAssigned(mechanic)), slot.ScheduleDate))
        idle = daysSince & "d"
    End If
    partner = AssignedPartner(slot, slots, slotCount)
    If Len(partner) > 0 Then
        repeats = 0
        If pairCounts.Exists(PairKey(mechanic, partner)) Then repeats = CLng(pairCounts(PairKey(mechanic, partner)))
        If repeats = 0 Then
            pairText = "new pairing w/ " & partner
        Else
            pairText = "paired w/ " & partner & " x" & repeats
        End If
    Else
        pairText = "pairing pending"
    End If
    locationGap = LocationImbalanceAfter(mechanic, slot.Location, locationCounts)
    BuildSelectionNote = "lifetime " & lifetime & ", recent " & recent & ", idle " & idle & _
                         ", " & pairText & ", loc-gap " & locationGap
End Function

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
    If Len(partner) > 0 Then Call Increment(pairCounts, PairKey(mechanic, partner))
End Sub

' Balancing pass: swaps two provisional assignments only when the full fairness
' objective decreases, repeatedly, until a complete pass finds no improvement.
Private Sub ImproveSchedule(ByRef slots() As ScheduleSlot, ByVal slotCount As Long, _
                            ByVal baseLocations As Object)
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
                        beforeScore = LocalFairnessScore(slots, slotCount, baseLocations)
                        slots(firstSlot).Mechanic = secondMechanic
                        slots(secondSlot).Mechanic = firstMechanic
                        afterScore = LocalFairnessScore(slots, slotCount, baseLocations)
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
' location imbalance. Location imbalance is measured against lifetime history so
' the balancing passes rotate each mechanic between the two buildings over time.
Private Function LocalFairnessScore(ByRef slots() As ScheduleSlot, ByVal slotCount As Long, _
                                    ByVal baseLocations As Object) As Double
    Dim totals As Object, locations As Object, pairs As Object, seen As Object, index As Long
    Dim maxCount As Long, minCount As Long, mechanic As Variant, score As Double
    Dim key As Variant, oneCount As Long, twoCount As Long, personName As String
    Set totals = NewDictionary()
    Set locations = NewDictionary()
    Set pairs = NewDictionary()
    Set seen = NewDictionary()
    For Each key In baseLocations.Keys
        locations(key) = CLng(baseLocations(key))
    Next key
    For index = 1 To slotCount
        Call Increment(totals, slots(index).Mechanic)
        Call Increment(locations, slots(index).Mechanic & "|" & slots(index).Location)
        If Len(AssignedPartner(slots(index), slots, slotCount)) > 0 Then _
            Call Increment(pairs, PairKey(slots(index).Mechanic, AssignedPartner(slots(index), slots, slotCount)))
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
    For Each key In locations.Keys
        personName = Left$(CStr(key), InStrRev(CStr(key), "|") - 1)
        If Not seen.Exists(personName) Then
            seen.Add personName, True
            oneCount = 0
            twoCount = 0
            If locations.Exists(personName & "|" & LOCATION_ONE) Then oneCount = CLng(locations(personName & "|" & LOCATION_ONE))
            If locations.Exists(personName & "|" & LOCATION_TWO) Then twoCount = CLng(locations(personName & "|" & LOCATION_TWO))
            score = score + Abs(oneCount - twoCount) * 10#
        End If
    Next key
    LocalFairnessScore = score + (maxCount - minCount) * 1000000000#
End Function

' Writes the two people for each location in the paper-compatible "Name / Name"
' format. It never changes populated schedule cells.
Private Sub WriteSlotsToSchedule(ByRef slots() As ScheduleSlot, ByVal slotCount As Long)
    Dim schedule As ListObject, index As Long, scheduleRow As ListRow
    Dim grouped As Object, key As Variant, cellValue As String
    Set schedule = GetTable("ScheduleTable")
    Set grouped = NewDictionary()
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

' Adds every newly generated assignment to the permanent History table, together
' with an audit note recording why the mechanic was selected. The duplicate key
' makes rerunning Generate Schedule safe.
Private Sub CommitAssignmentToHistory(ByRef slots() As ScheduleSlot, ByVal slotCount As Long)
    Dim history As ListObject, existing As Object, index As Long, historyRow As ListRow
    Dim key As String, historySheet As Worksheet, priorVisibility As Long
    Set history = GetTable("HistoryTable")

    ' Adding table rows is most reliable when the sheet is visible; the History
    ' sheet is normally hidden, so make it visible for the write and restore it.
    Set historySheet = history.Parent
    priorVisibility = historySheet.Visible
    If priorVisibility <> xlSheetVisible Then historySheet.Visible = xlSheetVisible

    Set existing = HistoryKeys(history)
    For index = 1 To slotCount
        key = DateKey(slots(index).ScheduleDate) & "|" & slots(index).Location & "|" & slots(index).Mechanic
        If Not existing.Exists(key) Then
            Set historyRow = history.ListRows.Add
            historyRow.Range.Cells(1, GetColumnIndex(history, "Date")).Value = slots(index).ScheduleDate
            historyRow.Range.Cells(1, GetColumnIndex(history, "Location")).Value = slots(index).Location
            historyRow.Range.Cells(1, GetColumnIndex(history, "Mechanic")).Value = slots(index).Mechanic
            historyRow.Range.Cells(1, GetColumnIndex(history, "Assignment Created On")).Value = Now
            historyRow.Range.Cells(1, GetColumnIndex(history, "Selection Note")).Value = slots(index).Note
            existing.Add key, True
        End If
    Next index

    If priorVisibility <> xlSheetVisible Then historySheet.Visible = priorVisibility
End Sub

' Statistics -----------------------------------------------------------------

' Reads History once and fills every statistic dictionary the assignment engine
' needs. A single pass avoids repeated table access and name-collision bugs.
Private Sub LoadHistoryStatistics(ByVal mechanics As Collection, ByVal recentCutoff As Date, _
                                  ByRef outTotal As Object, ByRef outRecent As Object, _
                                  ByRef outLocations As Object, ByRef outLastAssigned As Object, _
                                  ByRef outPairs As Object)
    Dim history As ListObject, row As ListRow, mechanic As Variant
    Dim dateCol As Long, locationCol As Long, mechanicCol As Long
    Dim name As String, scheduled As Date, groupKey As String, locationText As String
    Dim pairGroups As Object, pairGroupKey As Variant
    Dim peopleText As String, people() As String

    Set outTotal = NewDictionary()
    Set outRecent = NewDictionary()
    Set outLocations = NewDictionary()
    Set outLastAssigned = NewDictionary()
    Set outPairs = NewDictionary()
    Set pairGroups = NewDictionary()

    For Each mechanic In mechanics
        name = CStr(mechanic)
        outTotal.Add name, 0
        outRecent.Add name, 0
        outLocations.Add name & "|" & LOCATION_ONE, 0
        outLocations.Add name & "|" & LOCATION_TWO, 0
        outLastAssigned.Add name, DateSerial(1900, 1, 1)
    Next mechanic

    Set history = GetTable("HistoryTable")
    dateCol = GetColumnIndex(history, "Date")
    locationCol = GetColumnIndex(history, "Location")
    mechanicCol = GetColumnIndex(history, "Mechanic")

    For Each row In history.ListRows
        If Not IsDateValue(row.Range.Cells(1, dateCol).Value) Then GoTo ContinueHistoryRow
        scheduled = CDate(row.Range.Cells(1, dateCol).Value)
        name = TableCellText(row, mechanicCol)
        If Len(name) = 0 Or Not outTotal.Exists(name) Then GoTo ContinueHistoryRow

        outTotal(name) = CLng(outTotal(name)) + 1
        If scheduled >= recentCutoff Then outRecent(name) = CLng(outRecent(name)) + 1

        locationText = TableCellText(row, locationCol)
        groupKey = name & "|" & locationText
        If outLocations.Exists(groupKey) Then outLocations(groupKey) = CLng(outLocations(groupKey)) + 1

        If scheduled > CDate(outLastAssigned(name)) Then outLastAssigned(name) = scheduled

        groupKey = DateKey(scheduled) & "|" & locationText
        If pairGroups.Exists(groupKey) Then
            pairGroups(groupKey) = CStr(pairGroups(groupKey)) & "|" & name
        Else
            pairGroups.Add groupKey, name
        End If
ContinueHistoryRow:
    Next row

    For Each pairGroupKey In pairGroups.Keys
        peopleText = CStr(pairGroups(pairGroupKey))
        people = Split(peopleText, "|")
        If UBound(people) - LBound(people) + 1 = 2 Then _
            Call Increment(outPairs, PairKey(Trim$(people(LBound(people))), Trim$(people(UBound(people)))))
    Next pairGroupKey
End Sub

Private Function BuildLifetimeCounts(ByVal mechanics As Collection) As Object
    Dim total As Object, recent As Object, locations As Object
    Dim lastAssigned As Object, pairs As Object
    Call LoadHistoryStatistics(mechanics, DateSerial(1900, 1, 1), total, recent, locations, lastAssigned, pairs)
    Set BuildLifetimeCounts = total
End Function

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

' Helpers --------------------------------------------------------------------

' Mac Excel does not provide Scripting.Dictionary and related COM support.
Private Function RequireWindowsExcel() As Boolean
    If InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0 Then
        MsgBox "This workbook requires Microsoft Excel for Windows." & vbCrLf & vbCrLf & _
               "Excel for Mac is not supported. The scheduling engine relies on " & _
               "Windows-only VBA features (for example Scripting.Dictionary).", _
               vbExclamation, "Windows Excel required"
        RequireWindowsExcel = False
    Else
        RequireWindowsExcel = True
    End If
End Function

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

' Initial allocation follows chronological schedule dates so "time since last
' assignment" is meaningful. The subsequent ImproveSchedule pass evaluates the
' completed plan across every generated date.
Private Function NextUnassignedDate(ByRef slots() As ScheduleSlot, ByVal slotCount As Long) As Date
    Dim index As Long, earliest As Date
    For index = 1 To slotCount
        If Len(slots(index).Mechanic) = 0 Then
            If earliest = 0 Or slots(index).ScheduleDate < earliest Then earliest = slots(index).ScheduleDate
        End If
    Next index
    NextUnassignedDate = earliest
End Function

Private Function IsUnavailable(ByVal mechanic As String, ByVal scheduleDate As Date) As Boolean
    Dim availability As ListObject, row As ListRow
    Set availability = GetTable("AvailabilityTable")
    For Each row In availability.ListRows
        If StrComp(Trim$(CStr(row.Range.Cells(1, GetColumnIndex(availability, "Mechanic")).Value)), mechanic, vbTextCompare) = 0 _
           And IsDateValue(row.Range.Cells(1, GetColumnIndex(availability, "Date")).Value) Then
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

Private Function NewDictionary() As Object
    Set NewDictionary = CreateObject("Scripting.Dictionary")
    If NewDictionary Is Nothing Then _
        Err.Raise vbObjectError + 112, , "Scripting.Dictionary is not available. Use Microsoft Excel for Windows."
End Function

Private Function CopyDictionary(ByVal source As Object) As Object
    Dim result As Object, key As Variant
    Set result = NewDictionary()
    For Each key In source.Keys
        result(CStr(key)) = source(key)
    Next key
    Set CopyDictionary = result
End Function

Private Function IsDateValue(ByVal value As Variant) As Boolean
    IsDateValue = (Not IsEmpty(value)) And (Not IsNull(value)) And IsDate(value)
End Function

Private Function TableCellText(ByVal row As ListRow, ByVal columnIndex As Long) As String
    TableCellText = Trim$(CStr(row.Range.Cells(1, columnIndex).Value))
End Function

Private Function HistoryKeys(ByVal history As ListObject) As Object
    Dim result As Object, row As ListRow, key As String
    Dim dateCol As Long, locationCol As Long, mechanicCol As Long
    Set result = NewDictionary()
    dateCol = GetColumnIndex(history, "Date")
    locationCol = GetColumnIndex(history, "Location")
    mechanicCol = GetColumnIndex(history, "Mechanic")
    For Each row In history.ListRows
        If IsDateValue(row.Range.Cells(1, dateCol).Value) Then
            key = DateKey(CDate(row.Range.Cells(1, dateCol).Value)) & "|" & _
                  TableCellText(row, locationCol) & "|" & TableCellText(row, mechanicCol)
            result(key) = True
        End If
    Next row
    Set HistoryKeys = result
End Function

Private Function ExistingScheduleDates(ByVal schedule As ListObject, ByVal dateIndex As Long) As Object
    Dim result As Object, row As ListRow
    Set result = NewDictionary()
    For Each row In schedule.ListRows
        If IsDateValue(row.Range.Cells(1, dateIndex).Value) Then _
            result(DateKey(CDate(row.Range.Cells(1, dateIndex).Value))) = True
    Next row
    Set ExistingScheduleDates = result
End Function

Private Function FirstEmptyScheduleRow(ByVal schedule As ListObject, ByVal dateIndex As Long) As ListRow
    Dim row As ListRow
    For Each row In schedule.ListRows
        If IsBlank(row.Range.Cells(1, dateIndex).Value) Then Set FirstEmptyScheduleRow = row: Exit Function
    Next row
End Function

Private Sub SortScheduleByDate(ByVal schedule As ListObject, ByVal dateIndex As Long)
    If schedule.ListRows Is Nothing Then Exit Sub
    If schedule.ListRows.Count < 2 Then Exit Sub
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
    Dim sheet As Worksheet, listObj As ListObject
    For Each sheet In ThisWorkbook.Worksheets
        For Each listObj In sheet.ListObjects
            If StrComp(listObj.Name, tableName, vbTextCompare) = 0 Then
                Set GetTable = listObj
                Exit Function
            End If
        Next listObj
    Next sheet
    Err.Raise vbObjectError + 104, , "Workbook table '" & tableName & "' was not found."
End Function

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
End Function

Private Function GetColumnIndex(ByVal table As ListObject, ByVal columnName As String) As Long
    GetColumnIndex = table.ListColumns(columnName).Index
End Function

' Resolves a workbook-level named cell (the three settings) to a Range and
' raises a clear, named error if the name is missing, instead of a cryptic
' "object variable not set" failure.
Private Function SettingRange(ByVal nameText As String) As Range
    On Error GoTo Fail
    Set SettingRange = ThisWorkbook.Names(nameText).RefersToRange
    If SettingRange Is Nothing Then GoTo Fail
    Exit Function
Fail:
    Err.Raise vbObjectError + 108, , "The workbook setting '" & nameText & "' could not be found. " & _
        "Rebuild the workbook with build_workbook.py."
End Function

Private Function GetNamedDate(ByVal nameText As String) As Date
    If Not IsDate(SettingRange(nameText).Value) Then _
        Err.Raise vbObjectError + 105, , "Enter a valid Start Date."
    GetNamedDate = CDate(SettingRange(nameText).Value)
End Function

Private Function GetNamedPositiveLong(ByVal nameText As String) As Long
    If Not IsNumeric(SettingRange(nameText).Value) Then _
        Err.Raise vbObjectError + 106, , "Enter a whole number of weeks."
    GetNamedPositiveLong = CLng(SettingRange(nameText).Value)
    If GetNamedPositiveLong < 1 Then Err.Raise vbObjectError + 107, , "Number of Weeks must be at least 1."
End Function

Private Function GetNamedText(ByVal nameText As String) As String
    GetNamedText = CStr(SettingRange(nameText).Value)
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

' Transparency & verification ------------------------------------------------

' Refreshes the Fairness Summary block on the Schedule sheet with a statistical
' analysis of lifetime workload spread across active mechanics, read live from
' History. Standard deviation and the max-to-min gap quantify fairness.
Private Sub UpdateFairnessSummary()
    Dim eligible As Collection, counts As Object, mechanic As Variant
    Dim total As Long, maxCount As Long, minCount As Long, n As Long
    Dim mean As Double, sumSq As Double, stdev As Double, gap As Long

    Set eligible = ActiveMechanics()
    Set counts = BuildLifetimeCounts(eligible)
    n = 0
    total = 0
    maxCount = 0
    minCount = 2147483647
    For Each mechanic In counts.Keys
        n = n + 1
        total = total + CLng(counts(mechanic))
        maxCount = WorksheetFunction.Max(maxCount, CLng(counts(mechanic)))
        minCount = WorksheetFunction.Min(minCount, CLng(counts(mechanic)))
    Next mechanic

    If n = 0 Then
        gap = 0
        stdev = 0
    Else
        gap = maxCount - minCount
        mean = total / n
        sumSq = 0
        For Each mechanic In counts.Keys
            sumSq = sumSq + (CLng(counts(mechanic)) - mean) ^ 2
        Next mechanic
        stdev = Sqr(sumSq / n)
    End If

    SetNamedValue "FairnessActiveCount", n
    SetNamedValue "FairnessTotal", total
    SetNamedValue "FairnessGap", gap
    SetNamedValue "FairnessStdDev", Round(stdev, 3)
End Sub

Private Sub SetNamedValue(ByVal nameText As String, ByVal value As Variant)
    On Error Resume Next
    ThisWorkbook.Names(nameText).RefersToRange.Value = value
    On Error GoTo 0
End Sub

' Post-generation work verification. Confirms, over every future row: (1) zero
' duplicate mechanics on the same date, (2) zero inactive/unavailable mechanics
' scheduled, and (3) all future rows filled. Returns a human-readable log.
Private Function VerifyGeneratedSchedule() As String
    Dim schedule As ListObject, row As ListRow, eligible As Collection
    Dim dateIndex As Long, firstLocationIndex As Long, secondLocationIndex As Long
    Dim activeSet As Object, mechanic As Variant, peopleToday As Object
    Dim duplicateDates As Long, illegalCount As Long, unfilledCount As Long
    Dim scheduleDate As Date, colIndex As Variant, cellValue As String
    Dim parts() As String, i As Long, name As String

    Set eligible = ActiveMechanics()
    Set activeSet = NewDictionary()
    For Each mechanic In eligible
        activeSet(CStr(mechanic)) = True
    Next mechanic

    Set schedule = GetTable("ScheduleTable")
    dateIndex = GetColumnIndex(schedule, "Date")
    firstLocationIndex = GetColumnIndex(schedule, LOCATION_ONE)
    secondLocationIndex = GetColumnIndex(schedule, LOCATION_TWO)

    For Each row In schedule.ListRows
        If IsFutureScheduleDate(row.Range.Cells(1, dateIndex).Value) Then
            scheduleDate = CDate(row.Range.Cells(1, dateIndex).Value)
            Set peopleToday = NewDictionary()
            For Each colIndex In Array(firstLocationIndex, secondLocationIndex)
                cellValue = Trim$(CStr(row.Range.Cells(1, CLng(colIndex)).Value))
                If Len(cellValue) = 0 Then
                    unfilledCount = unfilledCount + 1
                Else
                    parts = Split(cellValue, PAIR_SEPARATOR)
                    For i = LBound(parts) To UBound(parts)
                        name = Trim$(parts(i))
                        If Len(name) > 0 Then
                            If peopleToday.Exists(name) Then
                                peopleToday(name) = CLng(peopleToday(name)) + 1
                            Else
                                peopleToday.Add name, 1
                            End If
                            If Not activeSet.Exists(name) Then
                                illegalCount = illegalCount + 1
                            ElseIf IsUnavailable(name, scheduleDate) Then
                                illegalCount = illegalCount + 1
                            End If
                        End If
                    Next i
                End If
            Next colIndex
            For Each mechanic In peopleToday.Keys
                If CLng(peopleToday(mechanic)) > 1 Then
                    duplicateDates = duplicateDates + 1
                    Exit For
                End If
            Next mechanic
        End If
    Next row

    VerifyGeneratedSchedule = "Verification results:" & vbCrLf & _
        "  1. No duplicate mechanic on a date: " & PassFail(duplicateDates) & vbCrLf & _
        "  2. No inactive/unavailable scheduled: " & PassFail(illegalCount) & vbCrLf & _
        "  3. All future rows filled: " & PassFail(unfilledCount)
End Function

Private Function PassFail(ByVal violations As Long) As String
    If violations = 0 Then
        PassFail = "PASS"
    Else
        PassFail = "FAIL (" & violations & ")"
    End If
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
