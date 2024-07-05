' SHEETS REQUIRED:
' 1. PastUnits [Name, Unit]
' 2. ModuleGrades [Name, Module, Grade]
' 3. UnitRequirements [Unit, Module Preference (CSV format), Quota]
' 4. NewAssignments [Name, New Unit]

Sub AssignUnits()
    ' DECLARATIONS:
    ' Worksheets
    Dim wsNewAssgn As Worksheet
    Dim wsPastUnits As Worksheet
    Dim wsMods As Worksheet
    Dim wsUnitReq As Worksheet
    
    ' Last Worksheet Rows
    Dim lastRowNewAssgn As Long
    Dim lastRowPastUnits As Long
    Dim lastRowMods As Long
    Dim lastRowUnitReq As Long
    
    ' Looping Variables
    Dim i As Long, j As Long
    
    ' Required Dictionaries
    Dim pastUnitsDict As Object
    Dim newAssgnDict As Object
    Dim modsTakenDict As Object
    Dim availUnitsDict As Object
    Dim unitReqDict As Object
    Dim unitQuotaDict As Object
    
    Dim bestPossUnits As Object
    Dim yetToGoUnits As Object
    
    
    ' SETTING VARIABLES:
    ' Set worksheets
    Set wsNewAssgn = ThisWorkbook.Sheets("NewAssignments")
    Set wsPastUnits = ThisWorkbook.Sheets("PastUnits")
    Set wsMods = ThisWorkbook.Sheets("ModuleGrades")
    Set wsUnitReq = ThisWorkbook.Sheets("UnitRequirements")
    
    ' Set Dictionaries
    Set pastUnitsDict = CreateObject("Scripting.Dictionary")
    Set newAssgnDict = CreateObject("Scripting.Dictionary")
    Set modsTakenDict = CreateObject("Scripting.Dictionary")
    Set availUnitsDict = CreateObject("Scripting.Dictionary")
    Set unitReqDict = CreateObject("Scripting.Dictionary")
    Set unitQuotaDict = CreateObject("Scripting.Dictionary")
    Set bestPossUnits = CreateObject("Scripting.Dictionary")
    Set yetToGoUnits = CreateObject("Scripting.Dictionary")
    
    ' Set last worksheet rows
    lastRowNewAssgn = wsNewAssgn.Cells(wsNewAssgn.Rows.Count, "A").End(xlUp).Row
    lastRowPastUnits = wsPastUnits.Cells(wsPastUnits.Rows.Count, "A").End(xlUp).Row
    lastRowMods = wsMods.Cells(wsMods.Rows.Count, "A").End(xlUp).Row
    lastRowUnitReq = wsUnitReq.Cells(wsUnitReq.Rows.Count, "A").End(xlUp).Row
    
    
    ' INITIALIZING DICTIONARIES:
    ' Populate newAssgnDict with names from NewAssgn worksheet
    For Each cell In wsNewAssgn.Range("A2:A" & lastRowNewAssgn)
        newAssgnDict.Add cell.Value, "0"
        ' RESULTANT FORMAT = "Name, 0"
    Next cell
    
    ' Populate pastUnitsDict
    Dim unitName As String
    Dim personName As String
    Dim unitArr As String
    For Each cell In wsPastUnits.Range("A2:A" & lastRowPastUnits) ' Assume names are in Col A
        personName = cell.Value
        unitName = cell.Offset(0, 1).Value ' Gets the value in the column to the right
        If Not pastUnitsDict.exists(personName) Then
            ' Create a new person key, or add the unit to their name if they already exist in the dict
            pastUnitsDict(personName) = unitName
        Else
            ' If the person already exists in the dict, then add the new past unit to their existing list
            unitArr = pastUnitsDict(personName)
            unitArr = unitArr & "," & unitName
            pastUnitsDict(personName) = unitArr
            ' RESULTANT FORMAT = "Name: PastUnit1,PastUnit2,PatUnit3"
            ' If a person has no past units, then they will not be in this dictionary
        End If
    Next cell
    
    ' Populate unitReq and unitQuota dictionaries
    Dim unitReq As String
    Dim unitQuota As Long
    For Each cell In wsUnitReq.Range("A2:A" & lastRowUnitReq)
        unitName = cell.Value
        unitReq = cell.Offset(0, 1).Value
        unitQuota = cell.Offset(0, 2).Value
        ' Add keys and values to unitReq and unitQuota dictionaries
        unitReqDict.Add unitName, unitReq ' FORMAT = "unitName: mod1,mod2,..."
        unitQuotaDict.Add unitName, unitQuota ' FORMAT = "unitName: quota"
    Next cell
    
    ' Populate modsTaken dictionary
    Dim moduleTaken As String
    Dim modsTaken As String
    Dim modsArr() As String
    For Each cell In wsMods.Range("A2:A" & lastRowMods)
        personName = cell.Value
        moduleTaken = cell.Offset(0, 1).Value
        If Not modsTakenDict.exists(personName) Then
            ' If the person has not yet been logged in the dictionary, add their name
            modsTakenDict(personName) = moduleTaken
        Else
            ' For those who already have entries in the dictionary, add on to their list of mods
            modsTaken = modsTakenDict(personName)
            modsTaken = modsTaken & "," & moduleTaken
            modsTakenDict(personName) = modsTaken
            ' RESULTANT FORMAT = "Name: module1,module2,module3..."
        End If
    Next cell
    
    
    ' FINDING POSSIBLE UNITS FOR EACH PERSON:
    ' Create dictionaries that best match people's modules taken
    Dim name As Variant, unit As Variant, currUnit As Variant
    Dim unitMod As Variant, personMod As Variant
    Dim possibleUnits As String
    Dim pastUnits As String
    Dim pastUnitsArr() As String
    Dim goodUnit As Boolean
    Dim personMods() As String
    Dim unitMods() As String
    Dim persMod As Variant
    
    ' Populating "bestPossUnits"
    For Each name In newAssgnDict.keys
        possibleUnits = "0"
        personMods = Split(modsTakenDict(name), ",")
        For Each unit In unitReqDict.keys
            goodUnit = False
            unitMods = Split(unitReqDict(unit), ",")
            For Each unitMod In unitMods
                For Each persMod In personMods
                    If LCase(persMod) Like LCase("*" & unitMod & "*") Then
                        goodUnit = True
                        ' Exit persMod loop
                        Exit For
                    End If
                Next persMod
                ' Exit unitMod loop
                If goodUnit Then
                    Exit For
                End If
            Next unitMod
            ' Add goodUnit to possibleUnits
            If goodUnit Then
                If possibleUnits = "0" Then
                    possibleUnits = unit
                Else
                    possibleUnits = possibleUnits & "," & unit
                End If
            End If
        Next unit
        If Not possibleUnits = "0" Then
            bestPossUnits.Add name, possibleUnits
        End If
        ' RESULTANT FORMAT = "Name: possibleUnit1,possibleUnit2,..."
    Next name
    
    ' Populating "yetToGoUnits"
    For Each name In newAssgnDict.keys
        possibleUnits = "0"
        If pastUnitsDict.exists(name) Then
            pastUnitsArr = Split(pastUnitsDict(name), ",")
            For Each unit In unitReqDict.keys
                If Not IsInArray(pastUnitsArr, CStr(unit)) Then
                    If possibleUnits = "0" Then
                        possibleUnits = unit
                    Else
                        possibleUnits = possibleUnits & "," & unit
                    End If
                End If
            Next unit
        Else
            For Each unit In unitReqDict.keys
                If possibleUnits = "0" Then
                    possibleUnits = unit
                Else
                    possibleUnits = possibleUnits & "," & unit
                End If
            Next unit
        End If
        yetToGoUnits.Add name, possibleUnits
        ' RESULTANT FORMAT = "Name: possibleUnit1,possibleUnit2,..."
    Next name
    
    
    ' UNIT ALLOCATION:
    Dim prefUnits() As String
    Dim maxUnit As String
    Dim maxQuota As Long
    Dim remNum As Long
    Dim qwerty As String, temp As String, yes() As String
    Dim length As Long
    Dim currUnits() As String, cfmUnit As String, tmp As String
    
    ' Assign units to those with matching modules
    For Each name In bestPossUnits.keys
        maxQuota = 0
        maxUnit = "poo"
        prefUnits = Split(bestPossUnits(name), ",")
        ' Loop through the units that match a person's mods taken
        For Each unit In prefUnits
                currUnits = Split(yetToGoUnits(name), ",")
                ' Check that the selected unit has yet to be gone to by the person
                If IsInArray(currUnits, CStr(unit)) Then
                    If maxQuota <= unitQuotaDict(unit) And unitQuotaDict(unit) > 0 Then
                        maxUnit = unit
                    End If
                End If
        Next unit
        If Not maxUnit = "poo" Then
            newAssgnDict(name) = maxUnit
            unitQuotaDict(maxUnit) = unitQuotaDict(maxUnit) - 1
        End If
    Next name
    
    ' Assign units to other people
    For Each name In yetToGoUnits.keys
        If newAssgnDict(name) = "0" Then
            maxQuota = 0
            maxUnit = "poo"
            prefUnits = Split(yetToGoUnits(name), ",")
            For Each unit In prefUnits
                If maxQuota <= unitQuotaDict(unit) And unitQuotaDict(unit) > 0 Then
                    maxUnit = unit
                End If
            Next unit
            If Not maxUnit = "poo" Then
                newAssgnDict(name) = maxUnit
                unitQuotaDict(maxUnit) = unitQuotaDict(maxUnit) - 1
            End If
        End If
    Next name
    
    ' Assign units randomly to the remaining people who still do not have units
    Dim nameList() As String
    Dim names As String
    Dim poopy As Variant
    Dim randomIndex As Long
    names = "l"
    For Each name In newAssgnDict.keys
        If newAssgnDict(name) = "0" Then
            If names = "l" Then
                names = name
            Else
                names = names & "," & name
            End If
        End If
    Next name
    nameList = Split(names, ",")
    Do While UBound(nameList) >= LBound(nameList)
        randomIndex = Int((UBound(nameList) - LBound(nameList) + 1) * Rnd + LBound(nameList))
        poopy = nameList(randomIndex)
        If randomIndex < UBound(nameList) Then
            nameList(randomIndex) = nameList(UBound(nameList))
        End If
        
        For Each unit In unitQuotaDict.keys
            maxQuota = 0
            maxUnit = "poo"
            If unitQuotaDict(unit) >= maxQuota And unitQuotaDict(unit) > 1 Then
                maxQuota = unitQuotaDict(unit)
                maxUnit = unit
            End If
        Next unit
        If Not maxUnit = "poo" Then
            newAssgnDict(poopy) = maxUnit
        Else
            newAssgnDict(poopy) = "UNASSIGNED"
        End If
        
        If UBound(nameList) > LBound(nameList) Then
            ReDim Preserve nameList(LBound(nameList) To UBound(nameList) - 1)
        Else
            Exit Do
        End If
    Loop
    
       
    ' OUTPUT:
    ' Print results out to NewAssignments sheet
    Dim cellLooper As Variant
    For Each cellLooper In wsNewAssgn.Range("A2:A" & lastRowNewAssgn)
        cellLooper.Offset(0, 1).Value = newAssgnDict(Trim(cellLooper.Value))
        cellLooper.Offset(0, 2).Value = Replace(modsTakenDict(Trim(cellLooper.Value)), ",", ", ")
    Next cellLooper
    
    
    ' CLEAN UP
    Set newAssgnDict = Nothing
    Set pastUnitsDict = Nothing
    Set newAssgnDict = Nothing
    Set modsTakenDict = Nothing
    Set availUnitsDict = Nothing
    Set unitReqDict = Nothing
    Set unitQuotaDict = Nothing
    Set bestPossUnits = Nothing
    Set yetToGoUnits = Nothing
    Set wsNewAssgn = Nothing
    Set wsPastUnits = Nothing
    Set wsMods = Nothing
    Set wsUnitReq = Nothing
End Sub


Function IsInArray(arr() As String, target As String):
    Dim x As Variant
    For Each x In arr
        If target = x Then
            IsInArray = True
            Exit Function
        End If
    Next x
    IsInArray = False
End Function
