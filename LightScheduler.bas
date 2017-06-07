Attribute VB_Name = "LightScheduler"
Option Explicit
Const NUMBER_OF_COLUMNS As Integer = 10
Const TEXT_FILE_DELIMITER As String = " "
Const DEFAULT_FILE_NAME As String = "Data.txt"
Const COLUMN_NAMES As String = "DATE,HOURS,MINUTES,SECONDS,UV%,DB%,BL%,GR%,RE%,IR%"
Const FILE_NAME_CELL_NUMBER As Integer = 14
Const LAST_ROW_EXECUTION_TIME_INTERVAL_CELL_NUMBER = 13
Const LAST_ROW_EXECUTION_TIME_UNITS_CELL_NUMBER = 14
Const LAST_ROW_EXECUTION_TIME_ROW_NUMBER = 7
Const REPEAT_INTERVAL_CELL_NUMBER As Integer = 14
Const REPEAT_UNITS_CELL_NUMBER As Integer = 15
Const LAST_COLUMN_LETTER As String = "J"
Const DATE_FORMATTING_STRING As String = "yyyy-m-d"
Const TIME_FORMATTING_STRING As String = "H:mm:ss"
Const PROTECT_PASSWORD As String = "Zukis_Cool1"
Const RASP_PI_DIRECTORY As String = "/home/pi/Desktop/"
Const RASP_PI_USERNAME As String = "pi"
Const RASP_PI_PASSWORD As String = "ERCraspberry@192.168.0.249"
Const PY_LIGHT_COMMAND_FILE As String = "RunLightCommand_v1.1.py"
Const WINSCP_PATH As String = "C:\Program Files (x86)\WinSCP\"
Const QUOTATION As String = """"
Const HOST_KEY As String = "ssh-rsa 2048 13:f0:b2:db:93:db:9d:30:6b:1a:b6:ac:15:76:dc:c3"
Const SESSION_NAME As String = "Raspberry_pi"
Const RASP_PI_INTERFACE_NAME As String = "HortiLight_v1.1.py"
Const RUNLIGHTCOMMAND_FILE_NAME As String = "RunLightCommand_v1.1.py"

'------------------------------------------------------------------------------------------------------------
'Sub: WriteToOutput
'Coded by: Matt Urschel
'Date : 3 May 2017
'Description: Code for button "Write To Output" on Input worksheet - Appends user-entered rows on Input
'             page to end of data on Output page, with start time after a user-entered interval since
'             last line. Repeats pattern for user-specified time interval (times are automatically advanced).
'------------------------------------------------------------------------------------------------------------
Public Sub WriteToOutput()
    On Error GoTo ERROR

   
    
    Dim XCelWorkbook As Excel.Workbook
    Dim XCelSheet1 As Excel.Worksheet
    Dim XCelSheet2 As Excel.Worksheet
    Dim lRowCounter1 As Long: lRowCounter1 = 0
    Dim lRowCounter2 As Long: lRowCounter2 = 0
    Dim lFirstBlankRow As Long
    Dim vArraySheet1(), vArraySheet1Intervals() As Variant
    
    Dim lArrayCounterRowsSheet1 As Long: lArrayCounterRowsSheet1 = 0
    Dim lColumn, lRow, lRowsSheet2 As Long: lColumn = 0: lRow = 0: lRowsSheet2 = 0
    Dim lRepeatInterval, lTimeAfterLastRowInterval, lTimeBetweenRepeatsInterval As Long: lRepeatInterval = 0: lTimeAfterLastRowInterval = 0: lTimeBetweenRepeatsInterval = 0
    Dim sRepeatUnit, sTimeAfterLastRowUnit, sTimeBetweenRepeatsUnit As String
    Dim lPatternInterval, lRepeatIntervalInSeconds, lTimeBetweenRepeatsIntervalInSeconds As Long: lPatternInterval = 0: lRepeatIntervalInSeconds = 0: lTimeBetweenRepeatsIntervalInSeconds = 0
    Dim lNumberOfRepetitions As Long: lNumberOfRepetitions = 0
    Dim sDateStart, sDateEnd, sTimeStart, sTimeEnd, sDateLastRow, sTimeLastRow, sDateLastRowArray, sTimeLastRowArray As String
    Dim lNumberOfNonEmptyRowsSheet1 As Long: lNumberOfNonEmptyRowsSheet1 = 0
    Dim lNumberOfNonEmptyRowsSheet2 As Long: lNumberOfNonEmptyRowsSheet2 = 0
    Dim lRowsInterval As Long: lRowsInterval = 0
    Dim dNewDate As Date
    Dim lTimeDiffLastRowToNewAppend As Long
    Dim lRepeatCounter As Long: lRepeatCounter = 0
    'Dim cbCheckBox As CheckBox
    
    '---------------------------------------------------
    'INITIALIZE EXCEL OBJECTS AND USER-DEFINED VARIABLES
    '---------------------------------------------------
    
    'Initialize workbook and worksheets
    Set XCelWorkbook = Application.ActiveWorkbook
    Set XCelSheet1 = XCelWorkbook.Sheets(1)
    Set XCelSheet2 = XCelWorkbook.Sheets(2)
    
    
    
    'Determine last populated row on worksheet 1
    lNumberOfNonEmptyRowsSheet1 = CountNonEmptyRows(XCelSheet1, NUMBER_OF_COLUMNS)
    
    'Determine last populated row on worksheet 2
    lNumberOfNonEmptyRowsSheet2 = CountNonEmptyRows(XCelSheet2, NUMBER_OF_COLUMNS)
    
    'Initialize row counter
    lRowCounter1 = 2
    
    'Get contents of interval cell if changed
    If Len(Trim(XCelSheet1.Cells(lRowCounter1, REPEAT_INTERVAL_CELL_NUMBER))) > 0 Then
       lRepeatInterval = CLng(XCelSheet1.Cells(lRowCounter1, REPEAT_INTERVAL_CELL_NUMBER))
    End If
    
    'Get contents of unit cell if changed
    If Len(Trim(XCelSheet1.Cells(lRowCounter1, REPEAT_UNITS_CELL_NUMBER))) > 0 Then
       sRepeatUnit = Trim(XCelSheet1.Cells(lRowCounter1, REPEAT_UNITS_CELL_NUMBER))
       
       'Convert interval unit to string for DateAdd function and convert repeat interval to seconds for later comparison to pattern interval
       Select Case sRepeatUnit
            Case "Weeks"
                sRepeatUnit = "ww"
                lRepeatIntervalInSeconds = lRepeatInterval * 604800
            Case "Days"
                sRepeatUnit = "d"
                lRepeatIntervalInSeconds = lRepeatInterval * 86400
            Case "Hours"
                sRepeatUnit = "h"
                lRepeatIntervalInSeconds = lRepeatInterval * 3600
            Case "Minutes"
                sRepeatUnit = "n"
                lRepeatIntervalInSeconds = lRepeatInterval * 60
            Case "Seconds"
                sRepeatUnit = "s"
                lRepeatIntervalInSeconds = lRepeatInterval
            Case "Repeats"
                lNumberOfRepetitions = lRepeatInterval
        End Select
    End If
    
    'Get contents of time after last row interval cell if changed
    If Len(Trim(XCelSheet1.Cells(lRowCounter1 + 1, REPEAT_INTERVAL_CELL_NUMBER))) > 0 Then
       lTimeAfterLastRowInterval = CLng(XCelSheet1.Cells(lRowCounter1 + 1, REPEAT_INTERVAL_CELL_NUMBER))
    End If
    
    'Get contents of time after last row unit cell if changed
    If Len(Trim(XCelSheet1.Cells(lRowCounter1 + 1, REPEAT_UNITS_CELL_NUMBER))) > 0 Then
       sTimeAfterLastRowUnit = Trim(XCelSheet1.Cells(lRowCounter1 + 1, REPEAT_UNITS_CELL_NUMBER))
       
       'Convert interval unit to string for DateAdd function
       Select Case sTimeAfterLastRowUnit
            Case "Weeks"
                sTimeAfterLastRowUnit = "ww"
            Case "Days"
                sTimeAfterLastRowUnit = "d"
            Case "Hours"
                sTimeAfterLastRowUnit = "h"
            Case "Minutes"
                sTimeAfterLastRowUnit = "n"
            Case "Seconds"
                sTimeAfterLastRowUnit = "s"
        End Select
    End If
    
    'Get contents of time between repeats interval cell if changed
    If Len(Trim(XCelSheet1.Cells(lRowCounter1 + 2, REPEAT_INTERVAL_CELL_NUMBER))) > 0 Then
       lTimeBetweenRepeatsInterval = CLng(XCelSheet1.Cells(lRowCounter1 + 2, REPEAT_INTERVAL_CELL_NUMBER))
    End If
    
    'Get contents of time between repeats unit cell if changed
    If Len(Trim(XCelSheet1.Cells(lRowCounter1 + 2, REPEAT_UNITS_CELL_NUMBER))) > 0 Then
       sTimeBetweenRepeatsUnit = Trim(XCelSheet1.Cells(lRowCounter1 + 2, REPEAT_UNITS_CELL_NUMBER))
       
       'Convert interval unit to string for DateAdd function and convert repeat interval to seconds for later comparison to pattern interval
       Select Case sTimeBetweenRepeatsUnit
            Case "Weeks"
                sTimeBetweenRepeatsUnit = "ww"
                lTimeBetweenRepeatsIntervalInSeconds = lTimeBetweenRepeatsInterval * 604800
            Case "Days"
                sTimeBetweenRepeatsUnit = "d"
                lTimeBetweenRepeatsIntervalInSeconds = lTimeBetweenRepeatsInterval * 86400
            Case "Hours"
                sTimeBetweenRepeatsUnit = "h"
                lTimeBetweenRepeatsIntervalInSeconds = lTimeBetweenRepeatsInterval * 3600
            Case "Minutes"
                sTimeBetweenRepeatsUnit = "n"
                lTimeBetweenRepeatsIntervalInSeconds = lTimeBetweenRepeatsInterval * 60
            Case "Seconds"
                sTimeBetweenRepeatsUnit = "s"
                lTimeBetweenRepeatsIntervalInSeconds = lTimeBetweenRepeatsInterval
        End Select
    End If
    
    '---------------
    'DATA VALIDATION
    '---------------
    
    'DO GENERAL WORKSHEET VALIDATION
    If Not CommonDataValidation(XCelSheet1) Then
        Exit Sub
    End If
    
    'IF REPEAT PATTERN INTERVAL, TIME AFTER LAST ROW, OR TIME BETWEEN REPEATS FIELDS ARE EMPTY, THROW ERROR
    If lRepeatInterval = 0 Then

        
        MsgBox "Please enter Repeat pattern interval.", vbExclamation, "Data Entry Error"
        Exit Sub
    End If
    
    If Len(Trim(sRepeatUnit)) = 0 Then

        
        MsgBox "Please enter Repeat pattern units.", vbExclamation, "Data Entry Error"
        Exit Sub
    End If
    
    If (lTimeAfterLastRowInterval = 0) And (lNumberOfNonEmptyRowsSheet2 > 1) Then

        
        MsgBox "Please enter Time after last row interval.", vbExclamation, "Data Entry Error"
        Exit Sub
    End If
    
    If (Len(Trim(sTimeAfterLastRowUnit)) = 0) And (lNumberOfNonEmptyRowsSheet2 > 1) Then

        
        MsgBox "Please enter Time after last row units.", vbExclamation, "Data Entry Error"
        Exit Sub
    End If
    
'    If lTimeBetweenRepeatsInterval = 0 Then
'        'Protect Output worksheet
'        XCelSheet2.Protect (PROTECT_PASSWORD)
'
'        MsgBox "Please enter Time between repeats interval.", vbExclamation, "Data Entry Error"
'        Exit Sub
'    End If
'
'    If Len(Trim(sTimeBetweenRepeatsUnit)) = 0 Then
'        'Protect Output worksheet
'        XCelSheet2.Protect (PROTECT_PASSWORD)
'
'        MsgBox "Please enter Time between repeats units.", vbExclamation, "Data Entry Error"
'        Exit Sub
'    End If
    
    'IF REQUESTED REPEAT INTERVAL IS SMALLER THAN PATTERN INTERVAL, OR IF THERE IS NO DIFFERENCE BETWEEN START AND END TIME, THROW ERROR
    
    'Format start and end date of pattern and convert to string
    sDateStart = Format(Trim(XCelSheet1.Cells(2, 1)), DATE_FORMATTING_STRING)
    sDateEnd = Format(Trim(XCelSheet1.Cells(lNumberOfNonEmptyRowsSheet1, 1)), DATE_FORMATTING_STRING)

    'Format start and end times of pattern and convert to string
    sTimeStart = Format(TimeSerial(XCelSheet1.Cells(2, 2), XCelSheet1.Cells(2, 3), XCelSheet1.Cells(2, 4)), TIME_FORMATTING_STRING)
    sTimeEnd = Format(TimeSerial(XCelSheet1.Cells(lNumberOfNonEmptyRowsSheet1, 2), XCelSheet1.Cells(lNumberOfNonEmptyRowsSheet1, 3), XCelSheet1.Cells(lNumberOfNonEmptyRowsSheet1, 4)), TIME_FORMATTING_STRING)
    
    'Determine time interval between first and last rows (in seconds)
    lPatternInterval = DateDiff("s", CDate(sDateStart & " " & sTimeStart), CDate(sDateEnd & " " & sTimeEnd))
    
    'If user specified the number of repeats, don't worry about time intervals
    If sRepeatUnit <> "Repeats" Then
        'If user did not specify number of repeats, make sure time intervals make sense
        If (lPatternInterval > 0) Then
            'If pattern interval is smaller than repeat interval, throw error
            If (lRepeatIntervalInSeconds < lPatternInterval) Then
            
               MsgBox "Please enter repeat time interval that is larger than pattern time interval.", vbExclamation, "Data Entry Error"
               Exit Sub
            ElseIf ((lPatternInterval + lTimeBetweenRepeatsIntervalInSeconds) > lRepeatIntervalInSeconds) Then
            
               MsgBox "The sum of the duration of the repeated pattern and the time between repeats must be smaller than the repeat time interval.", vbExclamation, "Data Entry Error"
               Exit Sub
            Else
               'Number of times pattern can be repeated in given time interval
               
               lNumberOfRepetitions = Round(lRepeatIntervalInSeconds / (lPatternInterval + lTimeBetweenRepeatsIntervalInSeconds))
            End If
            
        Else
           MsgBox "Time difference between first and last row must be greater than zero.", vbExclamation, "Data Entry Error"
           Exit Sub
        End If
    End If
    
    If lNumberOfRepetitions > 1 Then
        If lTimeBetweenRepeatsInterval = 0 Then

        
            MsgBox "Please enter Time between repeats interval.", vbExclamation, "Data Entry Error"
            Exit Sub
        End If
    
        If Len(Trim(sTimeBetweenRepeatsUnit)) = 0 Then

            
            MsgBox "Please enter Time between repeats units.", vbExclamation, "Data Entry Error"
            Exit Sub
        End If
    End If
        

        
    '------------------------------------------
    'POPULATE ARRAY WITH CONTENTS OF INPUT PAGE
    '------------------------------------------
      
    'Determine array dimensions from Input worksheet
'    lArrayCounterRowsSheet1 = lNumberOfNonEmptyRowsSheet1
'
'    ReDim vArraySheet1(1 To lArrayCounterRowsSheet1 - 1, 1 To NUMBER_OF_COLUMNS)
    
    'Populate array from Input worksheet
'    lRowCounter1 = 2
'    lArrayCounterRowsSheet1 = 1
'
'    Do While Len(Trim(XCelSheet1.Cells(lRowCounter1, 1))) > 0
'
'        For lColumn = 1 To UBound(vArraySheet1, 2)
'            vArraySheet1(lArrayCounterRowsSheet1, lColumn) = Trim(XCelSheet1.Cells(lRowCounter1, lColumn))
'        Next lColumn
'
'        lArrayCounterRowsSheet1 = lArrayCounterRowsSheet1 + 1
'        lRowCounter1 = lRowCounter1 + 1
'    Loop
    
    
    vArraySheet1 = PopulateWorksheetArray(XCelSheet1, lNumberOfNonEmptyRowsSheet1 - 1, NUMBER_OF_COLUMNS)
    
    ReDim vArraySheet1Intervals(1 To lNumberOfNonEmptyRowsSheet1 - 1)
    
    'Populate intervals array with intervals between rows in worksheet 1 array
    For lArrayCounterRowsSheet1 = 2 To UBound(vArraySheet1)
        'Get date of previous row in worksheet 1 array
        sDateStart = Format(vArraySheet1(lArrayCounterRowsSheet1 - 1, 1), DATE_FORMATTING_STRING)
        
        'Get time of previous row in worksheet 1 array
        sTimeStart = Format(TimeSerial(vArraySheet1(lArrayCounterRowsSheet1 - 1, 2), vArraySheet1(lArrayCounterRowsSheet1 - 1, 3), vArraySheet1(lArrayCounterRowsSheet1 - 1, 4)), TIME_FORMATTING_STRING)
               
        'Get date of this row in worksheet 1 array
        sDateEnd = Format(vArraySheet1(lArrayCounterRowsSheet1, 1), DATE_FORMATTING_STRING)
        
        'Get time of this row in worksheet 1 array
        sTimeEnd = Format(TimeSerial(vArraySheet1(lArrayCounterRowsSheet1, 2), vArraySheet1(lArrayCounterRowsSheet1, 3), vArraySheet1(lArrayCounterRowsSheet1, 4)), TIME_FORMATTING_STRING)
       
        'Populate appropriate row in intervals array difference between times
        vArraySheet1Intervals(lArrayCounterRowsSheet1 - 1) = DateDiff("s", CDate(sDateStart & " " & sTimeStart), CDate(sDateEnd & " " & sTimeEnd))
    Next lArrayCounterRowsSheet1
    
    '------------------------------------------------------------------------------------------------------------------------------------------
    'CHANGE DATE AND TIME VALUES BASED ON REPETITION INTERVAL AND APPEND CONTENTS OF INPUT PAGE TO OUTPUT PAGE WITH REQUESTED NUMBER OF REPEATS
    '------------------------------------------------------------------------------------------------------------------------------------------
    
    'Unprotect Output worksheet
    XCelSheet2.Unprotect (PROTECT_PASSWORD)
    
    'Initialize counters
    lFirstBlankRow = lNumberOfNonEmptyRowsSheet2 + 1
    
    'DETERMINE DATE/TIME OF LAST ROW ON OUTPUT PAGE
    
    'If Output page has at least one row
    If (lFirstBlankRow > 2) Then
    
        'Get date of last row on output page
        sDateLastRow = Format(Trim(XCelSheet2.Cells(lFirstBlankRow - 1, 1)), DATE_FORMATTING_STRING)
        
        'Get time of last row on output page
        sTimeLastRow = Format(TimeSerial(XCelSheet2.Cells(lFirstBlankRow - 1, 2), XCelSheet2.Cells(lFirstBlankRow - 1, 3), XCelSheet2.Cells(lFirstBlankRow - 1, 4)), TIME_FORMATTING_STRING)
        
        'Get date of last row in Input worksheet array
        sDateLastRowArray = Format(vArraySheet1(UBound(vArraySheet1, 1), 1), DATE_FORMATTING_STRING)
        
        'Get time of last row in Input worksheet array
        sTimeLastRowArray = Format(TimeSerial(vArraySheet1(UBound(vArraySheet1, 1), 2), vArraySheet1(UBound(vArraySheet1, 1), 3), vArraySheet1(UBound(vArraySheet1, 1), 4)), TIME_FORMATTING_STRING)
                
        lTimeDiffLastRowToNewAppend = DateDiff("s", CDate(sDateLastRow & " " & sTimeLastRow), CDate(sDateLastRowArray & " " & sTimeLastRowArray))
        'For each repeat
        For lRepeatCounter = 1 To lNumberOfRepetitions
            'If this is the first repeat
            If (lRepeatCounter = 1) Then
'                'Get date of last row on output page
'                sDateLastRow = Format(Trim(XCelSheet2.Cells(lFirstBlankRow - 1, 1)), DATE_FORMATTING_STRING)
'
'                'Get time of last row on output page
'                sTimeLastRow = Format(TimeSerial(XCelSheet2.Cells(lFirstBlankRow - 1, 2), XCelSheet2.Cells(lFirstBlankRow - 1, 3), XCelSheet2.Cells(lFirstBlankRow - 1, 4)), TIME_FORMATTING_STRING)
                If lTimeDiffLastRowToNewAppend <= 0 Then
                    'Add user defined time after last row to first row of repeat
                    dNewDate = DateAdd(sTimeAfterLastRowUnit, CDbl(lTimeAfterLastRowInterval), CDate(sDateLastRow & " " & sTimeLastRow))
                
                    'Change first row of array
                    vArraySheet1(1, 1) = Format(dNewDate, "m/d/yyyy")
                    vArraySheet1(1, 2) = Format(dNewDate, "HH")
                    vArraySheet1(1, 3) = Format(dNewDate, "nn")
                    vArraySheet1(1, 4) = Format(dNewDate, "ss")
                
            
                    For lArrayCounterRowsSheet1 = 2 To UBound(vArraySheet1, 1)
                    
                        'Get date of last row on output page
                        sDateLastRow = Format(vArraySheet1(lArrayCounterRowsSheet1 - 1, 1), DATE_FORMATTING_STRING)
            
                        'Get time of last row on output page
                        sTimeLastRow = Format(TimeSerial(vArraySheet1(lArrayCounterRowsSheet1 - 1, 2), vArraySheet1(lArrayCounterRowsSheet1 - 1, 3), vArraySheet1(lArrayCounterRowsSheet1 - 1, 4)), TIME_FORMATTING_STRING)
                                 
                                 
                        dNewDate = DateAdd("s", CDbl(vArraySheet1Intervals(lArrayCounterRowsSheet1 - 1)), CDate(sDateLastRow & " " & sTimeLastRow))
                        
                        'Change first row of array
                        vArraySheet1(lArrayCounterRowsSheet1, 1) = Format(dNewDate, "m/d/yyyy")
                        vArraySheet1(lArrayCounterRowsSheet1, 2) = Format(dNewDate, "HH")
                        vArraySheet1(lArrayCounterRowsSheet1, 3) = Format(dNewDate, "nn")
                        vArraySheet1(lArrayCounterRowsSheet1, 4) = Format(dNewDate, "ss")
            
                    Next lArrayCounterRowsSheet1
                End If
            
            Else 'If this is not the first repeat
            
                'Get date of last row
                sDateLastRow = Format(vArraySheet1(UBound(vArraySheet1, 1), 1), DATE_FORMATTING_STRING)
        
                'Get time of last row
                sTimeLastRow = Format(TimeSerial(vArraySheet1(UBound(vArraySheet1, 1), 2), vArraySheet1(UBound(vArraySheet1, 1), 3), vArraySheet1(UBound(vArraySheet1, 1), 4)), TIME_FORMATTING_STRING)
                             
                dNewDate = DateAdd(sTimeBetweenRepeatsUnit, CDbl(lTimeBetweenRepeatsInterval), CDate(sDateLastRow & " " & sTimeLastRow))
                
                vArraySheet1(1, 1) = Format(dNewDate, "m/d/yyyy")
                vArraySheet1(1, 2) = Format(dNewDate, "HH")
                vArraySheet1(1, 3) = Format(dNewDate, "nn")
                vArraySheet1(1, 4) = Format(dNewDate, "ss")
                
                For lArrayCounterRowsSheet1 = 2 To UBound(vArraySheet1, 1)
                
                    'Get date of last row on output page
                    sDateLastRow = Format(vArraySheet1(lArrayCounterRowsSheet1 - 1, 1), DATE_FORMATTING_STRING)
        
                    'Get time of last row on output page
                    sTimeLastRow = Format(TimeSerial(vArraySheet1(lArrayCounterRowsSheet1 - 1, 2), vArraySheet1(lArrayCounterRowsSheet1 - 1, 3), vArraySheet1(lArrayCounterRowsSheet1 - 1, 4)), TIME_FORMATTING_STRING)
                             
                             
                    dNewDate = DateAdd("s", CDbl(vArraySheet1Intervals(lArrayCounterRowsSheet1 - 1)), CDate(sDateLastRow & " " & sTimeLastRow))
                    
                    'Change first row of array
                    vArraySheet1(lArrayCounterRowsSheet1, 1) = Format(dNewDate, "m/d/yyyy")
                    vArraySheet1(lArrayCounterRowsSheet1, 2) = Format(dNewDate, "HH")
                    vArraySheet1(lArrayCounterRowsSheet1, 3) = Format(dNewDate, "nn")
                    vArraySheet1(lArrayCounterRowsSheet1, 4) = Format(dNewDate, "ss")
        
                Next lArrayCounterRowsSheet1
            End If
            
            
            'Append contents of worksheet1 to end of rows in worksheet2
            For lColumn = 1 To UBound(vArraySheet1, 2)
                lRowCounter2 = lFirstBlankRow
                For lRowsSheet2 = 1 To UBound(vArraySheet1, 1)
                    XCelSheet2.Cells(lRowCounter2, lColumn).Value = vArraySheet1(lRowsSheet2, lColumn)
                    lRowCounter2 = lRowCounter2 + 1
                Next lRowsSheet2
            Next lColumn
            
                    
            lFirstBlankRow = lRowCounter2
        Next lRepeatCounter
        
    'If Output page has no rows
    Else
        For lRepeatCounter = 1 To lNumberOfRepetitions
            If lRepeatCounter > 1 Then 'If this is not the first repeat
                'Get date of last row in sheet array
                sDateLastRow = Format(vArraySheet1(UBound(vArraySheet1, 1), 1), DATE_FORMATTING_STRING)
    
                'Get time of last row in sheet array
                sTimeLastRow = Format(TimeSerial(vArraySheet1(UBound(vArraySheet1, 1), 2), vArraySheet1(UBound(vArraySheet1, 1), 3), vArraySheet1(UBound(vArraySheet1, 1), 4)), TIME_FORMATTING_STRING)
                         
                         
                dNewDate = DateAdd(sTimeBetweenRepeatsUnit, CDbl(lTimeBetweenRepeatsInterval), CDate(sDateLastRow & " " & sTimeLastRow))
                
                'Change first row of array
                vArraySheet1(1, 1) = Format(dNewDate, "m/d/yyyy")
                vArraySheet1(1, 2) = Format(dNewDate, "HH")
                vArraySheet1(1, 3) = Format(dNewDate, "nn")
                vArraySheet1(1, 4) = Format(dNewDate, "ss")
        
            
                For lArrayCounterRowsSheet1 = 2 To UBound(vArraySheet1, 1)
                    'Get date of last row on output page
                    sDateLastRow = Format(vArraySheet1(lArrayCounterRowsSheet1 - 1, 1), DATE_FORMATTING_STRING)
        
                    'Get time of last row on output page
                    sTimeLastRow = Format(TimeSerial(vArraySheet1(lArrayCounterRowsSheet1 - 1, 2), vArraySheet1(lArrayCounterRowsSheet1 - 1, 3), vArraySheet1(lArrayCounterRowsSheet1 - 1, 4)), TIME_FORMATTING_STRING)
                             
                             
                    dNewDate = DateAdd("s", CDbl(vArraySheet1Intervals(lArrayCounterRowsSheet1 - 1)), CDate(sDateLastRow & " " & sTimeLastRow))
                    
                    'Change first row of array
                    vArraySheet1(lArrayCounterRowsSheet1, 1) = Format(dNewDate, "m/d/yyyy")
                    vArraySheet1(lArrayCounterRowsSheet1, 2) = Format(dNewDate, "HH")
                    vArraySheet1(lArrayCounterRowsSheet1, 3) = Format(dNewDate, "nn")
                    vArraySheet1(lArrayCounterRowsSheet1, 4) = Format(dNewDate, "ss")
        
                Next lArrayCounterRowsSheet1
            End If
            
            
            
            'Append contents of worksheet1 to end of rows in worksheet2
            For lColumn = 1 To UBound(vArraySheet1, 2)
                lRowCounter2 = lFirstBlankRow
                For lRowsSheet2 = 1 To UBound(vArraySheet1, 1)
                    XCelSheet2.Cells(lRowCounter2, lColumn).Value = vArraySheet1(lRowsSheet2, lColumn)
                    lRowCounter2 = lRowCounter2 + 1
                Next lRowsSheet2
            Next lColumn
            
            
            
            lFirstBlankRow = lRowCounter2
        Next lRepeatCounter
    End If
    
    
    'Protect Output worksheet
    XCelSheet2.Protect (PROTECT_PASSWORD)
    
   
    On Error Resume Next
    'Close iFileNum1

    Set XCelSheet1 = Nothing
    Set XCelSheet2 = Nothing
    Set XCelWorkbook = Nothing
    On Error GoTo ERROR

Exit Sub
    
ERROR:
    MsgBox Err.Description, vbCritical, "Error"
    On Error Resume Next
    
    'Protect Output worksheet
    XCelSheet2.Protect (PROTECT_PASSWORD)

    Set XCelSheet1 = Nothing
    Set XCelSheet2 = Nothing

    Set XCelWorkbook = Nothing
    
End Sub

'--------------------------------------------------------------------------------------------------------------
'Sub: WriteToFile
'Coded by: Matt Urschel
'Date : 3 May 2017
'Description: Code for button "Write to File" on Output worksheet - Writes data on Output worksheet to
'             text file (user-defined file name or default file name defined by string constant).
'Change Log:
'5/10/2017 Added code to ask user before overwriting file and give user option to append data to output file
'5/17/2017 Added code to insert 'lights off' command (all channels set to -1) at end of file after user defined
'          time interval
'--------------------------------------------------------------------------------------------------------------
Public Sub WriteToFile()
    On Error GoTo ERROR
    
    Dim XCelWorkbook As Excel.Workbook
    Dim XCelSheet2 As Excel.Worksheet
    'Dim XCelSheet1 As Excel.Worksheet
    Dim lRowCounter2 As Long
    Dim iColumnCounter As Integer
    Dim sLine As String
    Dim varrLine As Variant
    Dim sDate, sTime As String
    Dim iFileNum, iFileNum1, i As Integer
    Dim lTimeToExecuteLastRowInterval As Long
    Dim sTimeToExecuteLastRowUnits As String
    Dim sLastRowDate As String
    
            
    Dim sOutputFile As String
    
    'Initialize workbook and worksheets
    Set XCelWorkbook = Application.ActiveWorkbook
    Set XCelSheet2 = XCelWorkbook.Sheets(2)
    'Set XCelSheet1 = XCelWorkbook.Sheets(1)
    
    'Initialize row counter
    lRowCounter2 = 2
    
    'Get free file handle
    iFileNum = 1
    iFileNum1 = 2
    
    'Set file name to user-entered name on Output worksheet
    
    If Len(Trim(XCelSheet2.Cells(2, FILE_NAME_CELL_NUMBER))) > 0 Then
        sOutputFile = ActiveWorkbook.Path & "\" & Trim(XCelSheet2.Cells(2, FILE_NAME_CELL_NUMBER))
    Else
        sOutputFile = ActiveWorkbook.Path & "\" & DEFAULT_FILE_NAME
    End If
    
    'Get contents of interval cell if changed
    If Len(Trim(XCelSheet2.Cells(LAST_ROW_EXECUTION_TIME_ROW_NUMBER, LAST_ROW_EXECUTION_TIME_INTERVAL_CELL_NUMBER))) > 0 Then
       lTimeToExecuteLastRowInterval = CLng(XCelSheet2.Cells(LAST_ROW_EXECUTION_TIME_ROW_NUMBER, LAST_ROW_EXECUTION_TIME_INTERVAL_CELL_NUMBER))
    Else
       MsgBox "Please enter amount of time to execute last row.", vbExclamation, "Last row execution time"
       Exit Sub
    End If
    
    'Get contents of unit cell if changed
    If Len(Trim(XCelSheet2.Cells(LAST_ROW_EXECUTION_TIME_ROW_NUMBER, LAST_ROW_EXECUTION_TIME_UNITS_CELL_NUMBER))) > 0 Then
       sTimeToExecuteLastRowUnits = Trim(XCelSheet2.Cells(LAST_ROW_EXECUTION_TIME_ROW_NUMBER, LAST_ROW_EXECUTION_TIME_UNITS_CELL_NUMBER))
       
       'Convert interval unit to string for DateAdd function and convert repeat interval to seconds for later comparison to pattern interval
       Select Case sTimeToExecuteLastRowUnits
            Case "Weeks"
                sTimeToExecuteLastRowUnits = "ww"
            Case "Days"
                sTimeToExecuteLastRowUnits = "d"
            Case "Hours"
                sTimeToExecuteLastRowUnits = "h"
            Case "Minutes"
                sTimeToExecuteLastRowUnits = "n"
            Case "Seconds"
                sTimeToExecuteLastRowUnits = "s"
        End Select
    Else
        MsgBox "Please enter units for last row execution time.", vbExclamation, "Last row execution time"
        Exit Sub
    End If
    
    
    '---------------
    'DATA VALIDATION
    '---------------
    
    'DO GENERAL WORKSHEET VALIDATION
    If Not CommonDataValidation(XCelSheet2) Then
        Exit Sub
    End If
    
    '---------------------------------------------------------------
    'ASK USER IF THEY WANT TO OVERWRITE OR APPEND EXISTING DATA FILE
    '---------------------------------------------------------------
    
    'Check if file exists
    If Dir(sOutputFile) <> "" Then
        'Confirm overwrite
        If MsgBox("File already exists. Do you wish to overwrite it?", vbYesNo + vbQuestion, "File Overwrite") = vbYes Then
            
            'Delete file if it already exists
            On Error Resume Next
            Close iFileNum
            Kill sOutputFile
            On Error GoTo ERROR
            
            'Open file for output
            Open sOutputFile For Output As iFileNum
        Else 'Ask about append
            If MsgBox("Do you want to append current file with new rows?", vbYesNo + vbQuestion, "File Overwrite") = vbYes Then
                Open sOutputFile For Input As iFileNum
                Open sOutputFile & "_temp" For Output As iFileNum1
                
                
                'Get last row of data from file
                Do While Not EOF(iFileNum)
                    Line Input #iFileNum, sLine
                    varrLine = Split(sLine, " ")
                    'Write existing rows to temporary data file (except lights of command)
                    If InStr(sLine, "-1") = 0 Then
                        Print #iFileNum1, sLine
                    End If
                Loop
                
                'Get date/time of first row on output page
                sDate = Format(Trim(XCelSheet2.Cells(2, 1)), DATE_FORMATTING_STRING)
                sTime = Format(TimeSerial(XCelSheet2.Cells(2, 2), XCelSheet2.Cells(2, 3), XCelSheet2.Cells(2, 4)), TIME_FORMATTING_STRING)
                    
                'If date/time on first row of output page is less than or equal to date/time on last row of data file, throw error
                If DateDiff("s", CDate(varrLine(0) & " " & varrLine(1)), CDate(sDate & " " & sTime)) <= 0 Then
                    MsgBox "First row of appended data must have a date and time that is greater than that of the last line in the data file.", vbExclamation, "Append Error"
                    Close iFileNum
                    On Error Resume Next
                    'Close and delete temporary data file
                    Close iFileNum1
                    Kill sOutputFile & "_temp"
                    On Error GoTo ERROR
                    Exit Sub
                End If
                
                Close iFileNum
                Close iFileNum1
                'Delete original data file
                Kill (sOutputFile)
                'Rename temp data file as original data file
                Name sOutputFile & "_temp" As sOutputFile
                'Open file for append
                Open sOutputFile For Append As iFileNum
            Else
                'Quit without saving
                MsgBox "Data not saved.", vbExclamation, "File Save"
                Exit Sub
                
            End If
        End If
    Else
        'Open file for output
        Open sOutputFile For Output As iFileNum
    End If
    
    '---------------------------------------------------------------
    'WRITE DATA ON OUTPUT WORKSHEET TO TEXT FILE
    '---------------------------------------------------------------
    
    Do While Len(Trim(XCelSheet2.Cells(lRowCounter2, 1))) > 0
    
        'Initialize string variables to null
        sLine = ""
        sDate = ""
        sTime = ""

        'Format date and convert to string
        sDate = Format(Trim(XCelSheet2.Cells(lRowCounter2, 1)), DATE_FORMATTING_STRING)
        sLine = sLine & sDate & TEXT_FILE_DELIMITER
        
        'Format time and convert to string
        sTime = Format(TimeSerial(XCelSheet2.Cells(lRowCounter2, 2), XCelSheet2.Cells(lRowCounter2, 3), XCelSheet2.Cells(lRowCounter2, 4)), TIME_FORMATTING_STRING)
        sLine = sLine & sTime & TEXT_FILE_DELIMITER
                
        'Build output line string
        For iColumnCounter = 5 To NUMBER_OF_COLUMNS
            sLine = sLine & Trim(XCelSheet2.Cells(lRowCounter2, iColumnCounter)) & TEXT_FILE_DELIMITER
        Next iColumnCounter
        
        sLine = sLine & "X"
        
        'Write line string to text file
        Print #iFileNum, sLine
        
        lRowCounter2 = lRowCounter2 + 1
    Loop
    
    'Initialize string variables to null
    sLine = ""
    
    'Insert lights off command at end of file with time advanced based on user-defined time interval after last row
    sLastRowDate = DateAdd(sTimeToExecuteLastRowUnits, CDbl(lTimeToExecuteLastRowInterval), CDate(sDate & " " & sTime))
    
    sDate = Format(Left(sLastRowDate, InStr(sLastRowDate, " ") - 1), DATE_FORMATTING_STRING)
    sLine = sLine & sDate & TEXT_FILE_DELIMITER
    
    sTime = Format(Right(sLastRowDate, Len(sLastRowDate) - InStr(sLastRowDate, " ")), TIME_FORMATTING_STRING)
    sLine = sLine & sTime & TEXT_FILE_DELIMITER
    
    For i = 1 To 6
        sLine = sLine & "0" & TEXT_FILE_DELIMITER
    Next i
    
    sLine = sLine & "X"
    
    Print #iFileNum, sLine
    
    'Initialize string variables to null
    sLine = ""
    
    'Insert lights off command at end of file with time advanced based on user-defined time interval after last row
    sLastRowDate = DateAdd(sTimeToExecuteLastRowUnits, CDbl(lTimeToExecuteLastRowInterval), CDate(sDate & " " & sTime))
    
    sDate = Format(Left(sLastRowDate, InStr(sLastRowDate, " ") - 1), DATE_FORMATTING_STRING)
    sLine = sLine & sDate & TEXT_FILE_DELIMITER
    
    sTime = Format(Right(sLastRowDate, Len(sLastRowDate) - InStr(sLastRowDate, " ")), TIME_FORMATTING_STRING)
    sLine = sLine & sTime & TEXT_FILE_DELIMITER
    
    For i = 1 To 6
        sLine = sLine & "-1" & TEXT_FILE_DELIMITER
    Next i
    
    sLine = sLine & "X"
    
    Print #iFileNum, sLine
    
    Close iFileNum
    
    
Exit Sub
    
ERROR:
    MsgBox Err.Description, vbCritical, "Error"
    On Error Resume Next
    Close iFileNum
    Close iFileNum1
    Set XCelSheet2 = Nothing
    Set XCelWorkbook = Nothing
    
End Sub

Public Sub ClearOutput()
    On Error GoTo ERROR
    
    Dim XCelWorkbook As Excel.Workbook: Set XCelWorkbook = Application.ActiveWorkbook
    Dim XCelSheet2 As Excel.Worksheet: Set XCelSheet2 = XCelWorkbook.Sheets(2)
    Dim vColumnNamesArray As Variant
    Dim i As Integer
    
    'Unprotect Output worksheet
    XCelSheet2.Unprotect (PROTECT_PASSWORD)
    
    'Put column names from COLUMN_NAMES constant into array
    vColumnNamesArray = Split(COLUMN_NAMES, ",")
    
    
    'Clear cells and repopulate column names
    With XCelSheet2
        .Columns("A:" & LAST_COLUMN_LETTER).ClearContents
        
        For i = 1 To NUMBER_OF_COLUMNS
            .Rows("1").Columns(i).Value = vColumnNamesArray(i - 1)
        Next i
    End With
    
    'Clear file name cell
    XCelSheet2.Cells(2, FILE_NAME_CELL_NUMBER).Value = ""
    
    'Protect Output worksheet
    XCelSheet2.Protect (PROTECT_PASSWORD)
 
Exit Sub
    
ERROR:
    MsgBox Err.Description, vbCritical, "Error"
    On Error Resume Next
    'Protect Output worksheet
    XCelSheet2.Protect (PROTECT_PASSWORD)

    Set XCelSheet2 = Nothing

    Set XCelWorkbook = Nothing
End Sub

Public Sub ClearInput()
    On Error GoTo ERROR
    
    Dim XCelWorkbook As Excel.Workbook: Set XCelWorkbook = Application.ActiveWorkbook
    Dim XCelSheet1 As Excel.Worksheet: Set XCelSheet1 = XCelWorkbook.Sheets(1)
    Dim vColumnNamesArray As Variant
    Dim i As Integer
    
    'Unprotect Output worksheet
    XCelSheet1.Unprotect (PROTECT_PASSWORD)
    
    'Put column names from COLUMN_NAMES constant into array
    vColumnNamesArray = Split(COLUMN_NAMES, ",")
    
    
    'Clear cells and repopulate column names
    With XCelSheet1
        .Columns("A:" & LAST_COLUMN_LETTER).ClearContents
        
        For i = 1 To NUMBER_OF_COLUMNS
            .Rows("1").Columns(i).Value = vColumnNamesArray(i - 1)
        Next i
    End With
    
    XCelSheet1.Cells(2, REPEAT_INTERVAL_CELL_NUMBER).Value = ""
    XCelSheet1.Cells(2, REPEAT_INTERVAL_CELL_NUMBER + 1).Value = ""
    XCelSheet1.Cells(3, REPEAT_INTERVAL_CELL_NUMBER).Value = ""
    XCelSheet1.Cells(3, REPEAT_INTERVAL_CELL_NUMBER + 1).Value = ""
    XCelSheet1.Cells(4, REPEAT_INTERVAL_CELL_NUMBER).Value = ""
    XCelSheet1.Cells(4, REPEAT_INTERVAL_CELL_NUMBER + 1).Value = ""
        
    'Protect Output worksheet
    XCelSheet1.Protect (PROTECT_PASSWORD)
 
Exit Sub
    
ERROR:
    MsgBox Err.Description, vbCritical, "Error"
    On Error Resume Next

    Set XCelSheet1 = Nothing

    Set XCelWorkbook = Nothing
End Sub

'Function to determine last populated row on a worksheet
Private Function CountNonEmptyRows(xCelSheet As Excel.Worksheet, lNumberOfColumnsToCheck As Long) As Long
    Dim i, j, lNumberOfPopulatedColumns As Long
    Dim lRowCounter As Long: lRowCounter = 2
    Dim lNumberOfPopulatedRows As Long: lNumberOfPopulatedRows = 0
    
    lNumberOfPopulatedColumns = 1

    With xCelSheet
        Do While lNumberOfPopulatedColumns > 0
        
            lNumberOfPopulatedColumns = 0
            
            For i = 1 To lNumberOfColumnsToCheck
                If Len(Trim(xCelSheet.Cells(lRowCounter, i))) > 0 Then
                    lNumberOfPopulatedColumns = lNumberOfPopulatedColumns + 1
                End If
            Next i
        
            If lNumberOfPopulatedColumns = 0 Then
                lNumberOfPopulatedRows = lRowCounter - 1
                CountNonEmptyRows = lNumberOfPopulatedRows
                Exit Function
            End If
        
            lRowCounter = lRowCounter + 1
        Loop
    End With
    
    
End Function

Public Sub RemoveToolbars()

    On Error Resume Next

        With Application

           .DisplayFullScreen = True

           .CommandBars("Full Screen").Visible = False

           .CommandBars("Worksheet Menu Bar").Enabled = False

        End With
        

    On Error GoTo 0

End Sub

Public Sub RestoreToolbars()

    On Error Resume Next

        With Application

           .DisplayFullScreen = False

           .CommandBars("Worksheet Menu Bar").Enabled = True

        End With

    On Error GoTo 0

End Sub

Public Sub UploadFileToRaspberryPi()
    Dim sCommandLine As String
    Dim iFileNum As Integer: iFileNum = FreeFile
    Dim sFileName As String: sFileName = Application.ActiveWorkbook.Path & "\move_file.txt"
    Dim sLine As String
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    
    
    On Error GoTo ERROR
    
    If SystemOnline(Right(RASP_PI_PASSWORD, Len(RASP_PI_PASSWORD) - InStr(RASP_PI_PASSWORD, "@"))) Then
        'Open file for output
        Open sFileName For Output As iFileNum
        
        'Create WinSCP script file to upload data file to Raspberry Pi
        sLine = "open " & RASP_PI_USERNAME & ":" & RASP_PI_PASSWORD & "/ -hostkey=" & QUOTATION & HOST_KEY & QUOTATION & vbCrLf
        
        Print #iFileNum, sLine
        
        sLine = "put " & Application.ActiveWorkbook.Path & "\Data.txt " & RASP_PI_DIRECTORY & vbCrLf
        
        Print #iFileNum, sLine
        Print #iFileNum, "exit"
        
        Close iFileNum
        
        'Run WinSCP script file
        sCommandLine = QUOTATION & WINSCP_PATH & "winscp.com" & QUOTATION & " /ini=nul /script=" & QUOTATION & Application.ActiveWorkbook.Path & "\move_file.txt" & QUOTATION
        
        wsh.Run sCommandLine, windowStyle, waitOnReturn
                          
        'Call Shell(sCommandLine)
        'Application.Wait (Now + TimeValue("00:00:03"))
        'Delete file for security purposes
        On Error Resume Next
        Kill (sFileName)
        On Error GoTo ERROR
    Else
        'Application.Speech.Speak ("Please connect to Buffalo network.")
        MsgBox "You must be connected to domain 'BUFFALO' to use this feature.", vbCritical, "Connection Error"
    End If
            
    
    
    Exit Sub
    
ERROR:
        MsgBox Err.Description, vbCritical, "Error"
        On Error Resume Next
        Close iFileNum
        Kill (sFileName)

        
End Sub

Public Sub RunLightCommand()
    Dim sCommandLine As String
    Dim wsh, oExec As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    Dim i As Integer
    Dim XCelWorkbook As Excel.Workbook: Set XCelWorkbook = Application.ActiveWorkbook
    Dim XCelSheet3 As Excel.Worksheet: Set XCelSheet3 = XCelWorkbook.Sheets(3)
    Dim iColumn As Integer
    Dim bHortiLightShutOff As Boolean: bHortiLightShutOff = False

    On Error GoTo ERROR
    
    'XCelSheet3.Unprotect (PROTECT_PASSWORD)
    
    '---------------
    'DATA VALIDATION
    '---------------
    
    'If not all columns are populated, throw error
    For iColumn = 1 To NUMBER_OF_COLUMNS - 4
        If Len(Trim(XCelSheet3.Cells(2, iColumn))) = 0 Then
            'Protect Output worksheet
            'XCelSheet3.Protect (PROTECT_PASSWORD)
    
            MsgBox "Please fill in missing data in column " & iColumn & ".", vbExclamation, "Data Entry Error"
            Exit Sub
        End If
    Next iColumn
    
    '--------------------------------------------------------------------
    'RUN PYTHON SCRIPT ON PUTTY WITH USER-GIVEN CHANNEL PERCENTAGE VALUES
    '--------------------------------------------------------------------
    
    '**************************
    '*** NEEDS TO BE TESTED ***
    '**************************
    
    'If we're connected to BUFFALO
    If SystemOnline(Right(RASP_PI_PASSWORD, Len(RASP_PI_PASSWORD) - InStr(RASP_PI_PASSWORD, "@"))) Then
        'If HortiLight interface is running
        If ProcessRunning(RASP_PI_INTERFACE_NAME) Then
            'Ask if it's ok to kill HortiLight interface
            If MsgBox(RASP_PI_INTERFACE_NAME & " must be shut down to run commands. Is that OK?", vbYesNo + vbQuestion, "Shut down interface?") = vbYes Then
                sCommandLine = "plink " & RASP_PI_USERNAME & Right(RASP_PI_PASSWORD, Len(RASP_PI_PASSWORD) - InStr(RASP_PI_PASSWORD, "@") + 1) & _
                            " -pw " & Left(RASP_PI_PASSWORD, InStr(RASP_PI_PASSWORD, "@") - 1) & _
                            " -batch" & _
                            " pkill -f " & RASP_PI_INTERFACE_NAME
    
                wsh.Run sCommandLine, windowStyle, waitOnReturn
    
                bHortiLightShutOff = True
            Else
                Exit Sub
            End If
        End If

        'Run run python script
        sCommandLine = "plink " & RASP_PI_USERNAME & Right(RASP_PI_PASSWORD, Len(RASP_PI_PASSWORD) - InStr(RASP_PI_PASSWORD, "@") + 1) & _
                       " -pw " & Left(RASP_PI_PASSWORD, InStr(RASP_PI_PASSWORD, "@") - 1) & _
                       " python " & RASP_PI_DIRECTORY & RUNLIGHTCOMMAND_FILE_NAME & TEXT_FILE_DELIMITER
                       
         For i = 1 To NUMBER_OF_COLUMNS - 4
            sCommandLine = sCommandLine & Trim(XCelSheet3.Cells(2, i)) & TEXT_FILE_DELIMITER
         Next i


         wsh.Run sCommandLine, windowStyle, waitOnReturn
        

    Else
        MsgBox "You must be connected to domain 'BUFFALO' to use this feature.", vbCritical, "Connection Error"
    End If

    'Protect Output worksheet
    'XCelSheet3.Protect (PROTECT_PASSWORD)

    Exit Sub

ERROR:
        MsgBox Err.Description, vbCritical, "Error"
        On Error Resume Next
'        Close iFileNum
'        Kill (Me.Path & "\" & PUTTY_SCRIPT_NAME)
        Set XCelSheet3 = Nothing
        Set XCelWorkbook = Nothing
        'Protect Output worksheet
        'XCelSheet3.Protect (PROTECT_PASSWORD)
End Sub

'Determine if device is online
Function SystemOnline(ByVal ComputerName As String)
    Dim oShell, oExec As Variant
    Dim strText, strCmd As String
    strText = ""
    strCmd = "ping -n 3 -w 1000 " & ComputerName
    Set oShell = CreateObject("WScript.Shell")
    Set oExec = oShell.Exec(strCmd)
    Do While Not oExec.StdOut.AtEndOfStream
        strText = oExec.StdOut.ReadLine()
        If (InStr(strText, "Reply") > 0) And (InStr(strText, "unreachable") < 1) Then
            SystemOnline = True
            Exit Do
        End If
    Loop
End Function

'Determine if process is running
Function ProcessRunning(ByVal ProcessName As String)
    Dim oShell, oExec As Variant
    Dim strText, strCmd As String
    strText = ""
    strCmd = "ps aux | grep " & QUOTATION & ProcessName & QUOTATION
    Set oShell = CreateObject("WScript.Shell")
    Set oExec = oShell.Exec(strCmd)
    Do While Not oExec.StdOut.AtEndOfStream
        strText = oExec.StdOut.ReadLine()
        If (InStr(strText, "S+") > 0) Then
            ProcessRunning = True
            Exit Do
        End If
    Loop
End Function
Function PopulateWorksheetArray(xCelSheet As Excel.Worksheet, lNumRows As Long, lNumColumns As Long)
    Dim varrWorksheet() As Variant
    Dim lRowCounter, lArrayRowCounter, lArrayColumnCounter As Long
    
    'Redimension array to hold contents of worksheet
    ReDim varrWorksheet(1 To lNumRows, 1 To lNumColumns)
    
    'Populate array from worksheet
    lRowCounter = 2
    lArrayRowCounter = 1
    
    Do While Len(Trim(xCelSheet.Cells(lRowCounter, 1))) > 0
        
        For lArrayColumnCounter = 1 To UBound(varrWorksheet, 2)
            varrWorksheet(lArrayRowCounter, lArrayColumnCounter) = Trim(xCelSheet.Cells(lRowCounter, lArrayColumnCounter))
        Next lArrayColumnCounter

        lArrayRowCounter = lArrayRowCounter + 1
        lRowCounter = lRowCounter + 1
    Loop
    
    PopulateWorksheetArray = varrWorksheet
End Function

Public Sub ExportSourceFiles(destPath As String)
 
    Dim component As VBComponent
    For Each component In Application.VBE.ActiveVBProject.VBComponents
        If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Or component.Type = 100 Then
            component.Export destPath & component.Name & ToFileExtension(component.Type)
        End If
    Next
    'Application.VBE.ActiveVBProject.VBComponents
    'Application.ActiveWorkbook.VBProject.VBComponents
End Sub
 
Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
    Select Case vbeComponentType
        Case 100
            ToFileExtension = ".cls"
        Case vbext_ComponentType.vbext_ct_ClassModule
            ToFileExtension = ".cls"
        Case vbext_ComponentType.vbext_ct_StdModule
            ToFileExtension = ".bas"
        Case vbext_ComponentType.vbext_ct_MSForm
            ToFileExtension = ".frm"
        Case vbext_ComponentType.vbext_ct_ActiveXDesigner
        Case vbext_ComponentType.vbext_ct_Document
        Case Else
            ToFileExtension = vbNullString
    End Select
 
End Function

Private Function CommonDataValidation(xCelSheet As Excel.Worksheet) As Boolean
    Dim lRowCounter As Long: lRowCounter = 2
    Dim lRow, lColumn, lRowsInterval As Long
    Dim sDateStart, sDateEnd, sTimeStart, sTimeEnd As String
    Dim lNumberOfNonEmptyRows As Long
    
    lNumberOfNonEmptyRows = CountNonEmptyRows(xCelSheet, NUMBER_OF_COLUMNS)
    
    'IF THERE ARE FEWER THAN 2 POPULATED ROWS, THROW ERROR
    If lNumberOfNonEmptyRows < 3 Then
    
        MsgBox "Please enter at least two rows.", vbExclamation, "Data Entry Error"
        CommonDataValidation = False
        Exit Function
    End If
    
    'IF ANY ROWS ARE MISSING VALUES, THROW ERROR
        
    'Check all populated rows for complete data and throw error if any rows have incomplete data
    For lRow = 2 To lNumberOfNonEmptyRows
        'Count number of columns populated
        For lColumn = 1 To NUMBER_OF_COLUMNS
            If Len(Trim(xCelSheet.Cells(lRow, lColumn))) = 0 Then
        
                MsgBox "Please fill in missing data on row " & lRow & ".", vbExclamation, "Data Entry Error"
                CommonDataValidation = False
                Exit Function
            End If
        Next lColumn
    Next lRow
    
    'IF TIME INTERVAL BETWEEN ANY ROW AND PREVIOUS ROW IS LESS THAN OR EQUAL TO ZERO, THROW ERROR
    
    Do While Len(Trim(xCelSheet.Cells(lRowCounter, 1))) > 0
    
        If Len(Trim(xCelSheet.Cells(lRowCounter + 1, 1))) > 0 Then
            'Format start and end date of rows and convert to string
            sDateStart = Format(Trim(xCelSheet.Cells(lRowCounter, 1)), DATE_FORMATTING_STRING)
            sDateEnd = Format(Trim(xCelSheet.Cells(lRowCounter + 1, 1)), DATE_FORMATTING_STRING)

            'Format start and end times of rows and convert to string
            sTimeStart = Format(TimeSerial(xCelSheet.Cells(lRowCounter, 2), xCelSheet.Cells(lRowCounter, 3), xCelSheet.Cells(lRowCounter, 4)), TIME_FORMATTING_STRING)
            sTimeEnd = Format(TimeSerial(xCelSheet.Cells(lRowCounter + 1, 2), xCelSheet.Cells(lRowCounter + 1, 3), xCelSheet.Cells(lRowCounter + 1, 4)), TIME_FORMATTING_STRING)
            
            lRowsInterval = DateDiff("s", CDate(sDateStart & " " & sTimeStart), CDate(sDateEnd & " " & sTimeEnd))
            
            If lRowsInterval <= 0 Then
        
                MsgBox "Please make date/time on row " & (lRowCounter + 1) & " greater than row " & lRowCounter & ".", vbExclamation, "Data Entry Error"
                CommonDataValidation = False
                Exit Function
            End If
        End If
        
        lRowCounter = lRowCounter + 1
    Loop
    
    CommonDataValidation = True
End Function
