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
Const REPEAT_PATTERN_ROW_NUMBER As Integer = 3
Const TIME_FROM_LAST_ROW_ROW_NUMBER As Integer = 4
Const TIME_BETWEEN_REPEATS_ROW_NUMBER As Integer = 5
Const FREQUENCY_UPPER_BOUND As Double = 4
Const FREQUENCY_LOWER_BOUND As Double = 3
Const CHAOS_EXPERIMENT_DURATION_ROW_NUMBER As Integer = 14
Const CHAOS_TIME_FROM_LAST_ROW_ROW_NUMBER As Integer = 15
Const CHAOS_PHOTO_PERIOD_ROW_NUMBER As Integer = 16
Const CHAOS_DARK_PERIOD_ROW_NUMBER As Integer = 17
Const CHAOS_X0_ROW_NUMBER As Integer = 18
Const CHAOS_R_ROW_NUMBER As Integer = 19
Const CHAOS_MD1_ROW_NUMBER As Integer = 20
Const CHAOS_MD2_ROW_NUMBER As Integer = 21
Const CHAOS_START_DATETIME_ROW_NUMBER As Integer = 22
Const CHAOS_ARRAY_NUM_COLUMNS As Integer = 18
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
Const X0_UBOUND As Double = 1
Const X0_LBOUND As Double = 0
Const R_UBOUND As Double = 4
Const R_LBOUND As Double = 3.9
Const MD1_UBOUND As Double = 0.5
Const MD1_LBOUND As Double = 0
Const MD2_UBOUND As Double = 0.5
Const MD2_LBOUND As Double = 0
'*** NOTE: FUTURE VERSION SHOULD ALLOW USER ENTRY OF FOLLOWING CONSTANTS *********************
Const CH1_RATIO As Double = 0
Const CH2_RATIO As Double = 0
Const CH3_RATIO As Double = 20
Const CH4_RATIO As Double = 20
Const CH5_RATIO As Double = 60
Const CH6_RATIO As Double = 12
Const MIN_PERCENT1 As Double = 0
Const MIN_PERCENT2 As Double = 0
Const MIN_PERCENT3 As Double = 0
Const MIN_PERCENT4 As Double = 0
Const MIN_PERCENT5 As Double = 0
Const MIN_PERCENT6 As Double = 0
Const TOTAL_OUTPUT As Double = 14000000 'micromol photons/m2/day (or photoperiod)
Const CHAOS_BASE_FUNCTION As Integer = 0 '0 = flat line, 1 = sine wave
'*********************************************************************************************'
Const CHAOS_ROUNDING_DIGITS As Integer = 6
Const CHAOS_MAX_SWITCH_TIME As Integer = 180 'Maximum switching time in seconds
Const CHAOS_MIN_SWITCH_TIME As Integer = 30 'Minimum switching time in seconds

Const RASP_PI_INTERFACE_NAME As String = "HortiLight_v1.1.py"
Const RUNLIGHTCOMMAND_FILE_NAME As String = "RunLightCommand_v1.1.py"

'------------------------------------------------------------------------------------------------------------
'Sub: WriteToOutputPattern
'Coded by: Matt Urschel
'Date : 3 May 2017
'Description: Code for button "Write To Output" on Input worksheet - Appends user-entered rows on Input
'             page to end of data on Output page, with start time after a user-entered interval since
'             last line. Repeats pattern for user-specified time interval (times are automatically advanced).
'------------------------------------------------------------------------------------------------------------
Public Sub WriteToOutputPattern()
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
    If Len(Trim(XCelSheet1.Cells(REPEAT_PATTERN_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))) > 0 Then
       lRepeatInterval = CLng(XCelSheet1.Cells(REPEAT_PATTERN_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))
    End If
    
    'Get contents of unit cell if changed
    If Len(Trim(XCelSheet1.Cells(REPEAT_PATTERN_ROW_NUMBER, REPEAT_UNITS_CELL_NUMBER))) > 0 Then
       sRepeatUnit = Trim(XCelSheet1.Cells(REPEAT_PATTERN_ROW_NUMBER, REPEAT_UNITS_CELL_NUMBER))
       
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
    If Len(Trim(XCelSheet1.Cells(TIME_FROM_LAST_ROW_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))) > 0 Then
       lTimeAfterLastRowInterval = CLng(XCelSheet1.Cells(TIME_FROM_LAST_ROW_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))
    End If
    
    'Get contents of time after last row unit cell if changed
    If Len(Trim(XCelSheet1.Cells(TIME_FROM_LAST_ROW_ROW_NUMBER, REPEAT_UNITS_CELL_NUMBER))) > 0 Then
       sTimeAfterLastRowUnit = Trim(XCelSheet1.Cells(TIME_FROM_LAST_ROW_ROW_NUMBER, REPEAT_UNITS_CELL_NUMBER))
       
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
    If Len(Trim(XCelSheet1.Cells(TIME_BETWEEN_REPEATS_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))) > 0 Then
       lTimeBetweenRepeatsInterval = CLng(XCelSheet1.Cells(TIME_BETWEEN_REPEATS_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))
    End If
    
    'Get contents of time between repeats unit cell if changed
    If Len(Trim(XCelSheet1.Cells(TIME_BETWEEN_REPEATS_ROW_NUMBER, REPEAT_UNITS_CELL_NUMBER))) > 0 Then
       sTimeBetweenRepeatsUnit = Trim(XCelSheet1.Cells(TIME_BETWEEN_REPEATS_ROW_NUMBER, REPEAT_UNITS_CELL_NUMBER))
       
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
Public Sub WriteToOutputChaos()
    Dim XCelWorkbook As Excel.Workbook
    Dim XCelSheet1 As Excel.Worksheet
    Dim XCelSheet2 As Excel.Worksheet
    Dim lRowCounter1 As Long: lRowCounter1 = 2
    Dim lRowCounter2 As Long: lRowCounter2 = 2
    Dim lNumberOfNonEmptyRowsSheet2 As Long: lNumberOfNonEmptyRowsSheet2 = 0
    Dim lFirstBlankRow As Long
    Dim vArrayChaos(), vArrayChaosTemp() As Variant
    Dim dExperimentDuration As Double: dExperimentDuration = 0
    Dim sExperimentDurationUnits As String
    Dim dTimeFromLastRow As Double: dTimeFromLastRow = 0
    Dim sTimeFromLastRowUnits As String
    Dim dPhotoPeriod As Double: dPhotoPeriod = 0
    Dim sPhotoPeriodUnits As String
    Dim dDarkPeriod As Double: dDarkPeriod = 0
    Dim sDarkPeriodUnits As String
    Dim lNumberOfRepetitions, lRepeat As Long
    'Dim vArrayChaos(iArrayRowCounter, 5) As Double: vArrayChaos(iArrayRowCounter, 5) = 0
    Dim iTotalTime, iSwitchTime As Integer: iTotalTime = 0: iSwitchTime = 0
    Dim iArrayChaosRows As Integer: iArrayChaosRows = 1
    Dim dX0 As Double: dX0 = 0
    Dim dR As Double: dR = 0
    Dim dMD1, dMD2 As Double: dMD1 = 0: dMD2 = 0
    Dim sDateTime, sDate, sTime, sHH, sMM, sSS, sLine As String
    Dim bX0Random, bRRandom, bMD1Random, bMD2Random As Boolean: bX0Random = False: bRRandom = False: bMD1Random = False: bMD2Random = False
    Dim iArrayRowCounter, iArrayColumnCounter, iTempArrayCounterRows, iTempArrayCounterColumns As Integer: iArrayRowCounter = 0: iArrayColumnCounter = 0: iTempArrayCounterRows = 0: iTempArrayCounterColumns = 0
    Dim Pi As Double: Pi = 4 * Atn(1)
    Dim iLogFileNum As Integer: iLogFileNum = FreeFile
    Dim sLogFileName As String: sLogFileName = Application.ActiveWorkbook.Path & "\Chaos_log.txt"
    Dim dCH1, dCH2, dCH3, dCH4, dCH5, dCH6 As Double
    Dim dCH1Output, dCH2Output, dCH3Output, dCH4Output, dCH5Output, dCH6Output As Double
    Dim dTotalOutput, dTotalOutputCH1, dTotalOutputCH2, dTotalOutputCH3, dTotalOutputCH4, dTotalOutputCH5, dTotalOutputCH6, dDesiredTotalOutput As Double
    Dim dAdjustedTotalOutput, dAdjustedTotalOutputCH1, dAdjustedTotalOutputCH2, dAdjustedTotalOutputCH3, dAdjustedTotalOutputCH4, dAdjustedTotalOutputCH5, dAdjustedTotalOutputCH6 As Double
    
    '---------------------
    'OBJECT INITIALIZATION
    '---------------------
    
    'Initialize workbook and worksheets
    Set XCelWorkbook = Application.ActiveWorkbook
    Set XCelSheet1 = XCelWorkbook.Sheets(1)
    Set XCelSheet2 = XCelWorkbook.Sheets(2)
    
    Open sLogFileName For Output As iLogFileNum
    
    '----------------------------------
    'DATA VALIDATION AND INITIALIZATION
    '----------------------------------
    
    'Experiment duration
    If Len(Trim(XCelSheet1.Cells(CHAOS_EXPERIMENT_DURATION_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))) < 1 Then
        MsgBox "Please enter experiment duration.", vbExclamation, "Data Entry Error"
        Exit Sub
    Else
        dExperimentDuration = CLng(XCelSheet1.Cells(CHAOS_EXPERIMENT_DURATION_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))
        sExperimentDurationUnits = XCelSheet1.Cells(CHAOS_EXPERIMENT_DURATION_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER + 1)
        
        'Convert experiment duration to seconds
        Select Case sExperimentDurationUnits
            Case "Weeks"
                dExperimentDuration = dExperimentDuration * 604800
            Case "Days"
                dExperimentDuration = dExperimentDuration * 86400
            Case "Hours"
                dExperimentDuration = dExperimentDuration * 3600
            Case "Minutes"
                dExperimentDuration = dExperimentDuration * 60
            Case "Repeats"
                lNumberOfRepetitions = dExperimentDuration
        End Select
    End If
    
    'Photoperiod
    If Len(Trim(XCelSheet1.Cells(CHAOS_PHOTO_PERIOD_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))) < 1 Then
        MsgBox "Please enter photoperiod.", vbExclamation, "Data Entry Error"
        Exit Sub
    Else
        dPhotoPeriod = CLng(XCelSheet1.Cells(CHAOS_PHOTO_PERIOD_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))
        sPhotoPeriodUnits = XCelSheet1.Cells(CHAOS_PHOTO_PERIOD_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER + 1)
        'Convert photo period to seconds
        Select Case sPhotoPeriodUnits
            Case "Weeks"
                dPhotoPeriod = dPhotoPeriod * 604800
            Case "Days"
                dPhotoPeriod = dPhotoPeriod * 86400
            Case "Hours"
                dPhotoPeriod = dPhotoPeriod * 3600
            Case "Minutes"
                dPhotoPeriod = dPhotoPeriod * 60
        End Select
    End If
    
    'Dark period
    If Len(Trim(XCelSheet1.Cells(CHAOS_DARK_PERIOD_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))) < 1 Then
        MsgBox "Please enter dark period.", vbExclamation, "Data Entry Error"
        Exit Sub
    Else
        dDarkPeriod = CLng(XCelSheet1.Cells(CHAOS_DARK_PERIOD_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))
        sDarkPeriodUnits = XCelSheet1.Cells(CHAOS_PHOTO_PERIOD_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER + 1)
        'Convert dark period to seconds
        Select Case sDarkPeriodUnits
            Case "Weeks"
                dDarkPeriod = dDarkPeriod * 604800
            Case "Days"
                dDarkPeriod = dDarkPeriod * 86400
            Case "Hours"
                dDarkPeriod = dDarkPeriod * 3600
            Case "Minutes"
                dDarkPeriod = dDarkPeriod * 60
        End Select
    End If
    
    'If experiment duration is shorter than the sum of the photo period and dark period, throw error
    If dExperimentDuration < (dPhotoPeriod + dDarkPeriod) Then
        MsgBox "Sum of photo period and dark period must be less than total experiment duration.", vbExclamation, "Data Entry Error"
        Exit Sub
    End If
    
    'Determine total number of photo/dark periods that need to be generated (if experiment duration units were not "Repeats")
    If sExperimentDurationUnits <> "Repeats" Then
        lNumberOfRepetitions = CLng(dExperimentDuration / (dPhotoPeriod + dDarkPeriod))
    End If
    
    'X0
    If Len(Trim(XCelSheet1.Cells(CHAOS_X0_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))) < 1 Then
        bX0Random = True
    Else
        dX0 = CDbl(XCelSheet1.Cells(CHAOS_X0_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))
    End If
    
    'R
    If Len(Trim(XCelSheet1.Cells(CHAOS_R_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))) < 1 Then
        bRRandom = True
    Else
        dR = CDbl(XCelSheet1.Cells(CHAOS_R_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))
    End If
    
    'MD1
    If Len(Trim(XCelSheet1.Cells(CHAOS_MD1_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))) < 1 Then
        bMD1Random = True
    Else
        dMD1 = CDbl(XCelSheet1.Cells(CHAOS_MD1_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))
    End If
    
    'MD2
    If Len(Trim(XCelSheet1.Cells(CHAOS_MD2_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))) < 1 Then
        bMD2Random = True
    Else
        dMD2 = CDbl(XCelSheet1.Cells(CHAOS_MD2_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))
    End If
    
    
    'Set frequency based on photoperiod
    'vArrayChaos(iArrayRowCounter, 5) = CDbl(dPhotoPeriod / 300)
    
    'Determine last populated row on worksheet 2
    lNumberOfNonEmptyRowsSheet2 = CountNonEmptyRows(XCelSheet2, NUMBER_OF_COLUMNS)
    
    'Set worksheet 2 row counter to first empty row on worksheet 2
    If lNumberOfNonEmptyRowsSheet2 >= 2 Then
        lRowCounter2 = lNumberOfNonEmptyRowsSheet2 + 1
    End If
    
    'If there are rows on output worksheet, set start date for chaos commands to that date.
    'Otherwise, set it to user entered date in chaos section of input worksheet.
    If lNumberOfNonEmptyRowsSheet2 > 1 Then
    
        If (Len(Trim(XCelSheet1.Cells(CHAOS_TIME_FROM_LAST_ROW_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER))) < 1) Or _
           (Len(Trim(XCelSheet1.Cells(CHAOS_TIME_FROM_LAST_ROW_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER + 1))) < 1) Then
            
            MsgBox "Please enter interval and units for time after last row.", vbExclamation, "Data Entry Error"
            Exit Sub
            
        Else
            dTimeFromLastRow = CDbl(Trim(XCelSheet1.Cells(CHAOS_TIME_FROM_LAST_ROW_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER)))
            sTimeFromLastRowUnits = Trim(XCelSheet1.Cells(CHAOS_TIME_FROM_LAST_ROW_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER + 1))
            
            Select Case sTimeFromLastRowUnits
                Case "Weeks"
                    dTimeFromLastRow = dTimeFromLastRow * 604800
                Case "Days"
                    dTimeFromLastRow = dTimeFromLastRow * 86400
                Case "Hours"
                    dTimeFromLastRow = dTimeFromLastRow * 3600
                Case "Minutes"
                    dTimeFromLastRow = dTimeFromLastRow * 60
            End Select
            
            'Get start date/time from last row on output worksheet
            sDateTime = Format(XCelSheet2.Cells(lNumberOfNonEmptyRowsSheet2, 1) & " " & _
                         XCelSheet2.Cells(lNumberOfNonEmptyRowsSheet2, 2) & ":" & _
                         XCelSheet2.Cells(lNumberOfNonEmptyRowsSheet2, 3) & ":" & _
                         XCelSheet2.Cells(lNumberOfNonEmptyRowsSheet2, 4), DATE_FORMATTING_STRING & " " & TIME_FORMATTING_STRING)
                         
            'Advance start date/time by user entered time from last row
            sDateTime = CStr(DateAdd("s", dTimeFromLastRow, CDate(sDateTime)))
        End If
    Else
        sDateTime = Format(XCelSheet1.Cells(CHAOS_START_DATETIME_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER), DATE_FORMATTING_STRING & " " & TIME_FORMATTING_STRING)
    End If
    
    For lRepeat = 1 To lNumberOfRepetitions
    
        'Initialize counters for total time and number of rows in chaos array
        iTotalTime = 0
        iArrayChaosRows = 1
        
        'Initialize chaos array
        ReDim vArrayChaos(1 To 1, 1 To CHAOS_ARRAY_NUM_COLUMNS)
        
        'Generate random switching times until photoperiod time reached or exceeded
        Do Until iTotalTime >= dPhotoPeriod
        
            'Generate random switching time for this step
            iSwitchTime = Int((CHAOS_MAX_SWITCH_TIME - CHAOS_MIN_SWITCH_TIME + 1) * Rnd + CHAOS_MIN_SWITCH_TIME)
            
            'Add switching time to total time
            iTotalTime = iTotalTime + iSwitchTime
            
            If iTotalTime > dPhotoPeriod Then
               iTotalTime = dPhotoPeriod
            End If
            
            'Populate chaos array step time with total time plus random switching time
            vArrayChaos(iArrayChaosRows, 5) = iTotalTime
            
            'Increment number of rows for chaos array
            iArrayChaosRows = iArrayChaosRows + 1
            
            'Put values in chaos array in new temporary array
            vArrayChaosTemp = vArrayChaos
            
            'Redimension chaos array to add row for current repeat
            ReDim vArrayChaos(1 To iArrayChaosRows, 1 To CHAOS_ARRAY_NUM_COLUMNS)
            
            'Copy values from temp array into redimensioned chaos array
            For iTempArrayCounterRows = 1 To UBound(vArrayChaosTemp, 1)
                For iTempArrayCounterColumns = 1 To UBound(vArrayChaosTemp, 2)
                    vArrayChaos(iTempArrayCounterRows, iTempArrayCounterColumns) = vArrayChaosTemp(iTempArrayCounterRows, iTempArrayCounterColumns)
                Next iTempArrayCounterColumns
            Next iTempArrayCounterRows
            
            
            
        Loop
        
        
        'Set total output for all channels to zero
        dTotalOutput = 0
        dTotalOutputCH1 = 0
        dTotalOutputCH2 = 0
        dTotalOutputCH3 = 0
        dTotalOutputCH4 = 0
        dTotalOutputCH5 = 0
        dTotalOutputCH6 = 0
        
        dAdjustedTotalOutputCH1 = 0
        dAdjustedTotalOutputCH2 = 0
        dAdjustedTotalOutputCH3 = 0
        dAdjustedTotalOutputCH4 = 0
        dAdjustedTotalOutputCH5 = 0
        dAdjustedTotalOutputCH6 = 0
        dAdjustedTotalOutput = 0
        
        
        'Randomize variables (if specified by user)
        'X0
        If bX0Random Then
            Randomize
            dX0 = (Rnd * (X0_UBOUND - X0_LBOUND)) + X0_LBOUND
        End If
        
        'r
        If bRRandom Then
            Randomize
            dR = (Rnd * (R_UBOUND - R_LBOUND)) + R_LBOUND
        End If
        
        'MD1
        If bMD1Random Then
            Randomize
            dMD1 = (Rnd * (MD1_UBOUND - MD1_LBOUND)) + MD1_LBOUND
        End If
        
        'MD2
        If bMD2Random Then
            Randomize
            dMD2 = (Rnd * (MD2_UBOUND - MD2_LBOUND)) + MD2_LBOUND
        End If
        
        For iArrayColumnCounter = 1 To 4
            For iArrayRowCounter = 1 To iArrayChaosRows - 1
                Select Case iArrayColumnCounter
                    Case 1
                        vArrayChaos(iArrayRowCounter, iArrayColumnCounter) = iArrayRowCounter
                    Case 2
                        If iArrayRowCounter = 1 Then
                            vArrayChaos(iArrayRowCounter, iArrayColumnCounter) = Application.WorksheetFunction.RoundUp((dR * dX0 * (1 - dX0)), CHAOS_ROUNDING_DIGITS)
                        Else
                            vArrayChaos(iArrayRowCounter, iArrayColumnCounter) = Application.WorksheetFunction.RoundUp(dR * vArrayChaos(iArrayRowCounter - 1, iArrayColumnCounter) * (1 - vArrayChaos(iArrayRowCounter - 1, iArrayColumnCounter)), CHAOS_ROUNDING_DIGITS)
                        End If
                    Case 3
                        vArrayChaos(iArrayRowCounter, iArrayColumnCounter) = Sin(vArrayChaos(iArrayRowCounter, 1) * Pi / iArrayChaosRows)
                    Case 4
                        vArrayChaos(iArrayRowCounter, iArrayColumnCounter) = vArrayChaos(iArrayRowCounter, 3) * (1 - vArrayChaos(iArrayRowCounter, 2) * dMD1)
                    'Case 5
                        'If iArrayRowCounter > 1 Then
                            'vArrayChaos(iArrayRowCounter, iArrayColumnCounter) = CDbl(vArrayChaos(iArrayRowCounter - 1, 5) + vArrayChaos(iArrayRowCounter, 5))
                        'Else
                            'vArrayChaos(iArrayRowCounter, iArrayColumnCounter) = iArrayRowCounter
                        'End If
                End Select
            
                
                
            Next iArrayRowCounter
        Next iArrayColumnCounter
        
        
        For iArrayColumnCounter = 6 To 8
            For iArrayRowCounter = 1 To iArrayChaosRows - 1
                Select Case iArrayColumnCounter
                    Case 6
                        If Int(vArrayChaos(iArrayRowCounter, 5)) <= iArrayChaosRows Then
                            vArrayChaos(iArrayRowCounter, iArrayColumnCounter) = vArrayChaos(iArrayRowCounter, 3) * (1 - vArrayChaos(Int(vArrayChaos(iArrayRowCounter, 5)), 2) * dMD2)
                        Else
                            vArrayChaos(iArrayRowCounter, iArrayColumnCounter) = vArrayChaos(iArrayRowCounter, 3) * (1 - vArrayChaos(iArrayChaosRows, 2) * dMD2)
                        End If
                    Case 7
                        vArrayChaos(iArrayRowCounter, iArrayColumnCounter) = (vArrayChaos(iArrayRowCounter, 6) + vArrayChaos(iArrayRowCounter, 4)) / 2
                    Case 8
                        If CHAOS_BASE_FUNCTION = 1 Then
                            vArrayChaos(iArrayRowCounter, iArrayColumnCounter) = vArrayChaos(iArrayRowCounter, 7) * 100
                        End If
                        
                        If CHAOS_BASE_FUNCTION = 0 Then
                            vArrayChaos(iArrayRowCounter, iArrayColumnCounter) = vArrayChaos(iArrayRowCounter, 2) * 100
                        End If
                End Select
            Next iArrayRowCounter
        Next iArrayColumnCounter
        
        'Output chaos calculation data to log file
        'Heading
        sLine = "Photoperiod " & lRepeat & " (r=" & dR & ", X0=" & dX0 & ", MD1=" & dMD1 & ", MD2=" & dMD2 & ", F=" & vArrayChaos(iArrayRowCounter, 5) & ")"
        Print #iLogFileNum, sLine
        sLine = "Theta,Chaos,Sin(Theta),Damp,Time,Time&MaxDamp,FinalChaos,FinalChaos*100%"
        Print #iLogFileNum, sLine
        sLine = ""
        
        'Array
        For iArrayRowCounter = 1 To iArrayChaosRows - 1
            sLine = ""
            For iArrayColumnCounter = 1 To 8
                sLine = sLine & vArrayChaos(iArrayRowCounter, iArrayColumnCounter) & ","
            Next iArrayColumnCounter
            Print #iLogFileNum, sLine
        Next iArrayRowCounter
        
        'Populate chaos array with worksheet output values
        For iArrayRowCounter = 1 To iArrayChaosRows - 1
            'Date and Time
            If iArrayRowCounter > 1 Then
                sDateTime = CStr(DateAdd("s", CDbl(vArrayChaos(iArrayRowCounter, 5) - vArrayChaos(iArrayRowCounter - 1, 5)), CDate(sDateTime)))
            End If
            
            sDate = Format(sDateTime, DATE_FORMATTING_STRING)
            sTime = Format(sDateTime, TIME_FORMATTING_STRING)
            
            sHH = Left(sTime, InStr(sTime, ":") - 1)
            sMM = Mid(sTime, InStr(sTime, ":") + 1, 2)
            sSS = Right(sTime, 2)
            
            'Channel 1 %
            dCH1 = CDbl(Application.WorksheetFunction.RoundUp(vArrayChaos(iArrayRowCounter, 8) * (CH1_RATIO / _
                   Application.WorksheetFunction.Max(CH1_RATIO, CH2_RATIO, CH3_RATIO, CH4_RATIO, CH5_RATIO, CH6_RATIO)), 2))
            'Channel 2 %
            dCH2 = CDbl(Application.WorksheetFunction.RoundUp(vArrayChaos(iArrayRowCounter, 8) * (CH2_RATIO / _
                   Application.WorksheetFunction.Max(CH1_RATIO, CH2_RATIO, CH3_RATIO, CH4_RATIO, CH5_RATIO, CH6_RATIO)), 2))
            'Channel 3 %
            dCH3 = CDbl(Application.WorksheetFunction.RoundUp(vArrayChaos(iArrayRowCounter, 8) * (CH3_RATIO / _
                   Application.WorksheetFunction.Max(CH1_RATIO, CH2_RATIO, CH3_RATIO, CH4_RATIO, CH5_RATIO, CH6_RATIO)), 2))
            'Channel 4 %
            dCH4 = CDbl(Application.WorksheetFunction.RoundUp(vArrayChaos(iArrayRowCounter, 8) * (CH4_RATIO / _
                   Application.WorksheetFunction.Max(CH1_RATIO, CH2_RATIO, CH3_RATIO, CH4_RATIO, CH5_RATIO, CH6_RATIO)), 2))
            'Channel 5 %
            dCH5 = CDbl(Application.WorksheetFunction.RoundUp(vArrayChaos(iArrayRowCounter, 8) * (CH5_RATIO / _
                   Application.WorksheetFunction.Max(CH1_RATIO, CH2_RATIO, CH3_RATIO, CH4_RATIO, CH5_RATIO, CH6_RATIO)), 2))
            'Channel 6 %
            dCH6 = CDbl(Application.WorksheetFunction.RoundUp(vArrayChaos(iArrayRowCounter, 8) * (CH6_RATIO / _
                   Application.WorksheetFunction.Max(CH1_RATIO, CH2_RATIO, CH3_RATIO, CH4_RATIO, CH5_RATIO, CH6_RATIO)), 2))
            
            'Adjust channel % for max photons
            
            If iArrayRowCounter > 1 Then
                'Calculate output photons of each channel (micromol photons / m^2)
                dCH1Output = (1.9726 * dCH1 - 16.916) * (vArrayChaos(iArrayRowCounter, 5) - vArrayChaos(iArrayRowCounter - 1, 5))
                dCH2Output = (2.4582 * dCH2 - 15.204) * (vArrayChaos(iArrayRowCounter, 5) - vArrayChaos(iArrayRowCounter - 1, 5))
    '            dCH3Output = (3.1604 * dCH3 - 8.6614) * vArrayChaos(iArrayRowCounter, 5)
    '            dCH4Output = ((-0.0076 * (dCH4 ^ 2)) + (2.2092 * dCH4) - 16.318) * vArrayChaos(iArrayRowCounter, 5)
    '            dCH5Output = (3.807 * dCH5 - 32.702) * vArrayChaos(iArrayRowCounter, 5)
    '            dCH6Output = (2.4158 * dCH6 - 17.079) * vArrayChaos(iArrayRowCounter, 5)
                
                dCH3Output = (-0.0052 * (dCH3 ^ 2) + (3.0623 * dCH3) - 16.067) * (vArrayChaos(iArrayRowCounter, 5) - vArrayChaos(iArrayRowCounter - 1, 5))
                dCH4Output = ((9 * (10 ^ -6)) * (dCH4 ^ 3) - (0.0058 * (dCH4 ^ 2)) + (1.5671 * dCH4) - 3.7667) * (vArrayChaos(iArrayRowCounter, 5) - vArrayChaos(iArrayRowCounter - 1, 5))
                dCH5Output = ((-5 * (10 ^ -5)) * (dCH5 ^ 3) + (0.0045 * (dCH5 * 2)) + (2.8477 * dCH5) - 12.603) * (vArrayChaos(iArrayRowCounter, 5) - vArrayChaos(iArrayRowCounter - 1, 5))
                dCH6Output = ((-4 * (10 ^ -5)) * (dCH6 ^ 3) + (0.003 * (dCH6 ^ 2)) + (1.9106 * dCH6) - 9.6238) * (vArrayChaos(iArrayRowCounter, 5) - vArrayChaos(iArrayRowCounter - 1, 5))
            Else
            
                dCH1Output = (1.9726 * dCH1 - 16.916) * vArrayChaos(iArrayRowCounter, 5)
                dCH2Output = (2.4582 * dCH2 - 15.204) * vArrayChaos(iArrayRowCounter, 5)
    '            dCH3Output = (3.1604 * dCH3 - 8.6614) * vArrayChaos(iArrayRowCounter, 5)
    '            dCH4Output = ((-0.0076 * (dCH4 ^ 2)) + (2.2092 * dCH4) - 16.318) * vArrayChaos(iArrayRowCounter, 5)
    '            dCH5Output = (3.807 * dCH5 - 32.702) * vArrayChaos(iArrayRowCounter, 5)
    '            dCH6Output = (2.4158 * dCH6 - 17.079) * vArrayChaos(iArrayRowCounter, 5)
                
                dCH3Output = (-0.0052 * (dCH3 ^ 2) + (3.0623 * dCH3) - 16.067) * vArrayChaos(iArrayRowCounter, 5)
                dCH4Output = ((9 * (10 ^ -6)) * (dCH4 ^ 3) - (0.0058 * (dCH4 ^ 2)) + (1.5671 * dCH4) - 3.7667) * vArrayChaos(iArrayRowCounter, 5)
                dCH5Output = ((-5 * (10 ^ -5)) * (dCH5 ^ 3) + (0.0045 * (dCH5 * 2)) + (2.8477 * dCH5) - 12.603) * vArrayChaos(iArrayRowCounter, 5)
                dCH6Output = ((-4 * (10 ^ -5)) * (dCH6 ^ 3) + (0.003 * (dCH6 ^ 2)) + (1.9106 * dCH6) - 9.6238) * vArrayChaos(iArrayRowCounter, 5)

            End If
                
            If dCH1Output < 0 Then
                dCH1Output = 0
            End If
            
            If dCH2Output < 0 Then
                dCH2Output = 0
            End If
            
            If dCH3Output < 0 Then
                dCH3Output = 0
            End If
            
            If dCH4Output < 0 Then
                dCH4Output = 0
            End If
            
            If dCH5Output < 0 Then
                dCH5Output = 0
            End If
            
            If dCH6Output < 0 Then
                dCH6Output = 0
            End If
            
            'Add output from this row to total output for all channels, all rows
            dTotalOutput = dTotalOutput + dCH1Output + dCH2Output + dCH3Output + dCH4Output + dCH5Output + dCH6Output
            dTotalOutputCH1 = dTotalOutputCH1 + dCH1Output
            dTotalOutputCH2 = dTotalOutputCH2 + dCH2Output
            dTotalOutputCH3 = dTotalOutputCH3 + dCH3Output
            dTotalOutputCH4 = dTotalOutputCH4 + dCH4Output
            dTotalOutputCH5 = dTotalOutputCH5 + dCH5Output
            dTotalOutputCH6 = dTotalOutputCH6 + dCH6Output
            
            'Set output date/time and uncorrected channel %
            vArrayChaos(iArrayRowCounter, 9) = sDate
            vArrayChaos(iArrayRowCounter, 10) = sHH
            vArrayChaos(iArrayRowCounter, 11) = sMM
            vArrayChaos(iArrayRowCounter, 12) = sSS
            vArrayChaos(iArrayRowCounter, 13) = dCH1
            vArrayChaos(iArrayRowCounter, 14) = dCH2
            vArrayChaos(iArrayRowCounter, 15) = dCH3
            vArrayChaos(iArrayRowCounter, 16) = dCH4
            vArrayChaos(iArrayRowCounter, 17) = dCH5
            vArrayChaos(iArrayRowCounter, 18) = dCH6
            
        Next iArrayRowCounter
        
        'Write worksheet values to output worksheet
        For iArrayRowCounter = 1 To iArrayChaosRows - 1
        
            '((Application.WorksheetFunction.Max(MIN_PERCENT1,MIN_PERCENT2,MIN_PERCENT3,MIN_PERCENT4,MIN_PERCENT5,MIN_PERCENT6)/100) * TOTAL_OUTPUT)
            dDesiredTotalOutput = TOTAL_OUTPUT - ((Application.WorksheetFunction.Sum(MIN_PERCENT1, MIN_PERCENT2, MIN_PERCENT3, MIN_PERCENT4, MIN_PERCENT5, MIN_PERCENT6) / 100) * TOTAL_OUTPUT)
        
            'Adjust channel percentages based on desired output
            vArrayChaos(iArrayRowCounter, 13) = vArrayChaos(iArrayRowCounter, 13) * dDesiredTotalOutput / dTotalOutput
            vArrayChaos(iArrayRowCounter, 14) = vArrayChaos(iArrayRowCounter, 14) * dDesiredTotalOutput / dTotalOutput
            vArrayChaos(iArrayRowCounter, 15) = vArrayChaos(iArrayRowCounter, 15) * dDesiredTotalOutput / dTotalOutput
            vArrayChaos(iArrayRowCounter, 16) = vArrayChaos(iArrayRowCounter, 16) * dDesiredTotalOutput / dTotalOutput
            vArrayChaos(iArrayRowCounter, 17) = vArrayChaos(iArrayRowCounter, 17) * dDesiredTotalOutput / dTotalOutput
            vArrayChaos(iArrayRowCounter, 18) = vArrayChaos(iArrayRowCounter, 18) * dDesiredTotalOutput / dTotalOutput
            
            
            
            'Adjust channel percentages based on minimum percent per channel
            vArrayChaos(iArrayRowCounter, 13) = ((vArrayChaos(iArrayRowCounter, 13) / 100) * (100 - MIN_PERCENT1)) + MIN_PERCENT1
            vArrayChaos(iArrayRowCounter, 14) = ((vArrayChaos(iArrayRowCounter, 14) / 100) * (100 - MIN_PERCENT2)) + MIN_PERCENT2
            vArrayChaos(iArrayRowCounter, 15) = ((vArrayChaos(iArrayRowCounter, 15) / 100) * (100 - MIN_PERCENT3)) + MIN_PERCENT3
            vArrayChaos(iArrayRowCounter, 16) = ((vArrayChaos(iArrayRowCounter, 16) / 100) * (100 - MIN_PERCENT4)) + MIN_PERCENT4
            vArrayChaos(iArrayRowCounter, 17) = ((vArrayChaos(iArrayRowCounter, 17) / 100) * (100 - MIN_PERCENT5)) + MIN_PERCENT5
            vArrayChaos(iArrayRowCounter, 18) = ((vArrayChaos(iArrayRowCounter, 18) / 100) * (100 - MIN_PERCENT6)) + MIN_PERCENT6
            
           
            'Add output from this row to total output for all channels, all rows
            If iArrayRowCounter > 1 Then
                If ((1.9726 * vArrayChaos(iArrayRowCounter, 13) - 16.916) * (vArrayChaos(iArrayRowCounter, 5) - vArrayChaos(iArrayRowCounter - 1, 5))) > 0 Then
                    dAdjustedTotalOutputCH1 = dAdjustedTotalOutputCH1 + ((1.9726 * vArrayChaos(iArrayRowCounter, 13) - 16.916) * (vArrayChaos(iArrayRowCounter, 5) - vArrayChaos(iArrayRowCounter - 1, 5)))
                End If
                
                If ((2.4582 * vArrayChaos(iArrayRowCounter, 14) - 15.204) * (vArrayChaos(iArrayRowCounter, 5) - vArrayChaos(iArrayRowCounter - 1, 5))) > 0 Then
                    dAdjustedTotalOutputCH2 = dAdjustedTotalOutputCH2 + ((2.4582 * vArrayChaos(iArrayRowCounter, 14) - 15.204) * (vArrayChaos(iArrayRowCounter, 5) - vArrayChaos(iArrayRowCounter - 1, 5)))
                End If
                
                If ((3.1604 * vArrayChaos(iArrayRowCounter, 15) - 8.6614) * (vArrayChaos(iArrayRowCounter, 5) - vArrayChaos(iArrayRowCounter - 1, 5))) > 0 Then
                    dAdjustedTotalOutputCH3 = dAdjustedTotalOutputCH3 + ((3.1604 * vArrayChaos(iArrayRowCounter, 15) - 8.6614) * (vArrayChaos(iArrayRowCounter, 5) - vArrayChaos(iArrayRowCounter - 1, 5)))
                End If
                
                If (((-0.0076 * (vArrayChaos(iArrayRowCounter, 16) ^ 2)) + (2.2092 * vArrayChaos(iArrayRowCounter, 16)) - 16.318) * (vArrayChaos(iArrayRowCounter, 5) - vArrayChaos(iArrayRowCounter - 1, 5))) > 0 Then
                    dAdjustedTotalOutputCH4 = dAdjustedTotalOutputCH4 + (((-0.0076 * (vArrayChaos(iArrayRowCounter, 16) ^ 2)) + (2.2092 * vArrayChaos(iArrayRowCounter, 16)) - 16.318) * (vArrayChaos(iArrayRowCounter, 5) - vArrayChaos(iArrayRowCounter - 1, 5)))
                End If
                
                If ((3.807 * vArrayChaos(iArrayRowCounter, 17) - 32.702) * (vArrayChaos(iArrayRowCounter, 5) - vArrayChaos(iArrayRowCounter - 1, 5))) > 0 Then
                    dAdjustedTotalOutputCH5 = dAdjustedTotalOutputCH5 + ((3.807 * vArrayChaos(iArrayRowCounter, 17) - 32.702) * (vArrayChaos(iArrayRowCounter, 5) - vArrayChaos(iArrayRowCounter - 1, 5)))
                End If
                
                If ((2.4158 * vArrayChaos(iArrayRowCounter, 18) - 17.079) * (vArrayChaos(iArrayRowCounter, 5) - vArrayChaos(iArrayRowCounter - 1, 5))) > 0 Then
                    dAdjustedTotalOutputCH6 = dAdjustedTotalOutputCH6 + ((2.4158 * vArrayChaos(iArrayRowCounter, 18) - 17.079) * (vArrayChaos(iArrayRowCounter, 5) - vArrayChaos(iArrayRowCounter - 1, 5)))
                End If
            Else
                If ((1.9726 * vArrayChaos(iArrayRowCounter, 13) - 16.916) * vArrayChaos(iArrayRowCounter, 5)) > 0 Then
                    dAdjustedTotalOutputCH1 = dAdjustedTotalOutputCH1 + ((1.9726 * vArrayChaos(iArrayRowCounter, 13) - 16.916) * vArrayChaos(iArrayRowCounter, 5))
                End If
                
                If ((2.4582 * vArrayChaos(iArrayRowCounter, 14) - 15.204) * vArrayChaos(iArrayRowCounter, 5)) > 0 Then
                    dAdjustedTotalOutputCH2 = dAdjustedTotalOutputCH2 + ((2.4582 * vArrayChaos(iArrayRowCounter, 14) - 15.204) * vArrayChaos(iArrayRowCounter, 5))
                End If
                
                If ((3.1604 * vArrayChaos(iArrayRowCounter, 15) - 8.6614) * vArrayChaos(iArrayRowCounter, 5)) > 0 Then
                    dAdjustedTotalOutputCH3 = dAdjustedTotalOutputCH3 + ((3.1604 * vArrayChaos(iArrayRowCounter, 15) - 8.6614) * vArrayChaos(iArrayRowCounter, 5))
                End If
                
                If (((-0.0076 * (vArrayChaos(iArrayRowCounter, 16) ^ 2)) + (2.2092 * vArrayChaos(iArrayRowCounter, 16)) - 16.318) * vArrayChaos(iArrayRowCounter, 5)) > 0 Then
                    dAdjustedTotalOutputCH4 = dAdjustedTotalOutputCH4 + (((-0.0076 * (vArrayChaos(iArrayRowCounter, 16) ^ 2)) + (2.2092 * vArrayChaos(iArrayRowCounter, 16)) - 16.318) * vArrayChaos(iArrayRowCounter, 5))
                End If
                
                If ((3.807 * vArrayChaos(iArrayRowCounter, 17) - 32.702) * vArrayChaos(iArrayRowCounter, 5)) > 0 Then
                    dAdjustedTotalOutputCH5 = dAdjustedTotalOutputCH5 + ((3.807 * vArrayChaos(iArrayRowCounter, 17) - 32.702) * vArrayChaos(iArrayRowCounter, 5))
                End If
                
                If ((2.4158 * vArrayChaos(iArrayRowCounter, 18) - 17.079) * vArrayChaos(iArrayRowCounter, 5)) > 0 Then
                    dAdjustedTotalOutputCH6 = dAdjustedTotalOutputCH6 + ((2.4158 * vArrayChaos(iArrayRowCounter, 18) - 17.079) * vArrayChaos(iArrayRowCounter, 5))
                End If
            End If
            
            
        
            'Output array to worksheet
            XCelSheet2.Cells(lRowCounter2, 1).Value = CStr(vArrayChaos(iArrayRowCounter, 9)) 'Date
            XCelSheet2.Cells(lRowCounter2, 2).Value = CStr(vArrayChaos(iArrayRowCounter, 10)) 'Hours
            XCelSheet2.Cells(lRowCounter2, 3).Value = CStr(vArrayChaos(iArrayRowCounter, 11)) 'Minutes
            XCelSheet2.Cells(lRowCounter2, 4).Value = CStr(vArrayChaos(iArrayRowCounter, 12)) 'Seconds
            XCelSheet2.Cells(lRowCounter2, 5).Value = CStr(Application.WorksheetFunction.RoundUp(vArrayChaos(iArrayRowCounter, 13), 2)) 'CH1 %
            XCelSheet2.Cells(lRowCounter2, 6).Value = CStr(Application.WorksheetFunction.RoundUp(vArrayChaos(iArrayRowCounter, 14), 2)) 'CH2 %
            XCelSheet2.Cells(lRowCounter2, 7).Value = CStr(Application.WorksheetFunction.RoundUp(vArrayChaos(iArrayRowCounter, 15), 2)) 'CH3 %
            XCelSheet2.Cells(lRowCounter2, 8).Value = CStr(Application.WorksheetFunction.RoundUp(vArrayChaos(iArrayRowCounter, 16), 2)) 'CH4 %
            XCelSheet2.Cells(lRowCounter2, 9).Value = CStr(Application.WorksheetFunction.RoundUp(vArrayChaos(iArrayRowCounter, 17), 2)) 'CH5 %
            XCelSheet2.Cells(lRowCounter2, 10).Value = CStr(Application.WorksheetFunction.RoundUp(vArrayChaos(iArrayRowCounter, 18), 2)) 'CH6 %
            
            lRowCounter2 = lRowCounter2 + 1
            
        Next iArrayRowCounter
        
        dAdjustedTotalOutput = dAdjustedTotalOutputCH1 + dAdjustedTotalOutputCH2 + dAdjustedTotalOutputCH3 + dAdjustedTotalOutputCH4 + dAdjustedTotalOutputCH5 + dAdjustedTotalOutputCH6
        
        'Add row for dark period
        sDateTime = CStr(DateAdd("h", 24, CDate(sDateTime)))
        
        
        sDate = Format(sDateTime, DATE_FORMATTING_STRING)
        sDateTime = Format(sDate & " 0:00:00", DATE_FORMATTING_STRING & " " & TIME_FORMATTING_STRING)
        sTime = Format(sDateTime, TIME_FORMATTING_STRING)
            
        sHH = Left(sTime, InStr(sTime, ":") - 1)
        sMM = Mid(sTime, InStr(sTime, ":") + 1, 2)
        sSS = Right(sTime, 2)
        
        XCelSheet2.Cells(lRowCounter2, 1).Value = sDate
        XCelSheet2.Cells(lRowCounter2, 2).Value = sHH
        XCelSheet2.Cells(lRowCounter2, 3).Value = sMM
        XCelSheet2.Cells(lRowCounter2, 4).Value = sSS
        XCelSheet2.Cells(lRowCounter2, 5).Value = CStr(0)
        XCelSheet2.Cells(lRowCounter2, 6).Value = CStr(0)
        XCelSheet2.Cells(lRowCounter2, 7).Value = CStr(0)
        XCelSheet2.Cells(lRowCounter2, 8).Value = CStr(0)
        XCelSheet2.Cells(lRowCounter2, 9).Value = CStr(0)
        XCelSheet2.Cells(lRowCounter2, 10).Value = CStr(0)
        
        lRowCounter2 = lRowCounter2 + 1
        
        'Advance time for dark period
        sDateTime = CStr(DateAdd("s", dDarkPeriod, CDate(sDateTime)))
        
        'Print total output per channel and overall to log file
        sLine = "Unadjusted output: CH1 Output = " & dTotalOutputCH1 & ", CH2 Output = " & dTotalOutputCH2 & ", CH3 Output = " & dTotalOutputCH3 & ", CH4 Output = " & dTotalOutputCH4 & ",  CH5 Output = " & dTotalOutputCH5 & ", CH6 Output = " & dTotalOutputCH6 & ", " & "Total output = " & dTotalOutput
        Print #iLogFileNum, sLine
        sLine = "Adjusted output: CH1 Output = " & dAdjustedTotalOutputCH1 & ", CH2 Output = " & dAdjustedTotalOutputCH2 & ", CH3 Output = " & dAdjustedTotalOutputCH3 & ", CH4 Output = " & dAdjustedTotalOutputCH4 & ",  CH5 Output = " & dAdjustedTotalOutputCH5 & ", CH6 Output = " & dAdjustedTotalOutputCH6 & ", " & "Total output = " & dAdjustedTotalOutput
        Print #iLogFileNum, sLine
        sLine = ""
        
    Next lRepeat
    
    
    Close iLogFileNum
    Set XCelSheet1 = Nothing
    Set XCelSheet2 = Nothing
    
    Set XCelWorkbook = Nothing
Exit Sub
    
ERROR:
    MsgBox Err.Description, vbCritical, "Error"
    On Error Resume Next
    
    'Protect Output worksheet
    XCelSheet2.Protect (PROTECT_PASSWORD)
    Close iLogFileNum
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
        sLine = sLine & "0" & TEXT_FILE_DELIMITER
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
    
    XCelSheet1.Cells(REPEAT_PATTERN_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER).Value = ""
    XCelSheet1.Cells(REPEAT_PATTERN_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER + 1).Value = ""
    XCelSheet1.Cells(TIME_FROM_LAST_ROW_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER).Value = ""
    XCelSheet1.Cells(TIME_FROM_LAST_ROW_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER + 1).Value = ""
    XCelSheet1.Cells(TIME_BETWEEN_REPEATS_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER).Value = ""
    XCelSheet1.Cells(TIME_BETWEEN_REPEATS_ROW_NUMBER, REPEAT_INTERVAL_CELL_NUMBER + 1).Value = ""
        
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
    Dim sFileName As String: sFileName = QUOTATION & Application.ActiveWorkbook.Path & "\move_file.txt" & QUOTATION
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
        MsgBox "You must be connected to domain 'PlantLab' to use this feature.", vbCritical, "Connection Error"
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

