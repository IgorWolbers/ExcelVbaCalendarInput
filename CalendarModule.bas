Attribute VB_Name = "CalendarModule"
    Option Explicit
    
    Public Const GWL_STYLE = -16
    Public Const WS_CAPTION = &HC00000
       
    #If VBA7 Then
        #If Win64 Then
            Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias _
            "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
            
            Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias _
            "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, _
            ByVal dwNewLong As LongPtr) As LongPtr
        #Else
            Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias _
            "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
            
            Private Declare Function SetWindowLongPtr Lib "user32" Alias _
            "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, _
            ByVal dwNewLong As LongPtr) As LongPtr
        #End If
        
        Public Declare PtrSafe Function DrawMenuBar Lib "user32" _
        (ByVal hwnd As LongPtr) As LongPtr
        
        Private Declare PtrSafe Function FindWindow Lib "user32" Alias _
        "FindWindowA" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr
        
        Private Declare PtrSafe Function SetTimer Lib "user32" _
        (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, _
        ByVal uElapse As LongPtr, ByVal lpTimerFunc As LongPtr) As LongPtr
    
        Public Declare PtrSafe Function KillTimer Lib "user32" _
        (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As LongPtr
        
        Public TimerID As LongPtr
        
        Dim lngWindow As LongPtr, lFrmHdl As LongPtr
    #Else
    
        Public Declare Function GetWindowLong _
        Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hwnd As Long, ByVal nIndex As Long) As Long
        
        Public Declare Function SetWindowLong _
        Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hwnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
        
        Public Declare Function DrawMenuBar _
        Lib "user32" (ByVal hwnd As Long) As Long
        
        Public Declare Function FindWindowA _
        Lib "user32" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
    
        Public Declare Function SetTimer Lib "user32" ( _
        ByVal hwnd As Long, ByVal nIDEvent As Long, _
        ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
        
        Public Declare Function KillTimer Lib "user32" ( _
        ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
        
        Public TimerID As Long
        Dim lngWindow As Long, lFrmHdl As Long
    #End If
    
    Public TimerSeconds As Single, tim As Boolean
    Public CurMonth As Integer, CurYear As Integer
    Public frmYr As Integer, ToYr As Integer
    
    Public f As frmCalendar
    
    Enum CalendarThemes
        Venom = 0
        MartianRed = 1
        ArcticBlue = 2
        Greyscale = 3
    End Enum
        
    Sub Launch()
        Set f = frmCalendar
        
        With f
            .Caltheme = ArcticBlue
            .LongDateFormat = "dddd dd. mmmm yyyy" '"dddd mmmm dd, yyyy" etc
            .ShortDateFormat = "dd/mm/yyyy"  '"mm/dd/yyyy" or "d/m/y" etc
            .Show
        End With
    End Sub
    
    '~~> Hide the title bar of the userform
    Sub HideTitleBar(frm As Object)
        #If VBA7 Then
            Dim lngWindow As LongPtr, lFrmHdl As LongPtr
            lFrmHdl = FindWindow(vbNullString, frm.Caption)
            lngWindow = GetWindowLongPtr(lFrmHdl, GWL_STYLE)
            lngWindow = lngWindow And (Not WS_CAPTION)
            Call SetWindowLongPtr(lFrmHdl, GWL_STYLE, lngWindow)
            Call DrawMenuBar(lFrmHdl)
        #Else
            Dim lngWindow As Long, lFrmHdl As Long
            lFrmHdl = FindWindow(vbNullString, frm.Caption)
            lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
            lngWindow = lngWindow And (Not WS_CAPTION)
            Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
            Call DrawMenuBar(lFrmHdl)
        #End If
    End Sub
    
    '~~> Start Timer
    Sub StartTimer()
        '~~ Set the timer for 1 second
        TimerSeconds = 1
        TimerID = SetTimer(0&, 0&, TimerSeconds * 1000&, AddressOf TimerProc)
    End Sub
    
    '~~> End Timer
    Sub EndTimer()
        On Error Resume Next
        KillTimer 0&, TimerID
    End Sub
        
    '~~> Update Time
    #If VBA7 And Win64 Then ' 64 bit Excel under 64-bit windows  ' Use LongLong and LongPtr
        Public Sub TimerProc(ByVal hwnd As LongPtr, ByVal uMsg As LongLong, _
        ByVal nIDEvent As LongPtr, ByVal dwTimer As LongLong)
            frmCalendar.Label1.Caption = Split(Format(Time, "h:mm:ss AM/PM"))(0)
            frmCalendar.Label2.Caption = Split(Format(Time, "h:mm:ss AM/PM"))(1)
        End Sub
    #ElseIf VBA7 Then ' 64 bit Excel in all environments
        Public Sub TimerProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, _
        ByVal nIDEvent As LongPtr, ByVal dwTimer As Long)
            frmCalendar.Label1.Caption = Split(Format(Time, "h:mm:ss AM/PM"))(0)
            frmCalendar.Label2.Caption = Split(Format(Time, "h:mm:ss AM/PM"))(1)
        End Sub
    #Else ' 32 bit Excel
        Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, _
        ByVal nIDEvent As Long, ByVal dwTimer As Long)
            frmCalendar.Label1.Caption = Split(Format(Time, "h:mm:ss AM/PM"))(0)
            frmCalendar.Label2.Caption = Split(Format(Time, "h:mm:ss AM/PM"))(1)
        End Sub
    #End If
    
    '~~> Improvement suggested by T.M (https://stackoverflow.com/users/6460297/t-m)
    '(1) Get weekday name
    Function wday(ByVal wd&, ByVal lang As String) As String
        ' Purpose: get weekday in "DDD" format
        wday = Application.Text(DateSerial(6, 1, wd), cPattern(lang) & "ddd")    ' the first day in year 1906 starts with a Sunday
    End Function
    
    '~~> Improvement suggested by T.M (https://stackoverflow.com/users/6460297/t-m)
    '(2) Get month name
    Function mon(ByVal mo&, ByVal lang As String) As String
        ' Example call: mon(12, "1031") or mon(12, "de")
        mon = Application.Text(DateSerial(6, mo, 1), cPattern(lang) & "mmm")
    End Function

    '~~> Improvement suggested by T.M (https://stackoverflow.com/users/6460297/t-m)
    '(3) International patterns
    Function cPattern(ByVal ctry As String) As String
        ' Purpose: return country code pattern for above functions mon() and wday()
        ' Codes: see https://msdn.microsoft.com/en-us/library/dd318693(VS.85).aspx
        ctry = LCase(Trim(ctry))
        Select Case ctry
            Case "1033", "en-us": cPattern = "[$-409]" ' English (US)
            Case "1031", "de": cPattern = "[$-C07]" ' German
            Case "1034", "es": cPattern = "[$-C0A]" ' Spanish
            Case "1036", "fr": cPattern = "[$-80C]" ' French
            Case "1040", "it": cPattern = "[$-410]" ' Italian
            ' more ...
        End Select
    End Function
