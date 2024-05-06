VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendar 
   Caption         =   "UserForm1"
   ClientHeight    =   7788
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   5532
   OleObjectBlob   =   "frmCalendar.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    Private TimerID As Long, TimerSeconds As Single, tim As Boolean
    Dim curDate As Date
    Dim i As Long
    Dim thisDay As Integer, thisMonth As Integer, thisYear As Integer
    Dim CBArray() As New CalendarClass
        
    Dim NewXpos As Single
    Dim NewYpos As Single

    Private Cal_theme As CalendarThemes
    Private LdtFormat As String, SdtFormat As String
    
    Public Property Let LongDateFormat(s As String)
        LdtFormat = s
        
        Label3.Caption = Format(Date, LdtFormat)
    End Property
    
    Public Property Get LongDateFormat() As String
        LongDateFormat = LdtFormat
    End Property
    
    Public Property Let ShortDateFormat(s As String)
        SdtFormat = s
        
        Label6.Caption = Format(Date, SdtFormat)
    End Property
    
    Public Property Get ShortDateFormat() As String
        ShortDateFormat = SdtFormat
    End Property
    
    Public Property Let Caltheme(Theme As CalendarThemes)
        Cal_theme = Theme
        
        '~~> Set the color of controls
        Select Case Cal_theme
            Case CalendarThemes.Venom
                Me.BackColor = RGB(69, 69, 69)
                Frame1.BackColor = RGB(69, 69, 69)
                Label2.ForeColor = RGB(182, 182, 182)
                Label3.ForeColor = RGB(66, 156, 227)
                Label6.ForeColor = RGB(66, 156, 227)
                Label4.ForeColor = RGB(223, 223, 223)
                CommandButton1.ForeColor = RGB(201, 201, 201)
                CommandButton2.ForeColor = RGB(201, 201, 201)
            Case CalendarThemes.MartianRed
                Me.BackColor = RGB(136, 0, 27)
                Frame1.BackColor = RGB(136, 0, 27)
                
                Label2.ForeColor = RGB(255, 255, 255)
                Label3.ForeColor = RGB(255, 128, 127)
                Label6.ForeColor = RGB(255, 255, 255)
                Label4.ForeColor = RGB(255, 255, 255)
                
                CommandButton1.ForeColor = RGB(0, 0, 0)
                CommandButton2.ForeColor = RGB(0, 0, 0)
                DTINSERT.ForeColor = RGB(0, 0, 0)
                
                For i = 1 To 7
                    With Me.Controls("WD" & i)
                        .BackStyle = fmBackStyleOpaque
                        .ForeColor = RGB(0, 0, 0)
                        .BackColor = RGB(198, 86, 85)
                    End With
                Next i
                
                For i = 1 To 42
                    With Me.Controls("CB" & i)
                        .ForeColor = RGB(255, 255, 255)
                    End With
                Next i
            Case CalendarThemes.ArcticBlue
                Me.BackColor = RGB(9, 148, 223)
                Frame1.BackColor = RGB(9, 148, 223)
                
                Label2.ForeColor = RGB(255, 255, 255)
                Label3.ForeColor = RGB(34, 66, 125)
                Label6.ForeColor = RGB(255, 255, 255)
                Label4.ForeColor = RGB(255, 255, 255)
                
                CommandButton1.ForeColor = RGB(0, 0, 0)
                CommandButton2.ForeColor = RGB(0, 0, 0)
                DTINSERT.ForeColor = RGB(0, 0, 0)
                
                For i = 1 To 7
                    With Me.Controls("WD" & i)
                        .BackStyle = fmBackStyleOpaque
                        .ForeColor = RGB(0, 0, 0)
                        .BackColor = RGB(128, 128, 192)
                    End With
                Next i
                
                For i = 1 To 42
                    With Me.Controls("CB" & i)
                        .ForeColor = RGB(255, 255, 255)
                    End With
                Next i
            Case CalendarThemes.Greyscale
                Me.BackColor = RGB(240, 240, 240)
                Frame1.BackColor = RGB(240, 240, 240)
                
                Label1.ForeColor = RGB(0, 0, 0)
                Label2.ForeColor = RGB(0, 0, 0)
                Label3.ForeColor = RGB(0, 0, 0)
                Label6.ForeColor = RGB(0, 0, 0)
                Label4.ForeColor = RGB(0, 0, 0)
                
                CommandButton1.ForeColor = RGB(0, 0, 0)
                CommandButton2.ForeColor = RGB(0, 0, 0)
                DTINSERT.ForeColor = RGB(0, 0, 0)
                
                For i = 1 To 7
                    With Me.Controls("WD" & i)
                        .BackStyle = fmBackStyleOpaque
                        .ForeColor = RGB(0, 0, 0)
                        .BackColor = RGB(240, 240, 240)
                    End With
                Next i
                
                For i = 1 To 42
                    If i < 13 Then
                        Me.Controls("M" & i).ForeColor = RGB(0, 0, 0)
                        Me.Controls("Y" & i).ForeColor = RGB(0, 0, 0)
                    End If
                    
                    Me.Controls("CB" & i).ForeColor = RGB(0, 0, 0)
                Next i
        End Select
        
        '~~> Populate this months calendar
        PopulateCalendar Date
    End Property
    
    Public Property Get Caltheme() As CalendarThemes
        Caltheme = Cal_theme
    End Property
    
    Private Sub CbLanguage_Click()
        If CbLanguage.ListIndex = -1 Then Exit Sub
        
        '~~> LCID Codes: https://www.science.co.il/language/Locale-codes.php
        Select Case CbLanguage.ListIndex
            Case 0: ChangeLanguage (1033)
            Case 1: ChangeLanguage (1031)
            Case 2: ChangeLanguage (1034)
            Case 3: ChangeLanguage (1040)
            Case 4: ChangeLanguage (1036)
        End Select
    End Sub
    
Private Sub Label3_Click()

End Sub

    Private Sub UserForm_Initialize()
        With CbLanguage
            .AddItem "EN"
            .AddItem "GER"
            .AddItem "SPA"
            .AddItem "ITL"
            .AddItem "FRE"
            
            .ListIndex = 0
        End With
        
        '~~> Hide the Title Bar
        HideTitleBar Me
        
        Me.Caltheme = Venom
        Me.LongDateFormat = "dddd mmmm dd, yyyy"
        Me.ShortDateFormat = "dd/mm/yyyy"
        
        '~~> Create a command button control array so that
        '~~> when we press escape, we can unload the userform
        Dim CBCtl As Control
        
        i = 0
        
        For Each CBCtl In Me.Controls
            If TypeOf CBCtl Is MSForms.CommandButton Then
                i = i + 1
                ReDim Preserve CBArray(1 To i)
                Set CBArray(i).CommandButtonEvents = CBCtl
            End If
        Next CBCtl
        Set CBCtl = Nothing
        
        '~~> Set the Time
        StartTimer
                  
        curDate = Date
        
        thisDay = Day(Date): thisMonth = Month(Date): thisYear = Year(Date)
         
        CurYear = Year(Date): CurMonth = Month(Date)
        
        '~~> Populate this months calendar
        PopulateCalendar curDate
    End Sub
    
    '~~> Improvement suggested by T.M (https://stackoverflow.com/users/6460297/t-m)
    Sub ChangeLanguage(ByVal LCID As Long)
        Dim i&
        '~~> Week Day Name
         For i = 1 To 7
             Me.Controls("WD" & i).Caption = Left(wday(i, LCID), 2)
         Next i
        '~~> Month Name
         For i = 1 To 12
             Me.Controls("M" & i).Caption = Left(mon(i, LCID), 3)
         Next i
    End Sub
    
    '~~> The below 4 procedures will assist in moving the borderless userform
    Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        If Button = 1 Then
            NewXpos = X
            NewYpos = Y
        End If
    End Sub
    Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        If Button And 1 Then
            Me.Left = Me.Left + (X - NewXpos)
            Me.Top = Me.Top + (Y - NewYpos)
        End If
    End Sub
    Private Sub Frame1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        If Button = 1 Then
            NewXpos = X
            NewYpos = Y
        End If
    End Sub
    Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        If Button And 1 Then
            Me.Left = Me.Left + (X - NewXpos)
            Me.Top = Me.Top + (Y - NewYpos)
        End If
    End Sub

    '~~> Insert Selected date
    Private Sub DTINSERT_Click()
        If Len(Trim(Label6.Caption)) = 0 Then
            MsgBox "Please select a date first", vbCritical, "No date selected"
            Exit Sub
        End If
        '~~> Change the code here to insert date where ever you want
        MsgBox Label6.Caption, vbInformation, "Date selected"
    End Sub
    
    '~~> Stop timer in the terminate event
    Private Sub UserForm_Terminate()
        EndTimer
    End Sub
    
    '~~> Unload the form when user presses escape
    Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
        If KeyAscii = 27 Then Unload Me
    End Sub
    Private Sub CbLanguage_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
        If KeyAscii = 27 Then Unload Me
    End Sub
    
    '~~> UP Button
    Private Sub CommandButton1_Click()
        Select Case Label5.Caption
            Case 1 '~~> When user presses the up button when the dates are displayed
                curDate = DateSerial(CurYear, CurMonth, 0)
                
                '~~> Check if date is >= 1/1/1919
                If curDate >= DateSerial(1919, 1, 1) Then
                    '~~> Populate prev months calendar
                    PopulateCalendar curDate
                End If
            Case 2 '<~~ Do nothing
            Case 3 '~~> When user presses the up button when the Year Range is displayed
                If frmYr > 1919 Then
                    ResetBlueColor
                    
                    Dim NewToYr As Integer
                    
                    ToYr = frmYr - 1
                    NewToYr = frmYr - 1
                    
                    For i = 1 To 12
                        Me.Controls("Y" & i).Caption = ""
                    Next i
                
                    For i = 12 To 1 Step -1
                        If Not NewToYr < 1919 Then
                            With Me.Controls("Y" & i)
                                .Caption = NewToYr
                                
                                If NewToYr = thisYear Then
                                    .BackStyle = fmBackStyleOpaque
                                    .BackColor = &H8000000D
                                End If
                                
                                .Visible = True
                                
                                NewToYr = NewToYr - 1
                            End With
                        End If
                    Next i
                    
                    frmYr = NewToYr + 1
                    Label4.Caption = (NewToYr + 1) & " - " & ToYr
                End If
        End Select
    End Sub
    
    '~~> Down Button
    Private Sub CommandButton2_Click()
        Select Case Label5.Caption
            Case 1 '~~> When user presses the down button when the dates are displayed
                curDate = DateAdd("m", 1, DateSerial(CurYear, CurMonth, 1))
                
                '~~> Check if date is <= 31/12/2119
                If curDate <= DateSerial(2119, 12, 31) Then
                    '~~> Populate prev months calendar
                    PopulateCalendar curDate
                End If
            Case 2 '<~~ Do nothing
            Case 3 '~~> When user presses the down button when the Year Range is displayed
                frmYr = Val(Split(Label4.Caption, "-")(0))
                ToYr = Val(Split(Label4.Caption, "-")(1))
                 
                If ToYr < 2119 Then
                    ResetBlueColor
                    
                    Dim NewFrmYr As Integer
                    
                    frmYr = ToYr + 1
                    NewFrmYr = ToYr + 1
                    
                    For i = 1 To 12
                        Me.Controls("Y" & i).Caption = ""
                    Next i
                
                    For i = 1 To 12
                        If NewFrmYr < 2119 Then
                            With Me.Controls("Y" & i)
                                .Caption = NewFrmYr
                                
                                If NewFrmYr = thisYear Then
                                    .BackStyle = fmBackStyleOpaque
                                    .BackColor = &H8000000D
                                End If
                                
                                .Visible = True
                                
                                NewFrmYr = NewFrmYr + 1
                            End With
                        ElseIf NewFrmYr = 2119 Then
                            With Me.Controls("Y" & i)
                                .Caption = NewFrmYr
                                .Visible = True
                                NewFrmYr = NewFrmYr + 1
                            End With
                        End If
                    Next i
                    
                    If NewFrmYr = 2119 Then ToYr = NewFrmYr Else ToYr = NewFrmYr - 1
                    Label4.Caption = frmYr & " - " & ToYr
                End If
        End Select
    End Sub
    
    '~~> Populate the calendar for a specific month
    Sub PopulateCalendar(d As Date)
        '~~> Get the day of 1st of the month
        Dim m As Integer, Y As Integer
        Dim i As Integer, j As Integer
        Dim LastDay As Integer, NextCounter As Integer, PrevCounter As Integer
        Dim dtOne As Date, dtLast As Date, dtNext As Date
        
        ResetBlueColor
        
        Select Case Cal_theme
            Case CalendarThemes.Venom
                For i = 1 To 42
                    Me.Controls("CB" & i).ForeColor = RGB(255, 255, 255)
                Next i
            Case CalendarThemes.MartianRed
                For i = 1 To 42
                    Me.Controls("CB" & i).ForeColor = RGB(255, 255, 255)
                Next i
        End Select
        
        CurYear = Year(d)
        CurMonth = Month(d)
        
        m = Month(d): Y = Year(d)
        
        dtOne = DateSerial(Y, m, 1)
        dtLast = DateSerial(Year(dtOne), Month(dtOne), 0)
        dtNext = DateAdd("m", 1, DateSerial(Year(dtOne), Month(dtOne), 1))
        
        Select Case Weekday(dtOne, 0)
            Case 1:
                With CB1
                    .Caption = 1
                    .Tag = Format(DateSerial(Year(d), Month(d), 1), frmCalendar.ShortDateFormat)
                    Select Case Cal_theme
                        Case CalendarThemes.ArcticBlue
                            .ForeColor = RGB(255, 255, 255)
                        Case Else
                            .ForeColor = RGB(0, 0, 0)
                    End Select
                End With
                NextCounter = 2: PrevCounter = 0
            Case 2
                With CB2
                    .Caption = 1
                    .Tag = Format(DateSerial(Year(d), Month(d), 1), frmCalendar.ShortDateFormat)
                    Select Case Cal_theme
                        Case CalendarThemes.ArcticBlue
                            .ForeColor = RGB(255, 255, 255)
                        Case Else
                            .ForeColor = RGB(0, 0, 0)
                    End Select
                End With
                NextCounter = 3: PrevCounter = 1
            Case 3
                With CB3
                    .Caption = 1
                    .Tag = Format(DateSerial(Year(d), Month(d), 1), frmCalendar.ShortDateFormat)
                    Select Case Cal_theme
                        Case CalendarThemes.ArcticBlue
                            .ForeColor = RGB(255, 255, 255)
                        Case Else
                            .ForeColor = RGB(0, 0, 0)
                    End Select
                End With
                NextCounter = 4: PrevCounter = 2
            Case 4
                With CB4
                    .Caption = 1
                    .Tag = Format(DateSerial(Year(d), Month(d), 1), frmCalendar.ShortDateFormat)
                    Select Case Cal_theme
                        Case CalendarThemes.ArcticBlue
                            .ForeColor = RGB(255, 255, 255)
                        Case Else
                            .ForeColor = RGB(0, 0, 0)
                    End Select
                End With
                NextCounter = 5: PrevCounter = 3
            Case 5
                With CB5
                    .Caption = 1
                    .Tag = Format(DateSerial(Year(d), Month(d), 1), frmCalendar.ShortDateFormat)
                    Select Case Cal_theme
                        Case CalendarThemes.ArcticBlue
                            .ForeColor = RGB(255, 255, 255)
                        Case Else
                            .ForeColor = RGB(0, 0, 0)
                    End Select
                End With
                NextCounter = 6: PrevCounter = 4
            Case 6
                With CB6
                    .Caption = 1
                    .Tag = Format(DateSerial(Year(d), Month(d), 1), frmCalendar.ShortDateFormat)
                    Select Case Cal_theme
                        Case CalendarThemes.ArcticBlue
                            .ForeColor = RGB(255, 255, 255)
                        Case Else
                            .ForeColor = RGB(0, 0, 0)
                    End Select
                End With
                NextCounter = 7: PrevCounter = 5
            Case 7
                With CB7
                    .Caption = 1
                    .Tag = Format(DateSerial(Year(d), Month(d), 1), frmCalendar.ShortDateFormat)
                    Select Case Cal_theme
                        Case CalendarThemes.ArcticBlue
                            .ForeColor = RGB(255, 255, 255)
                        Case Else
                            .ForeColor = RGB(0, 0, 0)
                    End Select
                End With
                NextCounter = 8: PrevCounter = 6
        End Select
        
        LastDay = Val(Format(Excel.Application.WorksheetFunction.EoMonth(dtOne, 0), "dd"))
        
        For i = 2 To LastDay
            Me.Controls("CB" & NextCounter).Caption = i
            Me.Controls("CB" & NextCounter).Tag = Format(DateSerial(Year(d), Month(d), i), frmCalendar.ShortDateFormat)
            
            Select Case Cal_theme
                Case CalendarThemes.ArcticBlue
                    Me.Controls("CB" & NextCounter).ForeColor = RGB(255, 255, 255)
                Case Else
                    Me.Controls("CB" & NextCounter).ForeColor = RGB(0, 0, 0)
            End Select
            If i = thisDay And Month(d) = thisMonth And Year(d) = thisYear Then
                With Me.Controls("CB" & NextCounter)
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = &H8000000D
                End With
            End If
    
            NextCounter = NextCounter + 1
        Next i
        
        j = 1
        
        If NextCounter < 43 Then
            For i = NextCounter To 42
                With Me.Controls("CB" & i)
                    .Caption = j
                    .Tag = Format(DateSerial(Year(dtNext), Month(dtNext), j), frmCalendar.ShortDateFormat)
                    
                    Select Case Cal_theme
                        Case CalendarThemes.ArcticBlue
                            .ForeColor = RGB(0, 0, 0)
                        Case Else
                            .ForeColor = RGB(132, 132, 132)
                    End Select
                End With
                j = j + 1
            Next i
        End If
        
        LastDay = Val(Format(dtLast, "dd"))
        
        If PrevCounter > 1 Then
            For i = PrevCounter To 1 Step -1
                With Me.Controls("CB" & i)
                    .Caption = LastDay
                    .Tag = Format(DateSerial(Year(dtLast), Month(dtLast), LastDay), frmCalendar.ShortDateFormat)
                    Select Case Cal_theme
                        Case CalendarThemes.ArcticBlue
                            .ForeColor = RGB(0, 0, 0)
                        Case Else
                            .ForeColor = RGB(132, 132, 132)
                    End Select
                End With
                LastDay = LastDay - 1
            Next i
        ElseIf PrevCounter = 1 Then
            With Me.Controls("CB1")
                .Caption = LastDay
                .Tag = Format(DateSerial(Year(dtLast), Month(dtLast), LastDay), frmCalendar.ShortDateFormat)
                Select Case Cal_theme
                    Case CalendarThemes.ArcticBlue
                        .ForeColor = RGB(0, 0, 0)
                    Case Else
                        .ForeColor = RGB(132, 132, 132)
                End Select
            End With
        End If
        
        Label4.Caption = Format(d, "mmmm yyyy")
        
        CB1.SetFocus '~~> To allow user to press esc to quit
    End Sub
    
    '~~> Hide all controls
    Sub HideAllControls()
         DTINSERT.Visible = False
         Label6.Visible = False
         
        Select Case Cal_theme
            Case CalendarThemes.Venom
                For i = 1 To 7
                   With Me.Controls("WD" & i)
                       .Visible = False
                       .BackStyle = fmBackStyleTransparent
                       .BackColor = &H8000000F
                   End With
                Next i
            Case CalendarThemes.MartianRed
                For i = 1 To 7
                   With Me.Controls("WD" & i)
                       .Visible = False
                       .BackStyle = fmBackStyleOpaque
                       .BackColor = RGB(198, 86, 85)
                   End With
                Next i
            Case CalendarThemes.ArcticBlue
                For i = 1 To 7
                   With Me.Controls("WD" & i)
                       .Visible = False
                       .BackStyle = fmBackStyleOpaque
                       .BackColor = RGB(128, 128, 192)
                   End With
                Next i
             Case CalendarThemes.Greyscale
                For i = 1 To 7
                   With Me.Controls("WD" & i)
                       .Visible = False
                       '.BackStyle = fmBackStyleOpaque
                       '.BackColor = RGB(128, 128, 192)
                   End With
                Next i
        End Select
        
         For i = 1 To 42
            With Me.Controls("CB" & i)
                .Visible = False
                .BackStyle = fmBackStyleTransparent
                .BackColor = &H8000000F
            End With
         Next i
    
         For i = 1 To 12
            With Me.Controls("M" & i)
                .Visible = False
                .BackStyle = fmBackStyleTransparent
                .BackColor = &H8000000F
            End With
         Next i
         
         For i = 1 To 12
            With Me.Controls("Y" & i)
                .Visible = False
                .BackStyle = fmBackStyleTransparent
                .BackColor = &H8000000F
            End With
         Next i
    End Sub
    
    '~~> Show the months when user clicks on the date label
    Sub ShowMonthControls()
         For i = 1 To 12
            Me.Controls("M" & i).Visible = True
            
            If i = thisMonth Then
                With Me.Controls("M" & i)
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = &H8000000D
                End With
            End If
         Next i
    End Sub
    
    '~~> Show the details for specific month
    Sub ShowSpecificMonth()
        DTINSERT.Visible = True
        Label6.Visible = True
         
        For i = 1 To 42
            If i < 8 Then
                Me.Controls("WD" & i).Visible = True
            End If
            Me.Controls("CB" & i).Visible = True
        Next i
         
        Label4.Caption = Format(DateSerial(CurYear, CurMonth, 1), "mmm yyyy")
        Label5.Caption = 1
        
        CommandButton1.Visible = True
        CommandButton2.Visible = True
                
        PopulateCalendar DateSerial(CurYear, CurMonth, 1)
    End Sub
    
    '~~> Removes the blue color from current day/month/year
    Sub ResetBlueColor()
         For i = 1 To 42
            With Me.Controls("CB" & i)
                .BackStyle = fmBackStyleTransparent
                .BackColor = &H8000000F
            End With
         Next i
    
         For i = 1 To 12
            With Me.Controls("Y" & i)
                .BackStyle = fmBackStyleTransparent
                .BackColor = &H8000000F
            End With
         Next i
    End Sub
    
    '~~> Handles the month to year to year slab display
    Private Sub Label4_Click()
         Select Case Label5.Caption
            Case 1
                HideAllControls
           
                Label4.Caption = Split(Label4.Caption)(1)
                Label5.Caption = 2
                
                ShowMonthControls
                
                CommandButton1.Visible = False
                CommandButton2.Visible = False
            Case 2
                HideAllControls
                CommandButton1.Visible = True
                CommandButton2.Visible = True
                            
                ToYr = Val(Label4.Caption)
                frmYr = ToYr - 11
                
                If frmYr < 1919 Then frmYr = 1919
                
                Label4.Caption = frmYr & " - " & ToYr
                Label5.Caption = 3
                
                For i = 1 To 12
                    Me.Controls("Y" & i).Caption = ""
                Next i
                
                For i = 12 To 1 Step -1
                    If Not ToYr < 1919 Then
                        With Me.Controls("Y" & i)
                            .Caption = ToYr
                            .Visible = True
                            
                            If ToYr = thisYear Then
                                With Me.Controls("Y" & i)
                                    .BackStyle = fmBackStyleOpaque
                                    .BackColor = &H8000000D
                                End With
                            End If
                            
                            ToYr = ToYr - 1
                        End With
                    End If
                Next i
                
                Y1.SetFocus
                
                Label5.Caption = 3
            Case 3 'Do Nothing
         End Select
    End Sub

