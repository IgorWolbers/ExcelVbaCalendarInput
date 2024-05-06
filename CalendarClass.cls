VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalendarClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    '
    '~~> This section is used for handling Commandbutton Control Array
    '
    
    Public WithEvents CommandButtonEvents As MSForms.CommandButton
Attribute CommandButtonEvents.VB_VarHelpID = -1
    
    '~~> Unload the form when the user presses Escape
    Private Sub CommandButtonEvents_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
        If Not f Is Nothing Then If KeyAscii = 27 Then Unload f
    End Sub
    
    '~~> This section delas with showing/displaying controls
    '~~> and updating different labels
    Private Sub CommandButtonEvents_Click()
        f.Label6.Caption = CommandButtonEvents.Tag
        
        If Left(CommandButtonEvents.Name, 1) = "Y" Then
            If Len(Trim(CommandButtonEvents.Caption)) <> 0 Then
                CurYear = Val(CommandButtonEvents.Caption)
                
                With f
                    .HideAllControls
                    .ShowMonthControls
                
                    .Label4.Caption = CurYear
                    .Label5.Caption = 2
                
                    .CommandButton1.Visible = False
                    .CommandButton2.Visible = False
                End With
            End If
        ElseIf Left(CommandButtonEvents.Name, 1) = "M" Then
            Select Case UCase(CommandButtonEvents.Caption)
                Case "JAN": CurMonth = 1
                Case "FEB": CurMonth = 2
                Case "MAR": CurMonth = 3
                Case "APR": CurMonth = 4
                Case "MAY": CurMonth = 5
                Case "JUN": CurMonth = 6
                Case "JUL": CurMonth = 7
                Case "AUG": CurMonth = 8
                Case "SEP": CurMonth = 9
                Case "OCT": CurMonth = 10
                Case "NOV": CurMonth = 11
                Case "DEC": CurMonth = 12
            End Select
            
            f.HideAllControls
            f.ShowSpecificMonth
        End If
    End Sub
