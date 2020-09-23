Attribute VB_Name = "DateFunctions"
Const Saturday = 6
Const Sunday = 0

Type DDate
    Month As Integer
    Day As Integer
    Year As Integer
End Type

Private Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Const EM_GETLINECOUNT = 186




Function day_of_week(dtDate As DDate)

Dim Y As Long
Dim X As Long
Dim Z As Long

X = 12 * dtDate.Year + dtDate.Month - 3
Y = X / 12
Z = (Y / 4) - (Y / 100) + (Y / 400)

day_of_week = (((734 * X + 15) / 24) - 2 * Y + Z + dtDate.Day + 2) Mod 7

End Function

Function SeparateYear(strDate As String) As DDate
Dim Day As Integer
Dim Month  As Integer
Dim Year As Integer

SeparateYear.Month = CInt(Mid(strDate, 1, 2))
SeparateYear.Day = CInt(Mid(strDate, 4, 2))
SeparateYear.Year = CInt(Mid(strDate, 7))



End Function

Function DaysBetween(intDate As DDate, cmpDate As DDate) As Double
    Dim DateChange As DDate 'keep track of changes
    Dim Days As Integer 'increment for daysbetween
    
    
    Dim StartDay As Integer
    Dim EndDay As Integer
    
    Dim StartMonth As Integer
    Dim EndMonth As Integer
    
    Dim NewMonth As Boolean 'Flag if it is okay to change month
    
    Dim NewYear As Boolean 'flag if it is okay to change year
    
    'Set all "Change" variables to the same as intDate
    DateChange.Day = intDate.Day
    DateChange.Month = intDate.Month
    DateChange.Year = intDate.Year
        
    
    'Check to see if the start date is on a weekend => error
    If day_of_week(intDate) = Saturday Or day_of_week(intDate) = Sunday Then
        MsgBox "Work Started on a weekend!", vbOKOnly, "Action Canceled"
        DaysBetween = -1
        Exit Function
    End If
    
    'To Calculate Days between there are 4 cases
    'case 1 both month and year are the same
    'Case 2 month different year the same
    'Case 3 month same year different
    'Case 4 month different year different
    'Therefore code is for case 4 the rest will fall into place
    
    
    'Continually increment DayChange until it equals cmpDate
    'Set up the loop for the year
    Do While DateChange.Year <= cmpDate.Year
        
        'set up start and end month
        If NewYear = True Then
            StartMonth = 1
            NewYear = False
        Else
            StartMonth = intDate.Month
        End If
        
        If DateChange.Year = cmpDate.Year Then
            EndMonth = cmpDate.Month
        Else
            EndMonth = 12
        End If
        
        For DateChange.Month = StartMonth To EndMonth
        
        
            'Cover StartDay case when both in the same month
            If NewMonth Then
                StartDay = 1
                NewMonth = False
            Else
                StartDay = intDate.Day
            
            End If
         
            'Cover EndDay case when in the same month
            If DateChange.Month = cmpDate.Month Then
            
                EndDay = cmpDate.Day
            Else
                EndDay = LastDayofMonth(DateChange)
            End If

            
            For DateChange.Day = StartDay To EndDay
                'Most important -- If not weekend then add a day
                If day_of_week(DateChange) <> Saturday And _
                    day_of_week(DateChange) <> Sunday Then
                
                    'Not on weekend then increment day
                        Days = Days + 1
                End If
            
                'increment DayChange
            Next DateChange.Day
            
            'End of Month
            NewMonth = True

        Next DateChange.Month
        
            NewYear = True
            DateChange.Year = DateChange.Year + 1
    Loop
    
    DaysBetween = Days - 1
    'The - 1 is because the function says there are 2 days
    'between today and tomorrow -- Think Blockbuster Video rentals! :)
    
End Function

Function LastDayofMonth(A As DDate) As Integer
    If A.Month Mod 12 > 0 And A.Month Mod 12 < 9 Then 'january to july
        If A.Month Mod 2 = 1 Then ' has 31 days
            LastDayofMonth = 31
        Else
            If A.Month = 2 Then 'check for february
            
                If A.Year Mod 4 = 0 Then 'leap year
                     LastDayofMonth = 29
                Else 'not leap year
                    LastDayofMonth = 28
                End If
            Else
                LastDayofMonth = 30
            End If
        End If
    Else 'August to December
        If A.Month Mod 2 = 0 Then 'has 31 days
            LastDayofMonth = 31
        Else
            LastDayofMonth = 30
        End If
    End If

End Function
