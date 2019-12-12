Attribute VB_Name = "fonctions_base"
Option Explicit
Option Base 1

Function is_leap_year(year_ As Integer) As Boolean

    is_leap_year = Day(DateSerial((year_), 3, 0)) = 29
    
End Function

Public Function daysinyear(year_ As Integer) As Integer

    daysinyear = DateDiff("d", CDate("1/1/" & year_), CDate("31/12/" & year_)) + 1
    
End Function

Public Function daysinmonth(dat As Date) As Integer

    If month(dat) < 12 Then
        daysinmonth = DateDiff("d", CDate("1/" & month(dat) & "/" & year(dat)), CDate("1/" & month(dat) + 1 & "/" & year(dat)))
    Else
        daysinmonth = DateDiff("d", CDate("1/" & month(dat) & "/" & year(dat)), CDate("1/" & 1 & "/" & year(dat) + 1))
    End If

End Function

Function is_same_year(dat1 As Date, dat2 As Date) As Boolean

    If year(dat1) = year(dat2) Then
        is_same_year = True
    
    Else: is_same_year = False
    End If
    
End Function

Function day_count(dat1 As Date, dat2 As Date) As Integer
    day_count = dat2 - dat1
End Function

Function delta_t(ByVal date1 As Date, ByVal date2 As Date, ByVal conv As String) As Variant
    
    Select Case conv
    
    Case "Act/360"
    delta_t = (date2 - date1) / 360
    
    Case "Act/365"
    delta_t = (date2 - date1) / 365
    
    Case "Act/366"
    delta_t = (date2 - date1) / 366
    
    Case "30/360"
    delta_t = (WorksheetFunction.max(30 - Day(date1), 0) + WorksheetFunction.min(Day(date2), 30) + 360 * (year(date2) - year(date1)) + 30 * (month(date2) - month(date1) - 1)) / 360
    
    Case "Act/Act"
        'If Year(date1) = Year(date2) Then
       ' delta_t = WorksheetFunction.YearFrac(date1, Date2, 1)
        delta_t = (date2 - date1) / 365.25
        'delta_t = (date2 - date1) / dayinyear(date1)
        'Else
        'delta_t = (daytoendofyear(date1) - 0.5) / dayinyear(date1) + (dayfrombeginingofyear(date2) - 0.5) / dayinyear(date2) + Year(date2) - Year(date1) - 1
        'End If
    End Select

End Function




'---------------------------------------------- Manipulation tableaux, vecteurs -----------------------------------------
'------------------------------------------------------------------------------------------------------------------------

Function Inverse(ByVal Tabl As Variant) As Variant
    Dim temp() As Variant
    Dim valeur() As Variant
    valeur = Tabl
    ReDim temp(UBound(Tabl))
    Dim a As Integer, i As Integer, j As Integer
    
    
    a = UBound(Tabl)
    For i = 1 To a
        j = a - (i - 1)
        temp(i) = valeur(j)
    Next i
    
    Inverse = temp

End Function
Function Interpolation(ByVal T1 As Date, _
                       ByVal curveDate As Variant, _
                       ByVal curveRate As Variant) As Double

    Dim CourbeDate() As Variant
    Dim CourbeTaux() As Variant
    Dim i As Integer, a As Integer
    
    CourbeDate = curveDate
    CourbeTaux = curveRate
    
    If T1 <= CourbeDate(1, 1) Then
        Interpolation = CourbeTaux(1, 1)
    Exit Function
    
    ElseIf T1 >= CourbeDate(UBound(CourbeDate), 1) Then
        Interpolation = CourbeTaux(UBound(CourbeTaux), 1)
    Exit Function
    
    Else
    i = 1
    
    For a = 2 To UBound(CourbeDate, 1)
    
        If CourbeDate(i, 1) <= T1 Then 'récupère l'indice juste apres pour la date
            i = i + 1
        End If
        
        If i > UBound(CourbeDate, 1) Then
            i = i - 1
    Exit For
    
    End If
    Next a
    
    Interpolation = CourbeTaux(i - 1, 1) + (T1 - CourbeDate(i - 1, 1)) * ((CourbeTaux(i, 1) - CourbeTaux(i - 1, 1)) / (CourbeDate(i, 1) - CourbeDate(i - 1, 1)))
    Interpolation = Interpolation
    
    End If

End Function
'fonction pour transformer un range en tableau vba
Function Transforme(ByVal a As Range) As Variant

    Dim lignes As Integer
    Dim colonnes As Integer
    Dim i As Integer, j As Integer
    
    lignes = a.Rows.count
    colonnes = a.Columns.count
    
    Dim Trans() As Variant
    ReDim Trans(lignes, colonnes)
    
    For i = 1 To lignes
        For j = 1 To colonnes
            Trans(i, j) = a(i, j)
        Next j
    Next i
    
    Transforme = Trans
    
End Function




'---------------------------------Fonction de support -------------------------------------------
'---------------------------------------------------------------------------------------------------
Function Maxi(ByVal a As Double, ByVal b As Double) As Double

    
    If a >= b Then
        Maxi = a
    Else: Maxi = b
    End If

End Function



Function Mini(ByVal a As Double, ByVal b As Double) As Double

    
    If a >= b Then
        Mini = b
    Else: Mini = a
    End If

End Function
