Attribute VB_Name = "fonctions_base_finance"
Option Explicit
Option Base 1
'--------------------------- FONCTION POUR LES BUSINESS DAYS  -------------------------------------------------
'---------------------------------------------------------------------------------------------------------------


'passe au lundi si la date tombe en weekend
    Function Following(ByVal date1 As Date) As Date
    Following = date1
    
    If Weekday(date1) = 1 Then
        Following = DateAdd("d", 1, date1)
    
    ElseIf Weekday(date1) = 7 Then
        Following = DateAdd("d", 2, date1)


End If

End Function

'passe au vendredi si la date tombe en weekend
    
    Function Preceding(ByVal date1 As Date) As Date
    Preceding = date1
    
    If Weekday(date1) = 1 Then
        Preceding = DateAdd("d", -2, date1)
    ElseIf Weekday(date1) = 7 Then
    
    Preceding = DateAdd("d", -1, date1)


End If

End Function



'passe au lundi suivant sauf si on change de moi dans ce cas la passe au vendredi
Function ModFollowing(ByVal date1 As Date) As Date

    Dim temp As Date
    temp = date1
    ModFollowing = date1
    If Weekday(date1) = 1 Then
    
        If month(DateAdd("d", 1, date1)) = month(temp) Then
            ModFollowing = Following(temp)
        Else: ModFollowing = Preceding(temp)
        End If
        
    ElseIf Weekday(date1) = 7 Then
    
        If month(DateAdd("d", 2, date1)) = month(temp) Then
            ModFollowing = Following(temp)
        Else: ModFollowing = Preceding(temp)
        End If
    
    End If



End Function

'passe au vendredi precedent sauf si on change de moi dans ce cas la passe au lundi suivant
Function ModPreceding(ByVal date1 As Date) As Date

    Dim temp As Date
    temp = date1
    ModPreceding = date1
    If Weekday(date1) = 1 Then
    
        If month(DateAdd("d", -2, date1)) = month(temp) Then
            ModPreceding = Preceding(temp)
        Else: ModPreceding = Following(temp)
        End If
        
    ElseIf Weekday(date1) = 7 Then
    
        If month(DateAdd("d", -1, date1)) = month(temp) Then
            ModPreceding = Preceding(temp)
        Else:
        ModPreceding = Following(temp)
        End If
    
    End If

End Function

'parametre convention est pour la convention de jours a prendre
Function Business_Day(ByVal date1 As Date, ByVal convention As String)

    If convention = "Following" Then
    Business_Day = Following(date1)
    
    ElseIf convention = "Preceding" Then
    Business_Day = Preceding(date1)
    
    ElseIf convention = "Modified Following" Then
    Business_Day = ModFollowing(date1)
    
    ElseIf convention = "Modified Preceding" Then
    Business_Day = ModPreceding(date1)
    End If
    

End Function


'---------------------------------------------Generateur Echeancier --------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------


Function Affiche_Cash_Flow_Dates(ByVal dateValo As Date, _
                ByVal date1 As Date, _
              ByVal date2 As Date, _
              ByVal freq As Integer, _
              ByVal convention As String, _
              ByVal BrokenPeriod As String) As Variant
              
    Dim i As Integer
    Dim k As Integer
    Dim temp2 As Date
    Dim Echeancier() As Variant
    Dim echeaTemp() As Variant
    Dim j As Integer
    Dim echea1() As Variant
    
    k = 1
    i = 1
    j = 1
    ReDim Echeancier(1)
    Dim temp As Date
    temp = date1
    temp2 = date2
    
    echea1 = Cash_Flow_Dates(date1, date2, freq, convention, BrokenPeriod)
    
    i = 1
    If dateValo <= date1 Then
        Affiche_Cash_Flow_Dates = echea1
    Else
        
        While echea1(i) < dateValo
            i = i + 1
        Wend
        k = UBound(echea1) - i + 2
        For j = 1 To k
            ReDim Preserve echeaTemp(j)
            echeaTemp(j) = echea1(i - 1)
            i = i + 1
         
        Next j
       Affiche_Cash_Flow_Dates = echeaTemp
        
    End If
    
     
End Function

Function Cash_Flow_Dates(ByVal date1 As Date, _
              ByVal date2 As Date, _
              ByVal freq As Integer, _
              ByVal convention As String, _
              ByVal BrokenPeriod As String) As Variant

              
Dim i As Integer
i = 1

Dim Echeancier() As Variant
ReDim Echeancier(1)

Dim temp As Date
Dim temp2 As Date
temp = date1
temp2 = date2


If BrokenPeriod = "End" Then
     Echeancier(1) = date1
        While DateAdd("m", i * freq, temp) < date2
            ReDim Preserve Echeancier(i + 1)
            Echeancier(i + 1) = Business_Day(DateAdd("m", i * freq, date1), convention) 'a chaque case i la freq et en (i-1)
            i = i + 1
        Wend 'en sortie de boucle i contient le nombre de case dans lecheancier +1
        ReDim Preserve Echeancier(i + 1)
        Echeancier(i + 1) = date2
        Cash_Flow_Dates = Echeancier

ElseIf BrokenPeriod = "Start" Then
    
        Echeancier(1) = date2
        While DateAdd("m", -i * freq, temp2) > date1
            ReDim Preserve Echeancier(i + 1)
            Echeancier(i + 1) = Business_Day(DateAdd("m", -i * freq, date2), convention)
            i = i + 1
        Wend
        ReDim Preserve Echeancier(i + 1)
        Echeancier(i + 1) = date1
    
        Cash_Flow_Dates = Inverse(Echeancier)
End If

End Function


Function Tx_forward(ByVal date1 As Date, ByVal date2 As Date, ByVal df1 As Double, ByVal df2 As Double, ByVal conv_TxFwd As String, ByVal count As String)
'Attention à prendre les bons DF, correspondant à la bonne date de valo

    Dim forward As Double
    
    If (conv_TxFwd = "Taux linéaires") Then
        forward = (1 / delta_t(date1, date2, count)) * (df1 / df2 - 1)
    ElseIf (conv_TxFwd = "Taux actuariels") Then
        forward = (df1 / df2) ^ (1 / delta_t(date1, date2, count)) - 1
    Else
        forward = (1 / delta_t(date1, date2, count)) * Log(df1 / df2)
    End If
    
    Tx_forward = forward

End Function

'Passage de ZC au discount factor

Function ZC_DF(ByVal date_valo As Date, ByVal date_ZC As Date, ByVal ZC As Double, ByVal conv_zc As String, ByVal count As String)
    
    Dim DF_temp As Double
    If (conv_zc = "Taux linéaires") Then
        DF_temp = 1 / (1 + ZC * delta_t(date_valo, date_ZC, count))
    ElseIf (conv_zc = "Taux actuariels") Then
        DF_temp = 1 / ((1 + ZC) ^ delta_t(date_valo, date_ZC, count))
    Else
        DF_temp = Exp(-ZC * delta_t(date_valo, date_ZC, count))
    End If
    
    ZC_DF = DF_temp

End Function

'Passage du DF au ZC

Function DF_ZC2(ByVal date_valo As Date, ByVal date_DF As Date, ByVal df As Double, ByVal conv_zc As String, ByVal count As String)

    Dim ZC_temp As Double
    Dim ecart As Double
    
    If (date_valo = date_DF) Then
        
        ZC_temp = 0
        
    Else
            ecart = delta_t(date_valo, date_DF, count)
            If (conv_zc = "Taux actuariels") Then
                ZC_temp = df ^ (-1 / ecart) - 1
            ElseIf (conv_zc = "Taux linéaires") Then
                ZC_temp = (1 / ecart) * (1 / df - 1)
            Else
                ZC_temp = -Log(df) / ecart
            End If
    End If
    
    DF_ZC2 = ZC_temp

End Function


Function Coupon(ByVal nominal As Double, _
                 ByVal taux As Double, _
                 ByVal T1 As Date, _
                 ByVal T2 As Date, _
                 ByVal DayCount As String) As Double
                 
    Coupon = nominal * taux * delta_t(T1, T2, DayCount)
                 
End Function

