Attribute VB_Name = "fonctions_taux"
Option Explicit
Option Base 1

Function Affiche_Pricer_Bond_Fixe(ByVal t As Date, _
                             ByVal T1 As Date, _
                             ByVal T2 As Date, _
                             ByVal nominal As Double, _
                             ByVal taux As Double, _
                             ByVal Spread As Double, _
                             ByVal freq As Double, _
                             ByVal BrokenPeriod As String, _
                             ByVal CourbeDate As Variant, _
                             ByVal CourbeTaux As Variant, _
                             ByVal ConventionTx As String, _
                             ByVal DayCount As String, _
                             ByVal BusinessDay As String) As Variant


    
    Dim PrixObligationFixe As Double
    Dim temp() As Variant
    Dim eche() As Variant
    Dim Affiche() As Variant
    Dim DateCurve() As Variant
    Dim RateCurve() As Variant
    Dim i As Integer
    Dim j As Integer
    
    
    DateCurve = Transforme(CourbeDate)
    RateCurve = Transforme(CourbeTaux)
    PrixObligationFixe = 0
    i = 2
    j = 1
    eche = Cash_Flow_Dates(T1, T2, freq, BusinessDay, BrokenPeriod)
    Affiche = Affiche_Cash_Flow_Dates(t, T1, T2, freq, BusinessDay, BrokenPeriod)
    
    ReDim temp(UBound(Affiche) + 2, 4)
    
    While i <= UBound(eche)
    
        If eche(i) > t Then
          
            temp(j, 1) = Affiche(j)
            temp(j + 1, 2) = Coupon(nominal, taux, eche(i - 1), eche(i), DayCount)
            temp(j + 1, 3) = ZC_DF(t, eche(i), Interpolation(eche(i), DateCurve, RateCurve) + Spread, ConventionTx, DayCount)
            temp(j + 1, 4) = Coupon(nominal, taux, eche(i - 1), eche(i), DayCount) * ZC_DF(t, eche(i), Interpolation(eche(i), DateCurve, RateCurve) + Spread, ConventionTx, DayCount)
            j = j + 1
            PrixObligationFixe = PrixObligationFixe + Coupon(nominal, taux, eche(i - 1), eche(i), DayCount) * ZC_DF(t, eche(i), Interpolation(eche(i), DateCurve, RateCurve) + Spread, ConventionTx, DayCount)
        
        End If
        
        i = i + 1
      
    Wend
    
    PrixObligationFixe = PrixObligationFixe + nominal * ZC_DF(t, T2, Interpolation(T2, DateCurve, RateCurve), ConventionTx, DayCount)
    temp(UBound(Affiche), 1) = T2
    temp(j, 4) = temp(j, 4) + nominal * ZC_DF(t, T2, Interpolation(T2, DateCurve, RateCurve), ConventionTx, DayCount)
    temp(UBound(Affiche) + 1, 4) = PrixObligationFixe
    temp(UBound(Affiche) + 1, 3) = "Prix coupon plein"
    temp(UBound(Affiche) + 2, 4) = PrixObligationFixe - Coupon(nominal, taux, Affiche(1), t, DayCount)
    temp(UBound(Affiche) + 2, 3) = "Prix pied de coupon"
    Affiche_Pricer_Bond_Fixe = temp

End Function


Function Pricer_Bond_Fixe(ByVal t As Date, _
                             ByVal T1 As Date, _
                             ByVal T2 As Date, _
                             ByVal nominal As Double, _
                             ByVal taux As Double, _
                             ByVal Spread As Double, _
                             ByVal freq As Double, _
                             ByVal BrokenPeriod As String, _
                             ByVal CourbeDate As Variant, _
                             ByVal CourbeTaux As Variant, _
                             ByVal ConventionTx As String, _
                             ByVal DayCount As String, _
                             ByVal BusinessDay As String) As Variant
    
    
    
    Dim PrixObligationFixe As Double
    Dim temp() As Variant
    Dim eche() As Variant
    Dim Affiche() As Variant
    Dim DateCurve() As Variant
    Dim RateCurve() As Variant
    Dim i As Integer
    
    DateCurve = Transforme(CourbeDate)
    RateCurve = Transforme(CourbeTaux)
    PrixObligationFixe = 0
    i = 2
    eche = Cash_Flow_Dates(T1, T2, freq, BusinessDay, BrokenPeriod)
    Affiche = Affiche_Cash_Flow_Dates(t, T1, T2, freq, BusinessDay, BrokenPeriod)
    
    ReDim temp(2, 1)
    
    While i <= UBound(eche)
        If eche(i) > t Then
            PrixObligationFixe = PrixObligationFixe + Coupon(nominal, taux, eche(i - 1), eche(i), DayCount) * ZC_DF(t, eche(i), Interpolation(eche(i), DateCurve, RateCurve) + Spread, ConventionTx, DayCount)
        End If
        i = i + 1
    Wend
    
    PrixObligationFixe = PrixObligationFixe + nominal * ZC_DF(t, T2, Interpolation(T2, DateCurve, RateCurve), ConventionTx, DayCount)  'verifier l'indice mis dans la fonction dinterpolation
    
    temp(1, 1) = PrixObligationFixe
    temp(2, 1) = PrixObligationFixe - Coupon(nominal, taux, Affiche(1), t, DayCount)
    Pricer_Bond_Fixe = temp

End Function



Function Affiche_Pricer_Bond_Var(ByVal t As Date, _
                             ByVal T1 As Date, _
                             ByVal T2 As Date, _
                             ByVal nominal As Double, _
                             ByVal Spread As Double, _
                             ByVal marge As Double, _
                             ByVal freq As Double, _
                             ByVal BrokenPeriod As String, _
                             ByVal CourbeActuDate As Variant, _
                             ByVal CourbeActuTaux As Variant, _
                             ByVal courbeForwardDate As Variant, _
                             ByVal courbeForwardTaux As Variant, _
                             ByVal LastFix As Double, _
                             ByVal ConventionTx As String, _
                             ByVal DayCount As String, _
                             ByVal BusinessDay As String) As Variant
                             
    Dim CourbeTxForwardDate()  As Variant
    Dim CourbeTxForwardRate()  As Variant
    Dim CourbeTxActuDate()  As Variant
    Dim CourbeTxActuRate()  As Variant
    Dim tableau_ZC_Fwd() As Variant
    Dim tableau_DF_Fwd() As Variant
    Dim tableau_Tx_Fwd() As Variant
    Dim df1 As Double, df2 As Double
    
    
    
    Dim Affiche() As Variant
    Dim temp() As Variant
    Dim eche() As Variant
    Dim a As Double, b As Double, f As Double
    Dim PrixObligationVar As Double
    
    CourbeTxForwardDate = Transforme(courbeForwardDate)
    CourbeTxForwardRate = Transforme(courbeForwardTaux)
    CourbeTxActuDate = Transforme(CourbeActuDate)
    CourbeTxActuRate = Transforme(CourbeActuTaux)
    
    
    Dim i As Integer
    Dim j As Integer
    
    
    
    i = 2
    j = 1
    PrixObligationVar = 0
    eche = Cash_Flow_Dates(T1, T2, freq, BusinessDay, BrokenPeriod)
    Affiche = Affiche_Cash_Flow_Dates(t, T1, T2, freq, BusinessDay, BrokenPeriod)
    ReDim temp(UBound(Affiche) + 1, 5)
    
    
    If t < T1 Then 'premiere partie si la valorisation est avant la date de debut
            
            
        
        While i <= UBound(eche)
            df1 = ZC_DF(t, eche(i - 1), Interpolation(eche(i - 1), CourbeTxForwardDate, CourbeTxForwardRate), ConventionTx, DayCount)
            df2 = ZC_DF(t, eche(i), Interpolation(eche(i), CourbeTxForwardDate, CourbeTxForwardRate), ConventionTx, DayCount)
            a = Tx_forward(eche(i - 1), eche(i), df1, df2, "Taux linéaires", DayCount) + marge
            temp(j, 1) = eche(i - 1)
            temp(j, 2) = eche(i)
            temp(j, 3) = a
            temp(j, 4) = ZC_DF(t, eche(i), Interpolation(eche(i), CourbeTxActuDate, CourbeTxActuRate) + Spread, ConventionTx, DayCount)
            temp(j, 5) = Coupon(nominal, a, eche(i - 1), eche(i), DayCount) * ZC_DF(t, eche(i), Interpolation(eche(i), CourbeTxActuDate, CourbeTxActuRate) + Spread, ConventionTx, DayCount)
            j = j + 1
            PrixObligationVar = PrixObligationVar + Coupon(nominal, a, eche(i - 1), eche(i), DayCount) * ZC_DF(t, eche(i), Interpolation(eche(i), CourbeTxActuDate, CourbeTxActuRate) + Spread, ConventionTx, DayCount)
            i = i + 1
        Wend
        
        PrixObligationVar = PrixObligationVar + nominal * ZC_DF(t, T2, Interpolation(T2, CourbeTxActuDate, CourbeTxActuRate), ConventionTx, DayCount)
        temp(j - 1, 5) = temp(j - 1, 5) + nominal * ZC_DF(t, T2, Interpolation(T2, CourbeTxActuDate, CourbeTxActuRate), ConventionTx, DayCount)
        temp(UBound(Affiche), 5) = PrixObligationVar
        temp(UBound(Affiche), 4) = "Prix coupon plein"
        temp(UBound(Affiche) + 1, 5) = PrixObligationVar
        temp(UBound(Affiche) + 1, 4) = "Prix pied de coupon"
        Affiche_Pricer_Bond_Var = temp
    
    
    ElseIf t >= T1 Then 'deuxieme partie si la valoration est apres la date de debut
     
            b = marge + LastFix
            temp(j, 1) = Affiche(i - 1)
            temp(j, 2) = Affiche(i)
            temp(j, 3) = b
            temp(j, 4) = ZC_DF(t, Affiche(i), Interpolation(eche(i), CourbeTxActuDate, CourbeTxActuRate) + Spread, ConventionTx, DayCount)
            temp(j, 5) = Coupon(nominal, b, Affiche(i - 1), Affiche(i), DayCount) * ZC_DF(t, Affiche(i), Interpolation(Affiche(i), CourbeTxActuDate, CourbeTxActuRate) + Spread, ConventionTx, DayCount)
            PrixObligationVar = PrixObligationVar + Coupon(nominal, b, Affiche(i - 1), Affiche(i), DayCount) * ZC_DF(t, Affiche(i), Interpolation(Affiche(i), CourbeTxActuDate, CourbeTxActuRate) + Spread, ConventionTx, DayCount)
        j = j + 1
        i = i + 1
                 
        While i <= UBound(Affiche)
            df1 = ZC_DF(t, Affiche(i - 1), Interpolation(Affiche(i - 1), CourbeTxForwardDate, CourbeTxForwardRate), ConventionTx, DayCount)
            df2 = ZC_DF(t, Affiche(i), Interpolation(Affiche(i), CourbeTxForwardDate, CourbeTxForwardRate), ConventionTx, DayCount)
            f = Tx_forward(Affiche(i - 1), Affiche(i), df1, df2, "Taux linéaires", DayCount) + marge
            temp(j, 1) = Affiche(i - 1)
            temp(j, 2) = Affiche(i)
            temp(j, 3) = f
            temp(j, 4) = ZC_DF(t, Affiche(i), Interpolation(Affiche(i), CourbeTxActuDate, CourbeTxActuRate) + Spread, ConventionTx, DayCount)
            temp(j, 5) = Coupon(nominal, f, Affiche(i - 1), Affiche(i), DayCount) * ZC_DF(t, Affiche(i), Interpolation(Affiche(i), CourbeTxActuDate, CourbeTxActuRate) + Spread, ConventionTx, DayCount)
            PrixObligationVar = PrixObligationVar + Coupon(nominal, f, Affiche(i - 1), Affiche(i), DayCount) * ZC_DF(t, Affiche(i), Interpolation(Affiche(i), CourbeTxActuDate, CourbeTxActuRate) + Spread, ConventionTx, DayCount)
            i = i + 1
            j = j + 1
            
        Wend
        
        
        PrixObligationVar = PrixObligationVar + nominal * ZC_DF(t, T2, Interpolation(T2, CourbeTxActuDate, CourbeTxActuRate), ConventionTx, DayCount)
        temp(j - 1, 5) = temp(j - 1, 5) + nominal * ZC_DF(t, T2, Interpolation(T2, CourbeTxActuDate, CourbeTxActuRate), ConventionTx, DayCount)
        temp(UBound(Affiche), 5) = PrixObligationVar
        temp(UBound(Affiche), 4) = "Prix coupon plein"
        temp(UBound(Affiche) + 1, 5) = PrixObligationVar - Coupon(nominal, b, Affiche(1), t, DayCount)
        temp(UBound(Affiche) + 1, 4) = "Prix pied de coupon"
        Affiche_Pricer_Bond_Var = temp
        
    End If
    

      
End Function

Function Pricer_Bond_Var(ByVal t As Date, _
                             ByVal T1 As Date, _
                             ByVal T2 As Date, _
                             ByVal nominal As Double, _
                             ByVal Spread As Double, _
                             ByVal marge As Double, _
                             ByVal freq As Double, _
                             ByVal BrokenPeriod As String, _
                             ByVal CourbeActuDate As Variant, _
                             ByVal CourbeActuTaux As Variant, _
                             ByVal courbeForwardDate As Variant, _
                             ByVal courbeForwardTaux As Variant, _
                             ByVal LastFix As Double, _
                             ByVal ConventionTx As String, _
                             ByVal DayCount As String, _
                             ByVal BusinessDay As String) As Variant
                             
    Dim CourbeTxForwardDate()  As Variant
    Dim CourbeTxForwardRate()  As Variant
    Dim CourbeTxActuDate()  As Variant
    Dim CourbeTxActuRate()  As Variant
    Dim tableau_ZC_Fwd() As Variant
    Dim tableau_DF_Fwd() As Variant
    Dim tableau_Tx_Fwd() As Variant
    Dim df1 As Double, df2 As Double
    
    
    
    Dim Affiche() As Variant
    Dim temp() As Variant
    Dim eche() As Variant
    Dim a As Double, b As Double, f As Double
    Dim PrixObligationVar As Double
    
    CourbeTxForwardDate = Transforme(courbeForwardDate)
    CourbeTxForwardRate = Transforme(courbeForwardTaux)
    CourbeTxActuDate = Transforme(CourbeActuDate)
    CourbeTxActuRate = Transforme(CourbeActuTaux)
    
    
    Dim i As Integer
    Dim j As Integer
    
    
    
    i = 2
    j = 1
    PrixObligationVar = 0
    eche = Cash_Flow_Dates(T1, T2, freq, BusinessDay, BrokenPeriod)
    Affiche = Affiche_Cash_Flow_Dates(t, T1, T2, freq, BusinessDay, BrokenPeriod)
    ReDim temp(2, 1)
    
    
    If t < T1 Then 'premiere partie si la valorisation est avant la date de debut
            
            
        
        While i <= UBound(eche)
            df1 = ZC_DF(t, eche(i - 1), Interpolation(eche(i - 1), CourbeTxForwardDate, CourbeTxForwardRate), ConventionTx, DayCount)
            df2 = ZC_DF(t, eche(i), Interpolation(eche(i), CourbeTxForwardDate, CourbeTxForwardRate), ConventionTx, DayCount)
            a = Tx_forward(eche(i - 1), eche(i), df1, df2, "Taux linéaires", DayCount) + marge
           PrixObligationVar = PrixObligationVar + Coupon(nominal, a, eche(i - 1), eche(i), DayCount) * ZC_DF(t, eche(i), Interpolation(eche(i), CourbeTxActuDate, CourbeTxActuRate) + Spread, ConventionTx, DayCount)
            i = i + 1
        Wend
        
        PrixObligationVar = PrixObligationVar + nominal * ZC_DF(t, T2, Interpolation(T2, CourbeTxActuDate, CourbeTxActuRate), ConventionTx, DayCount)
        temp(1, 1) = PrixObligationVar
        temp(2, 1) = PrixObligationVar
    
        Pricer_Bond_Var = temp
    
    
    ElseIf t >= T1 Then 'deuxieme partie si la valoration est apres la date de debut
        
      
     
        b = marge + LastFix
        PrixObligationVar = PrixObligationVar + Coupon(nominal, b, Affiche(i - 1), Affiche(i), DayCount) * ZC_DF(t, Affiche(i), Interpolation(Affiche(i), CourbeTxActuDate, CourbeTxActuRate) + Spread, ConventionTx, DayCount)
        j = j + 1
        i = i + 1
                 
            
        While i <= UBound(Affiche)
            df1 = ZC_DF(t, Affiche(i - 1), Interpolation(Affiche(i - 1), CourbeTxForwardDate, CourbeTxForwardRate), ConventionTx, DayCount)
            df2 = ZC_DF(t, Affiche(i), Interpolation(Affiche(i), CourbeTxForwardDate, CourbeTxForwardRate), ConventionTx, DayCount)
            f = Tx_forward(Affiche(i - 1), Affiche(i), df1, df2, "Taux linéaires", DayCount) + marge
            PrixObligationVar = PrixObligationVar + Coupon(nominal, f, Affiche(i - 1), Affiche(i), DayCount) * ZC_DF(t, Affiche(i), Interpolation(Affiche(i), CourbeTxActuDate, CourbeTxActuRate) + Spread, ConventionTx, DayCount)
            i = i + 1
            j = j + 1
            
        Wend
        
        PrixObligationVar = PrixObligationVar + nominal * ZC_DF(t, T2, Interpolation(T2, CourbeTxActuDate, CourbeTxActuRate), ConventionTx, DayCount)
        temp(1, 1) = PrixObligationVar
        temp(2, 1) = PrixObligationVar - Coupon(nominal, b, Affiche(1), t, DayCount)
    
        Pricer_Bond_Var = temp
        
    End If
End Function



Public Function jambe_variable_fixe(ByVal t As Date, _
                             ByVal T1 As Date, _
                             ByVal T2 As Date, _
                             ByVal nominal As Double, _
                             ByVal Spread As Double, _
                             ByVal marge As Double, _
                             ByVal freq As Double, _
                             ByVal BrokenPeriod As String, _
                             ByVal CourbeActuDate As Variant, _
                             ByVal CourbeActuTaux As Variant, _
                             ByVal courbeForwardDate As Variant, _
                             ByVal courbeForwardTaux As Variant, _
                             ByVal LastFix As Double, _
                             ByVal ConventionTx As String, _
                             ByVal DayCount As String, _
                             ByVal BusinessDay As String, _
                             ByVal CourbeDate As Variant, _
                             ByVal CourbeTaux As Variant, _
                             ByVal taux As Double) As Variant

    Dim jambe_f As Double
    Dim jambe_v As Double
    Dim prix As Variant
    ReDim Preserve prix(2, 1)
    
    Spread = 0
    jambe_f = Pricer_Bond_Fixe(t, T1, T2, nominal, taux, Spread, freq, BrokenPeriod, CourbeDate, CourbeTaux, ConventionTx, DayCount, BusinessDay)(1, 1)
    jambe_v = Pricer_Bond_Var(t, T1, T2, nominal, Spread, marge, freq, BrokenPeriod, CourbeActuDate, CourbeActuTaux, courbeForwardDate, courbeForwardTaux, LastFix, ConventionTx, DayCount, BusinessDay)(1, 1)
    
    prix(1, 2) = jambe_v - jambe_f
    prix(1, 1) = "payeuse taux fixe"
    
    prix(2, 2) = jambe_v - jambe_f
    prix(2, 1) = "payeuse taux variable"
    
    jambe_variable_fixe = prix

End Function


Public Function jambe_variable_variable(t As Date, _
                             T1 As Date, _
                             T2 As Date, _
                             nominal As Double, _
                             Spread As Double, _
                             marge As Double, _
                             freq As Double, _
                             BrokenPeriod As String, _
                             CourbeActuDate As Variant, _
                             CourbeActuTaux As Variant, _
                             courbeForwardDate As Variant, _
                             courbeForwardTaux_1 As Variant, _
                             courbeForwardTaux_2 As Variant, _
                             LastFix As Double, _
                             ConventionTx As String, _
                             DayCount As String, _
                             BusinessDay As String) As Variant

    Dim jambe_v1 As Double
    Dim jambe_v2 As Double
    Dim prix As Variant
    ReDim Preserve prix(2, 2)
    
    Spread = 0
    jambe_v1 = Pricer_Bond_Var(t, T1, T2, nominal, Spread, marge, freq, BrokenPeriod, CourbeActuDate, CourbeActuTaux, courbeForwardDate, courbeForwardTaux_1, LastFix, ConventionTx, DayCount, BusinessDay)(1, 1)
    jambe_v2 = Pricer_Bond_Var(t, T1, T2, nominal, Spread, marge, freq, BrokenPeriod, CourbeActuDate, CourbeActuTaux, courbeForwardDate, courbeForwardTaux_2, LastFix, ConventionTx, DayCount, BusinessDay)(1, 1)
    
    prix(1, 2) = jambe_v2 - jambe_v1
    prix(1, 1) = "payeuse taux var 1"
    
    prix(2, 2) = jambe_v1 - jambe_v2
    prix(2, 1) = "payeuse taux var 2"
    
    jambe_variable_variable = prix

End Function


Public Function fx_forward(devise As String, _
                                    date_valo As Date, _
                                    date_matu As Date, _
                                    strike As Double, _
                                    taux_change As Double, _
                                    nominal As Integer, _
                                    courbeTaux_valo_dom As Variant, _
                                    report_deport As Variant, _
                                    ConventionTx As String, _
                                    DayCount As String) As Double

    Dim DF_DDD As Variant
    Dim DF_EEE As Variant
    
    Dim val1 As Double
    Dim val2 As Double
    Dim i As Integer
    Dim nouv_date As Date
    Dim prix_forward As Variant
    
    For i = 0 To 10
    
        nouv_date = DateAdd("yyyy", i, date_matu)
    
        val1 = ZC_DF(date_valo, nouv_date, courbeTaux_valo_dom(nouv_date), ConventionTx, DayCount)
        val2 = taux_change * val1
        DF_DDD(i) = val1
        DF_EEE(i) = val2
        
        prix_forward(i) = taux_change + (report_deport(i) / 10000)
        
    Next
    'pareil tableau avec differentes valeurs pour differentes dates, matuité à laquelle on ajoute 1 an à chaque fois
    
    Dim prix_contrat As Variant
    ReDim Preserve prix_contrat(2, 10)
    Dim contrat As Variant
    
    prix_contrat(1, 1) = "acheteur"
    prix_contrat(2, 1) = "vendeur"
    
    Dim valeur1 As Double
    Dim valeur2 As Double
    
    For i = 2 To 10 'à changer
    
    
        nouv_date = DateAdd("dd", i, date_matu)
        'tableau avec différentes maturités
        
        valeur1 = (nominal / strike) * DF_DDD(i) * (prix_forward(i) - strike)
        valeur2 = -valeur1
        
        prix_contrat(1, i) = valeur1
        prix_contrat(2, i) = valeur2
    
        'df_ddd va varier en fonction de chaque date, la date de matu est accrémentée dannée en année pendant 10 ans
        ' on récupère le fx forward pour toutes le smaturités allant de 1 à 10 ans
        contrat(i) = prix_contrat
    Next
    
    fx_forward = contrat
    
End Function


Public Function fx_call_put(devise As String, _
                                    date_valo As Date, _
                                    date_matu As Date, _
                                    strike As Double, _
                                    taux_change As Double, _
                                    nominal As Integer, _
                                    courbeTaux_valo_dom As Variant, _
                                    report_deport As Variant, _
                                    ConventionTx As String, _
                                    DayCount As String, _
                                    volatility As Double) As Variant

    Dim d1 As Double, d2 As Double
    Dim Call_option_price As Variant, put_option_price As Variant
    Dim fx()
    ReDim Preserve fx(10, 2)

    Dim DF_DDD As Variant
    Dim DF_EEE As Variant
    
    Dim val1 As Double
    Dim val2 As Double
    Dim i As Integer, j As Integer, k As Integer
    Dim nouv_date As Date
    
    
    For i = 0 To 10
        nouv_date = DateAdd("yyyy", i, date_matu)
    
        val1 = ZC_DF(date_valo, nouv_date, courbeTaux_valo_dom(nouv_date), ConventionTx, DayCount)
        val2 = taux_change * val1
        DF_DDD(i) = val1
        DF_EEE(i) = val2
    Next


    For j = 0 To 10
        nouv_date = DateAdd("yyyy", j, date_matu)
        
        d1 = (Log((DF_EEE(j) / DF_DDD(j)) * (taux_change / strike)) + 0.5 * delta_t(date_valo, nouv_date, DayCount) * volatility ^ 2) / volatility * Sqr(delta_t(date_valo, nouv_date, DayCount))
        d2 = d1 - volatility * Sqr(delta_t(date_valo, nouv_date, DayCount))
        
        Call_option_price(j) = DF_EEE(j) * taux_change * WorksheetFunction.Norm_Dist(d1, 0, 0, False) - DF_DDD(j) * strike * WorksheetFunction.Norm_Dist(d2, 0, 0, False)
        put_option_price(j) = -DF_EEE(j) * taux_change * WorksheetFunction.Norm_Dist(-d1, 0, 0, False) + DF_DDD(j) * strike * WorksheetFunction.Norm_Dist(-d2, 0, 0, False)
    Next
    
    
    For k = 0 To 10
        fx(k + 1, 1) = Call_option_price(k)
        fx(k + 1, 2) = put_option_price(k)
    Next k
    
    fx_call_put = fx
    
End Function


Public Function interest_rate_cap_floor()




End Function











'-------------------------------------------------- FONCTIONS ANNEXES --------------------------------------


Public Function min(a As Double, b As Double) As Double

    Dim minimum As Double
    
    If a < b Then
        minimum = a
    Else
        minimum = b
    End If
    
    min = minimum

End Function

Public Function max(c As Double, d As Double) As Double
    
    Dim maximum As Double
    
    If c < d Then
        maximum = d
    Else
        maximum = c
    End If
    
    max = maximum

End Function

Public Function appartenance(mot As String, lettre As String) As Boolean
    
    Dim i As Integer
    i = 0
    Dim appart As Boolean
    appart = False
    
    For i = 0 To Len(mot)
    
     If Mid(mot, i, 1) = lettre Then
        appart = True
     End If
     
    Next
    
    appartenance = appart

End Function

Public Function nombre_e(mot As String) As Integer
    
    Dim S As Integer, i As Integer
    S = 0
    
    For i = 0 To Len(mot)
        If Mid(mot, i, 1) = "e" Then
            S = S + 1
        End If
    Next
    
    nombre_e = S

End Function


Public Function palindrome(mot As String) As Boolean

    Dim i As Integer
    Dim est_palindrome As Boolean
    est_palindrome = False
    
    Dim longueur As Integer
    longueur = Len(mot)
    
    For i = 0 To Int(longueur / 2)
        While Mid(mot, i, 1) = Mid(mot, longueur - i, 1)
            est_palindrome = True
        Wend
        est_palindrome = False
    Next
    
    palindrome = est_palindrome

End Function
