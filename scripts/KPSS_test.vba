Option Explicit



'#####################################################################################################
'# KPSStest is the main function, it is explained in usage sheet.                                    #
'#####################################################################################################

Function KPSStest(TimeSeriesRange As Variant, Optional lags As Variant = "short", Optional Trend As Boolean = False)
'=====================================================================================================
' declaring variables
'=====================================================================================================
Dim T, i, j As Long
Dim S, sum_S_sqr, mean_statistic, statistic, LM, sum_sigma_sqr As Double
Dim IndependentRegressor() As Double
Dim output(1 To 1, 1 To 5) As Double
Dim confidence() As Double
Dim RegressionOutput(), residuals() As Variant
Dim SS() As Double
''''''''''''''''''''''''''''''''''''''''    CODING      ''''''''''''''''''''''''''''''''''''''''''''''
    
    T = TimeSeriesRange.Rows.Count

    'checking lags
    
    If lags = "short" Then
            lags = Round(4 * (T / 100) ^ 0.25, 0)
    ElseIf lags = "long" Then
            lags = Round(12 * (T / 100) ^ 0.25, 0)
    ElseIf IsMissing(lags) = True Then
            lags = Round(4 * (T / 100) ^ 0.25, 0)
    ElseIf IsNumeric(lags) = True Then
    
            If lags < 0 Or lags > T Then
                Exit Function
            Else
                lags = Round(lags, 0)
            End If
            
    End If
        
    If Trend = True Then
       
        ReDim IndependentRegressor(1 To T, 1 To 1)
       
        For i = 1 To T
           
           IndependentRegressor(i, 1) = i
       
        Next i
                
        RegressionOutput = Application.LinEst(TimeSeriesRange, IndependentRegressor, True, 1)
        statistic = RegressionOutput(1, 1)
        mean_statistic = RegressionOutput(1, 2)
        
        ReDim residuals(1 To T)
        ReDim SS(1 To T)
        
        For i = 1 To T
            
           residuals(i) = TimeSeriesRange(i) - mean_statistic - statistic * IndependentRegressor(i, 1)
           S = S + residuals(i)
           SS(i) = S
           
        Next i
        
        sum_S_sqr = WorksheetFunction.SumProduct(SS, SS) / (T * T)
        
        For i = 1 To lags
            For j = i + 1 To T
            
                sum_sigma_sqr = sum_sigma_sqr + (2 / T) * (1 - (i / (lags + 1))) * residuals(j) * residuals(j - i)
            
            Next j
        Next i
        
        sum_sigma_sqr = sum_sigma_sqr + (1 / T) * WorksheetFunction.SumProduct(residuals, residuals)
        
        LM = sum_S_sqr / sum_sigma_sqr
        
        output(1, 1) = LM
        output(1, 2) = lags
        output(1, 3) = 0.216
        output(1, 4) = 0.146
        output(1, 5) = 0.119
'-----------------------------------------------------------------------------------------------------
' output - a 1 x 6 array
'-----------------------------------------------------------------------------------------------------
    KPSStest = output
        
    Else
    
        ReDim residuals(1 To T)
        ReDim SS(1 To T)
        
        mean_statistic = WorksheetFunction.Average(TimeSeriesRange)

        For i = 1 To T
            
           residuals(i) = TimeSeriesRange(i) - mean_statistic
           S = S + residuals(i)
           SS(i) = S
           
           
        Next i
        
        sum_S_sqr = WorksheetFunction.SumProduct(SS, SS) / (T * T)
        
        For i = 1 To lags
            For j = i + 1 To T
            
                sum_sigma_sqr = sum_sigma_sqr + (2 / T) * (1 - (i / (lags + 1))) * residuals(j) * residuals(j - i)
            
            Next j
        Next i
        
        sum_sigma_sqr = sum_sigma_sqr + (1 / T) * WorksheetFunction.SumProduct(residuals, residuals)
        
        LM = sum_S_sqr / sum_sigma_sqr
        
        output(1, 1) = LM
        output(1, 2) = lags
        output(1, 3) = 0.739
        output(1, 4) = 0.463
        output(1, 5) = 0.347
    
'-----------------------------------------------------------------------------------------------------
' output - a 1 x 6 array
'-----------------------------------------------------------------------------------------------------
    KPSStest = output
    
    End If
    
End Function
