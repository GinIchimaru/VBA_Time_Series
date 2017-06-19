Option Explicit


'#####################################################################################################
'# ADFtest is the main function, it is explained in usage sheet. It is mainly, a case statement of   #
'# ADFRegression function depending on lag and LagCriteria arguments. For the time being theese are  #
'# defined as variants. Basically it's a wrapper function                                            #
'#####################################################################################################
Function ADFtest(TimeSeriesRange As Range, Optional Lagg As Variant, Optional LagCriteria As Variant, _
Optional Intercept As Boolean = True, Optional Trend As Boolean = False)
'=====================================================================================================
' declaring variables
'=====================================================================================================

Dim Lag As Long
Dim T As Long
Dim TemporaryArray() As Double
Dim tmp() As Double
Dim i As Long, j As Long
Dim indx As Long
Dim indx1 As Long
Dim TemporaryColumn() As Variant
Dim MinIC As Double
Dim col As Integer
''''''''''''''''''''''''''''''''''''''     CODING     ''''''''''''''''''''''''''''''''''''''''''''''''
    If IsMissing(Lagg) = False Then
        Lag = Lagg
    End If
    
    If Trend = True Then
        Intercept = True 'this must be the case since we can have trend and no intercept
    End If

    'count the length of time series
    T = TimeSeriesRange.Rows.Count

    'if lag argument is missing and lag is not determined by information criteria
    'then use Schwer suggestion for lag length
    If IsMissing(LagCriteria) Then
        If IsMissing(Lagg) Then
        'if lag criteria is missing but we have specified the lag then
            Lag = Round(12 * (T / 100) ^ (1 / 4), 0) 'suggested by Schwert(1989)
            LagCriteria = "AIC"
GoTo continue
        Else 'return regression with that lag
            ADFtest = ADFRegression(TimeSeriesRange, Lag, Intercept, Trend)
        End If
    'if lag criteria is not missing then->
    ElseIf (LagCriteria = "AIC" Or LagCriteria = "BIC") Then
        If IsMissing(Lagg) = False Then '-> and lag is not missing then->
continue:
            If LagCriteria = "AIC" Then '->depending on choise (aic->col=5, bic->col=6) set up a column->
                col = 6
            Else
                col = 7
            End If
            
            ReDim TemporaryArray(0 To Lag, 1 To 9) '->and make a temp array of n=lag regressions results
            
            For i = 0 To Lag
                tmp = ADFRegression(TimeSeriesRange, i, Intercept, Trend, Lag - i) 'lag-i is triming'
                For j = 1 To 9
                    TemporaryArray(i, j) = tmp(1, j)
                Next j
            Next i
            'find the row index of maximum aic (col=5) or bic (col=6) criteria from column col of temoarray and ->
            TemporaryColumn = Application.WorksheetFunction.Index(TemporaryArray, 0, col)
            MinIC = Application.WorksheetFunction.Min(TemporaryColumn)
            indx = Application.WorksheetFunction.Match(MinIC, TemporaryColumn, 0)
            indx1 = indx - 1 'indx doesnt count from 0, instead it counts from 1, so 1 is 0 lag
            '->take the indx1 as lag and show the result
            ADFtest = ADFRegression(TimeSeriesRange, indx1, Intercept, Trend, 0)
            'ADFtest = Application.WorksheetFunction.Index(TemporaryArray, indx, 0)
        Else 'if lag criteria is not missing and lag is missing
        
            ADFtest = CVErr(xlErrName)
        
        End If
    'if lag criteria is wrong then assume
    ElseIf LagCriteria <> "AIC" Or LagCriteria <> "BIC" Then
        
        ADFtest = CVErr(xlErrName)
        
    Else
    ADFtest = CVErr(xlErrNA)
    'Range("result") = ADFtest
    
    End If
      
'-----------------------------------------------------------------------------------------------------
' output - a 1 x 9 array
'-----------------------------------------------------------------------------------------------------
End Function

'#####################################################################################################
'# PPtest is explained in usage sheet. It is mainly my own work but I did   #
'# borrow some of the code from R function ur.pp from urca package mainly as a check but some parts  #
'# were completely translated from R since I had a bug in my code and this way was much easier.      #
'#                                                                                                   #
'# For R ur.pp function see: https://cran.r-project.org/web/packages/urca/index.html                 #
'#####################################################################################################

Function PPtest(TimeSeriesRange As Range, Optional lags As Variant = "short", Optional Intercept As Boolean = True, Optional Trend As Boolean = False)
'=====================================================================================================
' declaring variables
'=====================================================================================================
Dim T, i, j, ub, n As Long
Dim DependentRegressor() As Double
Dim IndependentRegressor() As Double
Dim confidence() As Double
Dim RegressionOutput(), residuals() As Variant
Dim lambda, statistic, meanstatistic, sigma_sqr_e, first_sum, right_summand, sigma_sqr As Double
Dim z_rho, z_tau, sigma_statistic, t_statistic, mean, myybar, myy, mty, my, M As Double
Dim output(1 To 1, 1 To 6) As Double
''''''''''''''''''''''''''''''''''''''''    CODING      ''''''''''''''''''''''''''''''''''''''''''''''

    T = TimeSeriesRange.Rows.Count
    n = T - 1
    
    'checking lags

    If lags = "short" Then
            lags = Round(4 * (n / 100) ^ 0.25, 0)
    ElseIf lags = "long" Then
            lags = Round(12 * (n / 100) ^ 0.25, 0)
    ElseIf IsMissing(lags) = True Then
            lags = Round(4 * (n / 100) ^ 0.25, 0)
    ElseIf IsNumeric(lags) = True Then
    
            If lags < 0 Or lags > n Then
                Exit Function
            Else
                lags = Round(lags, 0)
            End If
            
    End If
    
    'checking trend
    If Trend = True Then
        Intercept = True 'this must be the case since we can have trend and no intercept
    End If
    
    'confidence intervals
    confidence = MacKinnon(n, Intercept, Trend)
    
    'creating lagged time series-dependent variable
    ReDim DependentRegressor(1 To n, 1 To 1)
    DependentRegressor = LagTimeSeries(TimeSeriesRange, 1)
        
    'creating differenced time series-independent variable
    ReDim IndependentRegressor(1 To n, 1 To 1)
    IndependentRegressor = LagTimeSeries(TimeSeriesRange, -1)
    
    'creating trend if Trend = true
    If Trend = True Then
        
        ReDim Preserve IndependentRegressor(1 To n, 1 To 2)
        
        For i = 1 To n
            IndependentRegressor(i, 2) = i
        Next i
        
    End If
    
    'running a regression
    RegressionOutput = Application.LinEst(DependentRegressor, IndependentRegressor, Intercept, 1)
    ub = UBound(RegressionOutput, 2)
    sigma_statistic = RegressionOutput(2, ub - 1)
    statistic = RegressionOutput(1, ub - 1)  'temporary variable, for convenience
    If Intercept = True Then
    meanstatistic = RegressionOutput(1, ub) / RegressionOutput(2, ub)
    End If
    t_statistic = (statistic - 1) / sigma_statistic

    ReDim residuals(1 To n)
    
    'calculating statistic depending on the case chosen, ie. trend and mean, only mean or no mean
    If Intercept = True And Trend = True Then
    
            mean = WorksheetFunction.Average(DependentRegressor)
            myybar = 0
            
            For i = 1 To n
                'creating residual series and myybar
                residuals(i) = DependentRegressor(i, 1) - (RegressionOutput(1, ub) + RegressionOutput(1, ub - 1) * IndependentRegressor(i, 1) _
                + RegressionOutput(1, ub - 2) * IndependentRegressor(i, 2))
                
                myybar = myybar + (1 / n ^ 2) * (DependentRegressor(i, 1) - mean) ^ 2
            Next i
            
            'creating myy
            myy = (1 / n ^ 2) * WorksheetFunction.SumProduct(DependentRegressor, DependentRegressor)
            
            'creating mty
            mty = (n ^ (-5 / 2)) * WorksheetFunction.SumProduct(WorksheetFunction.Index(IndependentRegressor, 0, 2), DependentRegressor)
            
            'creating my
            my = (n ^ (-3 / 2)) * WorksheetFunction.Sum(DependentRegressor)
            
            'creating sigma_sqr_e
            sigma_sqr_e = WorksheetFunction.SumProduct(residuals, residuals) / (n)
        
            'creating left summand of sigma_sqr
            right_summand = 0
            For j = 1 To lags
                first_sum = 0
                For i = j + 1 To n
                
                    first_sum = first_sum + residuals(i) * residuals(i - j)
                
                Next i
                right_summand = right_summand + (2 / n) * (1 - (j / (lags + 1))) * first_sum
            Next j
        
            'creating sigma squared
            
            sigma_sqr = sigma_sqr_e + right_summand 'sigma_sqr_e is left summand
            
            'creating lambda
            lambda = (sigma_sqr - sigma_sqr_e) / 2
            
            'creating M
            M = (1 - n ^ (-2)) * myy - 12 * mty ^ 2 + 12 * (1 + 1 / n) * _
            mty * my - (4 + 6 / n + 2 / n ^ 2) * my ^ 2
            
            z_rho = n * (statistic - 1) - lambda / M
            z_tau = Sqr(sigma_sqr_e / sigma_sqr) * t_statistic - lambda / Sqr(sigma_sqr * M)
'-----------------------------------------------------------------------------------------------------
' output - a 1 x 6 array
'-----------------------------------------------------------------------------------------------------
            output(1, 1) = z_tau
            output(1, 2) = statistic - 1
            output(1, 3) = lags
            output(1, 4) = confidence(1, 1)
            output(1, 5) = confidence(1, 2)
            output(1, 6) = confidence(1, 3)
            
            PPtest = output
            
            Erase DependentRegressor, IndependentRegressor, RegressionOutput, residuals
    
    Else
    
        If Intercept = True Then
        
            mean = WorksheetFunction.Average(DependentRegressor)
            myybar = 0
            
            For i = 1 To n
                'creating residual series and myybar
                residuals(i) = DependentRegressor(i, 1) - (RegressionOutput(1, ub) + RegressionOutput(1, ub - 1) * IndependentRegressor(i, 1))
                myybar = myybar + (1 / n ^ 2) * (DependentRegressor(i, 1) - mean) ^ 2
            Next i
            
            'creating sigma_sqr_e
            sigma_sqr_e = WorksheetFunction.SumProduct(residuals, residuals) / (n)
            
            'creating left summand of sigma_sqr
            right_summand = 0
            For j = 1 To lags
                first_sum = 0
                For i = j + 1 To n
                
                    first_sum = first_sum + residuals(i) * residuals(i - j)
                
                Next i
                right_summand = right_summand + (2 / n) * (1 - (j / (lags + 1))) * first_sum
            Next j
        
            'creating sigma squared
            
            sigma_sqr = sigma_sqr_e + right_summand 'sigma_sqr_e is left summand
            
            'creating lambda
            lambda = (sigma_sqr - sigma_sqr_e) / 2
            
            
            z_rho = n * (statistic - 1) - lambda / myybar
            z_tau = Sqr(sigma_sqr_e / sigma_sqr) * t_statistic - lambda / Sqr(sigma_sqr * myybar)
'-----------------------------------------------------------------------------------------------------
' output - a 1 x 6 array
'-----------------------------------------------------------------------------------------------------
            output(1, 1) = z_tau
            output(1, 2) = statistic - 1
            output(1, 3) = lags
            output(1, 4) = confidence(1, 1)
            output(1, 5) = confidence(1, 2)
            output(1, 6) = confidence(1, 3)
            
            PPtest = output
            
            Erase DependentRegressor, IndependentRegressor, RegressionOutput, residuals
            
        Else
            
            myybar = 0
            
            For i = 1 To n
                'creating residual series and myybar
                residuals(i) = DependentRegressor(i, 1) - (RegressionOutput(1, ub) + RegressionOutput(1, ub - 1) * IndependentRegressor(i, 1))
                myybar = myybar + (1 / n ^ 2) * (DependentRegressor(i, 1)) ^ 2
            Next i
            
            'creating sigma_sqr_e
            sigma_sqr_e = WorksheetFunction.SumProduct(residuals, residuals) / (n)
            
            'creating left summand of sigma_sqr
            right_summand = 0
            For j = 1 To lags
                first_sum = 0
                For i = j + 1 To n
                
                    first_sum = first_sum + residuals(i) * residuals(i - j)
                
                Next i
                right_summand = right_summand + (2 / n) * (1 - (j / (lags + 1))) * first_sum
            Next j
        
            'creating sigma squared
            
            sigma_sqr = sigma_sqr_e + right_summand 'sigma_sqr_e is left summand
            
            'creating lambda
            lambda = (sigma_sqr - sigma_sqr_e) / 2
            
            
            z_rho = n * (statistic - 1) - lambda / myybar
            z_tau = Sqr(sigma_sqr_e / sigma_sqr) * t_statistic - lambda / Sqr(sigma_sqr * myybar)
'-----------------------------------------------------------------------------------------------------
' output - a 1 x 6 array
'-----------------------------------------------------------------------------------------------------
            output(1, 1) = z_tau
            output(1, 2) = statistic - 1
            output(1, 3) = lags
            output(1, 4) = confidence(1, 1)
            output(1, 5) = confidence(1, 2)
            output(1, 6) = confidence(1, 3)
            
            PPtest = output
            Erase DependentRegressor, IndependentRegressor, RegressionOutput, residuals, confidence
            
        End If
    
    End If
    

End Function

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




'#####################################################################################################
'# Function that uses built in linnest function to calculate t statistic for ADF regression          #
'#####################################################################################################

Private Function ADFRegression(TimeSeries As Range, Lag As Long, Optional Intercept As Boolean = False, Optional Trend As Boolean = False, Optional trim As Long = 0) As Double()
'=====================================================================================================
' declaring variables
'=====================================================================================================
Dim T As Long
Dim i, j, n, k As Long
Dim DifferencedTimeSeriesArray() As Double
Dim LaggedTimeSeries() As Double
Dim DependentRegressor() As Double
Dim IndependentRegressors() As Double
Dim RegressionOutput() As Variant
Dim statistic As Double
Dim Pi As Double
Dim SumOfSquares As Double
Dim AIC, BIC, LogLik As Double
Dim output(1 To 1, 1 To 9) As Double
Dim confidence() As Double
Dim ub As Integer
Pi = Application.WorksheetFunction.Pi() 'the value of pi number

''''''''''''''''''''''''''''''''''''''''    CODING      '''''''''''''''''''''''''''''''''''''''''''''''

T = TimeSeries.Rows.Count

ReDim DependentRegressor(1 To T - 1 - Lag - trim, 1 To 1)
ReDim IndependentRegressors(1 To T - 1 - Lag - trim, 0 To Lag - Trend) 'value of boolean True is -1!-this goes for trend

    'create differenced time series
    DifferencedTimeSeriesArray = DifferenceTimeSeries(TimeSeries)
    ' create lagged time series
    LaggedTimeSeries = LagTimeSeries(TimeSeries, -1)
   
    'loop creates regressors for DF regression with corresponding lags and number of observations
    For i = 0 To Lag 'lagged level series (i=0) plus lagged differenced series delta1(i=1) + delta2...+ deltalag+trim
        For j = 1 To T - 1 - Lag - trim 'the length of tseries due to lag, difference and trim
            If i = 0 Then ' create lagged level and dependent differenced series
                DependentRegressor(j, 1) = DifferencedTimeSeriesArray(j + Lag - i + trim, 1) 'creating left side of regression
                IndependentRegressors(j, i) = LaggedTimeSeries(j + Lag - i + trim, 1) 'creating lagged in level variable as first regressor
            Else
                IndependentRegressors(j, i) = DifferencedTimeSeriesArray(j + Lag - i + trim, 1) 'creating differenced lagged series as rest of regressors
            End If
        Next j
    Next i
    
    ' if trend is present it will be created as last column of IndependentRegressors array
    If Trend = True Then
        For i = 1 To T - 1 - Lag - trim
            IndependentRegressors(i, Lag - Trend) = i
        Next i
    End If
    
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '''''''''       calculating regression and statistics      '''''''''''''''''''''''''''''''''''''''''
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    RegressionOutput = Application.LinEst(DependentRegressor, IndependentRegressors, Intercept, 1) 'temporary variable, for convenience
    ub = UBound(RegressionOutput, 2) - 1
    statistic = RegressionOutput(1, ub) / RegressionOutput(2, ub) 'temporary variable, for convenience
    SumOfSquares = RegressionOutput(5, 2) 'temporary variable, for convenience
    n = T - 1 - Lag - trim 'number of observations, temporary variable, for convenience
    k = 1 - Intercept - Trend + Lag 'number of parameters, 2 = variance + laged level variable? when comparing with eviews it was 1 so kept it that way
    LogLik = -(n / 2) * Application.Ln(SumOfSquares / n) - (n / 2) * Application.Ln(2 * Pi) - (n / 2) 'log likelihood
    AIC = 2 * (k / n) - 2 * LogLik / n '+ (2 * k * (k + 1)) / (n * (n - k - 1)) 'Akaike --AICc
    BIC = k * Application.Ln(n) / n - 2 * LogLik / n ' Schwarz criterion
    confidence = MacKinnon(T, Intercept, Trend)
    
    output(1, 1) = statistic
    output(1, 2) = RegressionOutput(1, ub)
    output(1, 3) = confidence(1, 1)
    output(1, 4) = confidence(1, 2)
    output(1, 5) = confidence(1, 3)
    output(1, 6) = AIC
    output(1, 7) = BIC
    output(1, 8) = Lag
    output(1, 9) = Application.Count(DependentRegressor)
    'erasing variables
    Erase confidence, RegressionOutput, DependentRegressor, _
    DependentRegressor, LaggedTimeSeries, DifferencedTimeSeriesArray
    trim = 0
    
    ADFRegression = output
'-----------------------------------------------------------------------------------------------------
' output - a 1 x 8 array
'-----------------------------------------------------------------------------------------------------
End Function

'#####################################################################################################
'# For a given input arguments it gives MacKinnon t values for 3 confidence levels                   #
'#####################################################################################################

Private Function MacKinnon(T As Long, Intercept As Boolean, Trend As Boolean) As Double()

Dim beta(1 To 3, 1 To 3) As Double
Dim tau(1 To 1, 1 To 3) As Double 'one for each confidence

If Intercept = False And Trend = False Then
    '1% confidence
    beta(1, 1) = -2.5658
    beta(2, 1) = -1.96
    beta(3, 1) = -10.04
    '5% confidence
    beta(1, 2) = -1.9393
    beta(2, 2) = -0.398
    beta(3, 2) = 0
    '10% confidence
    beta(1, 3) = -1.6156
    beta(2, 3) = -0.181
    beta(3, 3) = 0
ElseIf Intercept = True And Trend = False Then
    '1% confidence
    beta(1, 1) = -3.4336
    beta(2, 1) = -5.999
    beta(3, 1) = -29.25
    '5% confidence
    beta(1, 2) = -2.8621
    beta(2, 2) = -2.738
    beta(3, 2) = -8.36
    '10% confidence
    beta(1, 3) = -2.5671
    beta(2, 3) = -1.438
    beta(3, 3) = -4.48
ElseIf Intercept = True And Trend = True Then
    '1% confidence
    beta(1, 1) = -3.9638
    beta(2, 1) = -8.353
    beta(3, 1) = -47.44
    '5% confidence
    beta(1, 2) = -3.4126
    beta(2, 2) = -4.039
    beta(3, 2) = -17.83
    '10% confidence
    beta(1, 3) = -3.1279
    beta(2, 3) = -2.418
    beta(3, 3) = -7.58
End If

tau(1, 1) = beta(1, 1) + (beta(2, 1) / T) + (beta(3, 1) / (T * T))
tau(1, 2) = beta(1, 2) + (beta(2, 2) / T) + (beta(3, 2) / (T * T))
tau(1, 3) = beta(1, 3) + (beta(2, 3) / T) + (beta(3, 3) / (T * T))
MacKinnon = tau
'-----------------------------------------------------------------------------------------------------
' output - a 1 x 3 array
'-----------------------------------------------------------------------------------------------------
End Function

'#####################################################################################################
'# Function that translates range into array. Although this can be done with variant variable in a   #
'# single line without looping trough range, I prefer to work with double type arrays. data variable   #
'# is declared like variant for convenince so that same function can be used in sheet and in vba.    #
'#####################################################################################################
Private Function RangeToArray(data As Variant)
'=====================================================================================================
' declaring variables
'=====================================================================================================
Dim TempArray() As Double
Dim T As Long
Dim Cols As Long
Dim i As Long
Dim j As Long

   T = data.Rows.Count
   Cols = data.Columns.Count
    
ReDim TempArray(1 To T, 1 To Cols)
'looping
        For i = 1 To T
            For j = 1 To Cols
                TempArray(i, j) = data(i, j).Value
            Next j
        Next i
   
    RangeToArray = TempArray
    
'-----------------------------------------------------------------------------------------------------
' Reading data this way is approximately 0.2 seconds slower for every 15000 rows (read it somewhere).
' If data is 1 column range, the rusulting array is Tx1 array
'-----------------------------------------------------------------------------------------------------
End Function


'#####################################################################################################
'# Function uses given time series of length T as input and returns T-1 differenced time series      #
'#####################################################################################################
Private Function DifferenceTimeSeries(data As Variant)
'=====================================================================================================
' declaring variables
'=====================================================================================================
Dim T As Long
Dim i As Long
Dim TimeSeries() As Double
Dim TempArray() As Double

'Checking whether data is a range or another array and based on that assinghing new variables,
'isobject function is used. This way function can be used in vba and in worksheet.
    If IsObject(data) Then
        T = data.Rows.Count
        TimeSeries = RangeToArray(data)
    Else
        T = UBound(data)
        TimeSeries = data
    End If
        
ReDim TempArray(1 To T - 1, 1 To 1)

'looping
    For i = 1 To T - 1
       TempArray(i, 1) = TimeSeries(i + 1, 1) - TimeSeries(i, 1)
    Next i
    

    DifferenceTimeSeries = TempArray
'-----------------------------------------------------------------------------------------------------
' output - a (T-1) x 1 array of double.
'-----------------------------------------------------------------------------------------------------
End Function



'#####################################################################################################
'# Function uses given time series as input on position one and returns lagged time series depending #
'# on the argument on position two.                                                                  #
'#####################################################################################################
Private Function LagTimeSeries(data As Variant, Lag As Long)
'=====================================================================================================
' declaring variables
'=====================================================================================================
Dim T As Long
Dim i As Long
Dim TimeSeries() As Double
Dim TemporaryArray() As Double
Dim col As Long
Dim ii As Integer

'Checking whether data is a range or another array
    If IsObject(data) Then
        T = data.Rows.Count
        col = data.Columns.Count
        TimeSeries = RangeToArray(data)
    Else
        T = UBound(data)
        col = UBound(data, 2)
        TimeSeries = data
    End If
    
'generaly we will be using only negative values of lag in autoregresion, but in a general case one
'can use positive and negative values of lag. The convention used here is that -1 is interpreted as
'one period earlier.Ex: TimeSeries(-1) is 1 period lag time series and we will use function as:
'===========|            TimeSeries(-1) = LagTimeSeries(TimeSeries,-1)        |=======================

    If Lag <= 0 Then   'in case we are lagging time series
        ReDim TemporaryArray(1 To T + Lag, 1 To col) As Double
        For i = 1 To T + Lag
            For ii = 1 To col
            TemporaryArray(i, ii) = TimeSeries(i, ii)
            Next ii
        Next i
    Else 'we will be working only with negative values so this is unnecessary but anyway..
        ReDim TemporaryArray(1 To T - Lag, 1 To col) As Double
        For i = Lag + 1 To T
            For ii = 1 To col
            TemporaryArray(i - Lag, ii) = TimeSeries(i, ii)
            Next ii
        Next i
    End If
    
    Erase TimeSeries 'errasing variables
    
    LagTimeSeries = TemporaryArray
'-----------------------------------------------------------------------------------------------------
' output - a (T-1) x 1 array of double.
'-----------------------------------------------------------------------------------------------------
End Function



