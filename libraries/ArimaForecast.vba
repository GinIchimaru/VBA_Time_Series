Option Explicit

Function FORECAST_ARMA(TimeSeriesRange As Variant, _
                        lags As Variant, _
                        params As Variant, _
                        Optional nAhead As Variant = 5)
    
    Dim sigma                       As Double
    Dim confidence()                As Double
    Dim fitted()                    As Double
    Dim errors()                    As Double
    Dim independent_variables()     As Double
    Dim dependent_variables()       As Double
    Dim constant                    As Double
    Dim AR_lag                      As Long
    Dim MA_lag                      As Long
    Dim T                           As Long
    Dim i, j                        As Long
    Dim AR_coef()                   As Double
    Dim MA_coef()                   As Double
    Dim AR_sum                      As Double
    Dim MA_sum                      As Double

T = TimeSeriesRange.Rows.Count
AR_lag = lags(1)
MA_lag = lags(2)
constant = params(1)

If (TypeName(params) = "Range") Then
'if yes then it must be eather one row or one column range
    If params.Rows.Count > 2 Or params.Columns.Count > 2 Or (params.Rows.Count + params.Columns.Count) > 3 Then
        FORECAST_ARMA = CVErr(xlErrValue)
    End If
ElseIf NumberOfDimensions(params) > 1 Then
        FORECAST_ARMA = CVErr(xlErrValue)
        Exit Function
End If

If (TypeName(lags) = "Range") Then
'if yes then it must be eather one row or one column range
    If lags.Rows.Count > 2 Or lags.Columns.Count > 2 Or (lags.Rows.Count + lags.Columns.Count) > 3 Then
        FORECAST_ARMA = CVErr(xlErrValue)
    End If
Else
    If NumberOfDimensions(lags) > 1 Or UBound(lags) > 2 Or lags(1) < 0 Or lags(2) < 0 Then
       FORECAST_ARMA = CVErr(xlErrValue)
       Exit Function
    End If
End If

'if more than one column then exit
If TimeSeriesRange.Columns.Count > 1 Then
    FORECAST_ARMA = CVErr(xlErrNA)
    Exit Function
End If

ReDim independent_variables(1 To T)
ReDim dependent_variables(1 To T - AR_lag)

With Application.WorksheetFunction
    For i = 1 To T
      independent_variables(i) = .Index(TimeSeriesRange, i)
    Next i
    For i = 1 To T - AR_lag
      dependent_variables(i) = .Index(TimeSeriesRange, i + AR_lag)
    Next i
End With


errors = error(params, independent_variables, dependent_variables)

If AR_lag > 0 Then
    ReDim R_coef(1 To AR_lag)
    For i = 2 To (2 + AR_lag - 1)
        R_coef(i - 1) = params(i)
    Next i
Else
    ReDim R_coef(1 To 1)
    R_coef(1) = 0
End If

'create MA coefs
If MA_lag > 0 Then
    ReDim e_coef(1 To MA_lag)
    For i = 2 + AR_lag To (2 + AR_lag + MA_lag - 1)
        e_coef(i - 1 - AR_lag) = params(i)
    Next i
Else
    ReDim e_coef(1 To 1)
    e_coef(1) = 0
End If

'create AR terms
If AR_lag > 0 Then
    ReDim AR_laged(1 To AR_lag)
    For i = 1 To AR_lag
        AR_laged(i) = independent_variables(T - i + 1)
        AR_sum = AR_sum + AR_laged(i) * R_coef(i)
    Next i
Else
    ReDim AR_laged(1 To 1)
    AR_laged(1) = 0
End If

'create MA terms
If MA_lag > 0 Then
    ReDim MA_laged(1 To MA_lag)
    For i = 1 To MA_lag
        MA_laged(i) = errors(T - i + 1)
    Next i
Else
    ReDim MA_laged(1 To 1)
    MA_laged(1) = 0
End If
'dovle
ReDim function_values(1 To T - AR_lag)


For i = 1 To nAhead
    
    function_values(i) = constant + AR_sum + MA_sum
    
    AR_sum = 0
    MA_sum = 0
    
    If AR_lag > 0 Then
        For j = 1 To AR_lag
            AR_laged(j) = TimeSeriesRange(i + AR_lag - j + 1)
            AR_sum = AR_sum + AR_laged(j) * R_coef(j)
        Next j
    End If
    
    If MA_lag > 0 Then
        For j = 1 To MA_lag
            If i - j + 1 < 1 Then
                Exit For
            End If
            MA_laged(j) = TimeSeriesRange(i - j + 1 + AR_lag) - function_values(i - j + 1)
            MA_sum = MA_sum + e_coef(j) * MA_laged(j)
        Next j
    End If
    
Next i



End Function

Function NumberOfDimensions(ByVal vArray As Variant) As Long
Dim errorcheck As Long
Dim dimnum As Long
On Error GoTo FinalDimension

For dimnum = 1 To 60000
    errorcheck = LBound(vArray, dimnum)
Next

FinalDimension:
    NumberOfDimensions = dimnum - 1

End Function


Function error(params As Variant, independent_variables() As Double, dependent_variables() As Double)
    Dim errors() As Double
    Dim fitted() As Variant
    Dim i As Long
    Dim n As Long
    
    n = UBound(dependent_variables)
    fitted = Application.Run("arma_fitted", params, independent_variables)
    
    ReDim errors(1 To n)
    
    For i = 1 To n
        errors(i) = dependent_variables(i) - fitted(i)
    Next i
    
    error = errors
    
End Function
