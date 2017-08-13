Option Explicit

'###############################################################################
'#######################         Arma estimation                ################
'###############################################################################

'''''''''''''''''''''''''''''         usage          '''''''''''''''''''''''''''
'

'Function ARMA_CSS() :
'   This is the main function, it gives estimated coefficients with corresponding
'   p values as an array output. The coefficients are estimated by minimizing
'   conditional sum of squares.

'       Parameters:
'           'TimeSeriesRange- range time series for which we want to estimate the coeficients
'           'lags - an array of sort {AR_lag, MA_lag} or range of lags of length 2
'           'initial_values - an array or range of (AR_lag+MA_lag+1) values
'              starting from constant followed by AR_lags initial values for AR
'              terms and MA_lags initial values for MA terms, where AR_lags and
'              MA_lags are integers denoting number of lagged terms of AR and
'              MA components. If omited, all take values 0.2!!!

'Function arma_fitted:
'   Private function that calculates FITTED values of range TimeSeriesRange for given
'   coefficients. This is the functions that enters as an input in function
'   calibrateParameters from Levenberg class in ARMA_CS function
'
'       Parameters:
'           params - which are actually initial values from ARMA_CSS function
'           TimeSeries - range pased down from ARMA_CSS function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''    supporting functions       ''''''''''''''''''''''''''''
'NumberOfDimensions:
'   Private function that checks number of dimensions of array
'
'Hessian:
'   Private function that calculates finite differences hessian for
'   a given sum of squares function.

'       Parameters:
'           FunctName - name of the function that we want to calculate the
'             hessian
'           Parameters - parameters that enter the function, that is independent
'             variables of function FunctName
'           TimeSeries_independent - time series of T length,
'             in the global context this is TimeSeriesRange without first AR_lag points
'           TimeSeries_dependent - time series of T-AR_lag length
'             in the global context this is TimeSeriesRange without first AR_lag points
'           deltaX is increment dx, the precission of derivatives depends
'             on this value.
'           *Note:Exept for deltaX all parameters are here just to be passed
'             down to SS function

'SS - returns sum of squares for a given function with dependent,
'   independent variables by summing residuals


'On the module level 4 variables have been declared in order to be
'shared between ARMA_CSS and arma_fitted since otherwise
'Levenberg.calibrateParameters would treat them as an optimization variable



'######################################################################################
'###############################         code      ####################################
'######################################################################################


Dim AR_lag                       As Integer
Dim MA_lag                       As Integer
Dim Fixed_Parameters()           As Variant
Dim missing_fixed_parameters     As Boolean          'checks if fixed values for parameters are provided



Function ARMA_CSS(TimeSeriesRange As Range, _
                  lags As Variant, _
                  Optional initial_values As Variant, _
                  Optional Fixed_values As Variant)

Dim params()                   As Double           'parameters for optimization-initial values
Dim i As Integer, j As Integer, jj As Integer, ii As Integer, num_of_fixed As Integer
Dim T                          As Integer          'time series lenght
Dim independent_variables()    As Double           'whole time series, used in arma_fitted for lagged terms (rhs of ARMA equation)
Dim dependent_variables()      As Double           'level time series without first AR_lag point, lhs of ARMA equation
Dim res()                      As Double           'optimized parameters
Dim residuals                  As Variant
Dim SumOfSquares               As Double
Dim variance                   As Double
Dim Hessian                    As Variant
Dim inverseHessian             As Variant
Dim coefs_errors()             As Variant
Dim CoefErrorMatrix()          As Double
Dim size                       As Long
Dim LogLik                     As Double
Dim AIC                        As Double
Dim BIC                        As Double
Dim AIC_aug                    As Double
Dim BIC_aug                    As Double
Dim p_values()                 As Variant
Dim statistics()               As Variant
Dim final_result()             As Variant




'-------------------------------------------------------------------------------------|
'--------------------------   checking input data           --------------------------|
'-------------------------------------------------------------------------------------|


'check the dimension of lag
'is lag a range?
If (TypeName(lags) = "Range") Then
'if yes then it must be eather one row or one column range
    If lags.Rows.Count > 2 And lags.Columns.Count > 2 Then
        ARMA_CSS = CVErr(xlErrValue)
    End If
Else
    If NumberOfDimensions(lags) > 1 Or UBound(lags) > 2 Then
       ARMA_CSS = CVErr(xlErrValue)
       Exit Function
    End If
End If

'Check inputs for being range
If (TypeName(TimeSeriesRange) <> "Range") Then
    ARMA_CSS = CVErr(xlErrValue)
Exit Function
End If

'if more than one column then exit
If TimeSeriesRange.Columns.Count > 1 Then
    ARMA_CSS = CVErr(xlErrNA)
    Exit Function
End If

'check initial values to be one dimension
If IsMissing(initial_values) = False Then
'On Error GoTo keepOn
    If TypeName(initial_values) = "Range" Then
        If NumberOfDimensions(initial_values) > 1 Then
            ARMA_CSS = CVErr(xlErrValue)
            Exit Function
        End If
    ElseIf TypeName(initial_values) = "Variant()" Then
        If UBound(initial_values) > (lags(1) + lags(2) + 1) Then
            ARMA_CSS = CVErr(xlErrValue)
            Exit Function
        End If
'keepOn:
    End If
End If

'check fixed values to be one dimension
missing_fixed_parameters = IsMissing(Fixed_values)

If IsMissing(Fixed_values) = False Then
'On Error GoTo keepOn2
    If TypeName(Fixed_values) = "Range" Then
        If NumberOfDimensions(Fixed_values) > 1 Then
            ARMA_CSS = CVErr(xlErrValue)
            Exit Function
'keepOn2:
        End If
    ElseIf TypeName(Fixed_values) = "Variant()" Then
        If UBound(Fixed_values) > (lags(1) + lags(2) + 1) Then
            ARMA_CSS = CVErr(xlErrValue)
            Exit Function
        End If
    End If

    ReDim Fixed_Parameters(1 To (lags(1) + lags(2) + 1))
    For i = 1 To (lags(1) + lags(2) + 1)
        Fixed_Parameters(i) = Fixed_values(i)
    Next i
End If


'|-------------------------------------------------------------------------------|
'|-------------------------    creating lagged series        --------------------|
'|-------------------------------------------------------------------------------|

'assign lags to public variables in order to pass them to arima function
'in order to avoid passing them as parameters
AR_lag = lags(1)
MA_lag = lags(2)

'creating initial values for coefficients if omited in function
'if missing initial values, assign values
ReDim params(1 To (AR_lag + MA_lag + 1))

If IsMissing(initial_values) Then
    For i = 1 To (AR_lag + MA_lag + 1)
        If missing_fixed_parameters = False Then
            If Fixed_Parameters(i) <> vbNullString Then
                params(i) = Fixed_Parameters(i)
            Else
                params(i) = 0.2
            End If
        Else
            params(i) = 0.2
        End If
    Next i
Else
    For i = 1 To (AR_lag + MA_lag + 1)
        If missing_fixed_parameters = False Then
            If Fixed_Parameters(i) <> "" Then
                params(i) = Fixed_Parameters(i)
            Else
                params(i) = initial_values(i)
            End If
        Else
            params(i) = initial_values(i)
        End If
    Next i
End If

'calculate the length of time series
T = TimeSeriesRange.Rows.Count

'create lagged series to be used in as dependent variable
'create lagged series to be base for calculating right side of arma equation,
'   *note that this is column vector, there was no need of creating
'    multiple column of laged series since we are looping them anyway
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
'-------------------------------------------------------------------------------------|
'--------------------------   calibrate parameters          --------------------------|
'-------------------------------------------------------------------------------------|
'create new instance of Levenberg object
Dim levObj As New LevenbergMarquart

'estimate the coefficients by minimizing:
'           sum((y-arma_fitted(params, independent_variables)^2)
res = levObj.CalibrateParameters("arma_fitted", independent_variables, dependent_variables, params)

'|------------------------------------------------------------------------------------|
'|-------------------------   calculating statistics     -----------------------------|
'|------------------------------------------------------------------------------------|


'take out the residuals from levObj class
residuals = levObj.fvec

'calculate the sum of squares
SumOfSquares = Application.WorksheetFunction.SumProduct(residuals, residuals)

'calculate unbiased variance
size = T - AR_lag
variance = SumOfSquares / size

'calculate hessian of SS function for optimized coefficient values
Hessian = Hessian_("SS", res, independent_variables, dependent_variables, 0.00001)

'in case of fixed parameters exclude zeros from hessian in order to take the inverse
'   call the new created hessian HessianCorrected
'   assign it to be new hessian
If missing_fixed_parameters = False Then
    
    num_of_fixed = 0
    
    For i = 1 To UBound(Hessian)
        If Fixed_Parameters(i) = "" Then num_of_fixed = 1 + num_of_fixed
    Next i
    
    Dim HessianCorrected()
    ReDim HessianCorrected(1 To num_of_fixed, 1 To num_of_fixed) As Variant
    
    ii = 0
    
    For i = 1 To UBound(Hessian)
    If Fixed_Parameters(i) = "" Then
        ii = ii + 1
        jj = 0
        For j = 1 To UBound(Hessian)
             If Fixed_Parameters(j) = "" Then
                jj = jj + 1
                HessianCorrected(ii, jj) = Hessian(i, j)
            End If
        Next j
    End If
    Next i
    ReDim Hessian(1 To num_of_fixed, 1 To num_of_fixed)
    Hessian = HessianCorrected 'hessian without zeros
End If

inverseHessian = Application.WorksheetFunction.MInverse(Hessian)

'calculate covariance matrix of coefficients errors
ReDim CoefErrorMatrix(1 To UBound(Hessian), 1 To UBound(Hessian))
For i = 1 To UBound(Hessian)
    For j = 1 To UBound(Hessian)
        If UBound(Hessian) = 1 Then
            CoefErrorMatrix(i, j) = inverseHessian(i) * variance * 2
        Else
            CoefErrorMatrix(i, j) = inverseHessian(i, j) * variance * 2  'took me a while, still cant figure out this one
        End If
    Next j
Next i

'calculating coefficients errors
ReDim coefs_errors(1 To UBound(params))
ii = 0
For i = 1 To UBound(params)
If missing_fixed_parameters = False Then
    If Fixed_Parameters(i) <> "" Then
        ii = ii + 1
        coefs_errors(i) = CVErr(xlErrValue) 'put error on place of fixed parameter
    Else
        On Error GoTo 0 'in case of negative values under the sqr
        On Error GoTo errorHandler
        coefs_errors(i) = Sqr(CoefErrorMatrix(i - ii, i - ii))
    End If
Else
    coefs_errors(i) = Sqr(CoefErrorMatrix(i, i))
End If
Next i


'calculating other statistics
With Application
    AIC = size * .Ln(variance) + 2 * (2 + AR_lag + MA_lag)
    AIC_aug = size * (1 + .Ln(2 * .Pi())) + AIC
    LogLik = -(AIC_aug - 2 * (AR_lag + MA_lag + 2)) / 2
    BIC = size * .Ln(variance) + .Ln(size) * (AR_lag + MA_lag + 2)
    BIC_aug = size * (1 + .Ln(2 * .Pi())) + BIC
End With

'calculating p-values and statistics
ReDim p_values(1 To AR_lag + MA_lag + 1)
ReDim statistics(1 To AR_lag + MA_lag + 1)
With Application.WorksheetFunction
    For i = 1 To UBound(params)
        If missing_fixed_parameters = False Then
            If Fixed_Parameters(i) <> "" Then
                statistics(i) = CVErr(xlErrValue)
                p_values(i) = CVErr(xlErrValue)
            Else
                statistics(i) = res(i) / coefs_errors(i)
                p_values(i) = (1 - .T_Dist(Abs(statistics(i)), size - AR_lag - MA_lag - 1, True)) * 2
            End If
        Else
            statistics(i) = res(i) / coefs_errors(i)
            p_values(i) = (1 - .T_Dist(Abs(statistics(i)), size - AR_lag - MA_lag - 1, True)) * 2
        End If
    Next i
End With


'put NA value in places that are empty,
'   for this purpose dummy variable is used in order
'   to secure constant length of 5 places for variance,
'   SumOfSquares, LogLik, AIC, BIC
Dim dummy As Integer
If AR_lag + MA_lag + 1 < 5 Then
    dummy = 5 - (AR_lag + MA_lag + 1)
Else
    dummy = 0
End If


'creating output array
ReDim final_result(1 To 5, 1 To AR_lag + MA_lag + 1 + dummy)

For i = 1 To AR_lag + MA_lag + 1 + dummy
    If i <= AR_lag + MA_lag + 1 Then
        final_result(1, i) = res(i)
        final_result(2, i) = coefs_errors(i)
        final_result(3, i) = statistics(i)
        final_result(4, i) = p_values(i)
    Else
        final_result(1, i) = CVErr(xlErrNA)
        final_result(2, i) = CVErr(xlErrNA)
        final_result(3, i) = CVErr(xlErrNA)
        final_result(4, i) = CVErr(xlErrNA)
        final_result(5, i) = CVErr(xlErrNA)
    End If
Next i

final_result(5, 1) = variance
final_result(5, 2) = SumOfSquares
final_result(5, 3) = LogLik
final_result(5, 4) = AIC
final_result(5, 5) = BIC

Erase Fixed_Parameters 'IMPORTANT


ARMA_CSS = final_result

Exit Function

errorHandler:

coefs_errors(i) = CVErr(xlErrValue)
Resume Next

End Function


Private Function arma_fitted(params As Variant, TimeSeries() As Double)
Dim T                   As Long
Dim i                   As Long
Dim j                   As Long
Dim function_values     As Variant
Dim R_coef()            As Double
Dim e_coef()            As Double
Dim constant            As Double
Dim AR_laged()          As Double
Dim MA_laged()          As Double
Dim AR_sum              As Double
Dim MA_sum              As Double
'read in time series
T = UBound(TimeSeries)

'check if there are fixed values and assign them to params if there are any
If missing_fixed_parameters = False Then
    For i = 1 To UBound(params)
        If Fixed_Parameters(i) <> "" Then
            params(i) = Fixed_Parameters(i)
         End If
    Next i
End If
'create AR coefs
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
        AR_laged(i) = TimeSeries(AR_lag - i + 1)
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
        MA_laged(i) = 0
    Next i
Else
    ReDim MA_laged(1 To 1)
    MA_laged(1) = 0
End If

ReDim function_values(1 To T - AR_lag)

constant = params(1)


For i = 1 To (T - AR_lag)
    
    function_values(i) = constant + AR_sum + MA_sum
    
    AR_sum = 0
    MA_sum = 0
    
    If AR_lag > 0 Then
        For j = 1 To AR_lag
            AR_laged(j) = TimeSeries(i + AR_lag - j + 1)
            AR_sum = AR_sum + AR_laged(j) * R_coef(j)
        Next j
    End If
    
    If MA_lag > 0 Then
        For j = 1 To MA_lag
            If i - j + 1 < 1 Then
                Exit For
            End If
            MA_laged(j) = TimeSeries(i - j + 1 + AR_lag) - function_values(i - j + 1)
            MA_sum = MA_sum + e_coef(j) * MA_laged(j)
        Next j
    End If
    
Next i

arma_fitted = function_values

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'supporting functions                                              '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NumberOfDimensions(ByVal vArray As Variant) As Long
Dim errorcheck As Long
Dim dimnum As Long
On Error GoTo FinalDimension

For dimnum = 1 To 60000
    errorcheck = LBound(vArray, dimnum)
Next

FinalDimension:
    NumberOfDimensions = dimnum - 1

End Function


Private Function Hessian_(FunctName, parameters As Variant, TimeSeries_independent() As Double, TimeSeries_dependent() As Double, deltaX As Double)
'Second-order derivatives based on function calls only (Dennis and Schnabel 1983, p. 80, 104):
'for dense Hessian, n+n2/2 additional function calls are needed:
Dim i, j               As Integer
Dim NumOfDimensions    As Long
Dim x()                As Double
Dim ddy                As Variant
Dim f                  As Variant
Dim f2                 As Variant
Dim f3                 As Variant
Dim f4                 As Variant
Dim x2()               As Double
Dim x3()               As Double
Dim x4()               As Double
Dim hes()              As Double
Dim mis()               As Boolean

NumOfDimensions = UBound(parameters)

ReDim x(1 To NumOfDimensions)
ReDim x2(1 To NumOfDimensions)
ReDim x3(1 To NumOfDimensions)
ReDim x4(1 To NumOfDimensions)
ReDim hes(1 To NumOfDimensions, 1 To NumOfDimensions)

x = parameters

f = Application.Run(FunctName, parameters, TimeSeries_independent, TimeSeries_dependent)

For i = 1 To NumOfDimensions
    For j = 1 To NumOfDimensions
    
    x2 = parameters
    x3 = parameters
    x4 = parameters

    x2(i) = x2(i) + deltaX
    x2(j) = x2(j) + deltaX
    x3(i) = x3(i) + deltaX
    x4(j) = x4(j) + deltaX
    
    f2 = Application.Run(FunctName, x2, TimeSeries_independent, TimeSeries_dependent)
    f3 = Application.Run(FunctName, x3, TimeSeries_independent, TimeSeries_dependent)
    f4 = Application.Run(FunctName, x4, TimeSeries_independent, TimeSeries_dependent)
    
    ddy = f2 - f3 - f4 + f
    hes(i, j) = ddy / (deltaX * deltaX)
    
    Next j
Next i

Hessian_ = hes

End Function

Private Function SS(params As Variant, independent_variables() As Double, dependent_variables() As Double)
    Dim errors() As Double
    Dim fitted() As Variant
    Dim i As Long
    Dim n As Long
    
    n = UBound(dependent_variables)
    fitted = arma_fitted(params, independent_variables)
    
    ReDim errors(1 To n)
    
    For i = 1 To n
        errors(i) = dependent_variables(i) - fitted(i)
    Next i
    
    SS = Application.WorksheetFunction.SumProduct(errors, errors)
    
End Function
