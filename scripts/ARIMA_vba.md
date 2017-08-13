# VBA Project: **VBA_Time_Series**
## VBA Module: **[ARIMA](/scripts/ARIMA.vba "source is here")**
### Type: StdModule  

This procedure list for repo (VBA_Time_Series) was automatically created on 8/14/2017 1:06:14 AM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in ARIMA

---
VBA Procedure: **ARMA_CSS**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function ARMA_CSS(TimeSeriesRange As Range, lags As Variant, Optional initial_values As Variant, Optional Fixed_values As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TimeSeriesRange|Range|False||
lags|Variant|False||
initial_values|Variant|True||
Fixed_values|Variant|True||


---
VBA Procedure: **arma_fitted**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function arma_fitted(params As Variant, TimeSeries() As Double)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
params|Variant|False||
TimeSeries|Variant|False||


---
VBA Procedure: **NumberOfDimensions**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function NumberOfDimensions(ByVal vArray As Variant) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Variant|False||


---
VBA Procedure: **Hessian_**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function Hessian_(FunctName, parameters As Variant, TimeSeries_independent() As Double, TimeSeries_dependent() As Double, deltaX As Double)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
FunctName|Variant|False||
parameters|Variant|False||
TimeSeries_independent|Variant|False||
TimeSeries_dependent|Variant|False||
deltaX|Double|False||


---
VBA Procedure: **SS**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function SS(params As Variant, independent_variables() As Double, dependent_variables() As Double)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
params|Variant|False||
independent_variables|Variant|False||
dependent_variables|Variant|False||
