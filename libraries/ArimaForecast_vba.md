# VBA Project: **VBA_Time_Series**
## VBA Module: **[ArimaForecast](/libraries/ArimaForecast.vba "source is here")**
### Type: StdModule  

This procedure list for repo (VBA_Time_Series) was automatically created on 8/14/2017 10:32:36 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in ArimaForecast

---
VBA Procedure: **FORECAST_ARMA**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function FORECAST_ARMA(TimeSeriesRange As Variant, lags As Variant, params As Variant, Optional nAhead As Variant = 5)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TimeSeriesRange|Variant|False||
lags|Variant|False||
params|Variant|False||
nAhead|Variant|True| 5|


---
VBA Procedure: **NumberOfDimensions**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function NumberOfDimensions(ByVal vArray As Variant) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Variant|False||


---
VBA Procedure: **error**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function error(params As Variant, independent_variables() As Double, dependent_variables() As Double)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
params|Variant|False||
independent_variables|Variant|False||
dependent_variables|Variant|False||
