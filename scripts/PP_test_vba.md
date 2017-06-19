# VBA Project: **VBA_Time_Series**
## VBA Module: **[PP_test](/scripts/PP_test.vba "source is here")**
### Type: StdModule  

This procedure list for repo (VBA_Time_Series) was automatically created on 6/19/2017 9:54:35 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in PP_test

---
VBA Procedure: **PPtest**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function PPtest(TimeSeriesRange As Range, Optional lags As Variant = "short", Optional Intercept As Boolean = True, Optional Trend As Boolean = False)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TimeSeriesRange|Range|False||
lags|Variant|True| "short"|
Intercept|Boolean|True| True|
Trend|Boolean|True| False|


---
VBA Procedure: **LagTimeSeries**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function LagTimeSeries(data As Variant, Lag As Long)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
data|Variant|False||
Lag|Long|False||


---
VBA Procedure: **RangeToArray**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function RangeToArray(data As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
data|Variant|False||


---
VBA Procedure: **MacKinnon**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function MacKinnon(T As Long, Intercept As Boolean, Trend As Boolean) As Double()*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
T|Long|False||
Intercept|Boolean|False||
Trend|Boolean|False||
