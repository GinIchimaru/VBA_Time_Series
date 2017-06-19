# VBA Project: **VBA_Time_Series**
## VBA Module: **[ADF_test](/scripts/ADF_test.vba "source is here")**
### Type: StdModule  

This procedure list for repo (VBA_Time_Series) was automatically created on 6/19/2017 9:40:38 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in ADF_test

---
VBA Procedure: **ADFtest**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function ADFtest(TimeSeriesRange As Range, Optional Lagg As Variant, Optional LagCriteria As Variant, Optional Intercept As Boolean = True, Optional Trend As Boolean = False)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TimeSeriesRange|Range|False||
Lagg|Variant|True||
LagCriteria|Variant|True||
Intercept|Boolean|True| True|
Trend|Boolean|True| False|


---
VBA Procedure: **ADFRegression**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function ADFRegression(TimeSeries As Range, Lag As Long, Optional Intercept As Boolean = False, Optional Trend As Boolean = False, Optional trim As Long = 0) As Double()*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
TimeSeries|Range|False||
Lag|Long|False||
Intercept|Boolean|True| False|
Trend|Boolean|True| False|
trim|Long|True| 0) As Double(|


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
VBA Procedure: **DifferenceTimeSeries**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function DifferenceTimeSeries(data As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
data|Variant|False||


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
