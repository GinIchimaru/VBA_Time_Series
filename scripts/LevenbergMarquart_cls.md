# VBA Project: **VBA_Time_Series**
## VBA Module: **[LevenbergMarquart](/scripts/LevenbergMarquart.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (VBA_Time_Series) was automatically created on 6/20/2017 12:25:03 AM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in LevenbergMarquart

---
VBA Procedure: **lmdif**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub lmdif(fcn As String, m As Single, n As Single, x As Variant, fvec As Variant, ftol As Double, xtol As Double, gtol As Double, maxfev As Single, epsfcn As Double, diag As Variant, mode As Single, factor As Double, nprint As Single, info As Variant, nfev As Variant, fjac As Variant, ldfjac As Single, ipvt As Variant, qtf As Variant, wa1 As Variant, wa2 As Variant, wa3 As Variant, wa4 As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
fcn|String|False||
m|Single|False||
n|Single|False||
x|Variant|False||
fvec|Variant|False||
ftol|Double|False||
xtol|Double|False||
gtol|Double|False||
maxfev|Single|False||
epsfcn|Double|False||
diag|Variant|False||
mode|Single|False||
factor|Double|False||
nprint|Single|False||
info|Variant|False||
nfev|Variant|False||
fjac|Variant|False||
ldfjac|Single|False||
ipvt|Variant|False||
qtf|Variant|False||
wa1|Variant|False||
wa2|Variant|False||
wa3|Variant|False||
wa4|Variant|False||


---
VBA Procedure: **dsqrt**  
Type: **Function**  
Returns: **Double**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function dsqrt(x As Double) As Double*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
x|Double|False||


---
VBA Procedure: **dmax1**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function dmax1(a As Double, b As Double)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
a|Double|False||
b|Double|False||


---
VBA Procedure: **dmin1**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function dmin1(a As Double, b As Double)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
a|Double|False||
b|Double|False||


---
VBA Procedure: **min1single**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function min1single(a As Single, b As Single)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
a|Single|False||
b|Single|False||


---
VBA Procedure: **dabs**  
Type: **Function**  
Returns: **Double**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function dabs(x As Double) As Double*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
x|Double|False||


---
VBA Procedure: **testenorm**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub testenorm(x As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
x|Variant|False||


---
VBA Procedure: **qrfac**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub qrfac(m As Single, n As Single, ByRef a As Variant, lda As Single, pivot As Boolean, ByRef ipvt As Variant, lipvt As Single, ByRef rdiag As Variant, ByRef acnorm As Variant, ByRef wa As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
m|Single|False||
n|Single|False||
ByRef|Variant|False||
lda|Single|False||
pivot|Boolean|False||
ByRef|Variant|False||
lipvt|Single|False||
ByRef|Variant|False||
ByRef|Variant|False||
ByRef|Variant|False||


---
VBA Procedure: **mymod**  
Type: **Function**  
Returns: **Single**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function mymod(x As Single, y As Single) As Single*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
x|Single|False||
y|Single|False||


---
VBA Procedure: **myenorm**  
Type: **Function**  
Returns: **Double**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function myenorm(n As Single, x As Variant) As Double*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
n|Single|False||
x|Variant|False||


---
VBA Procedure: **lmpar**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub lmpar(n As Single, ByRef r As Variant, ldr As Single, ipvt As Variant, diag As Variant, qtb As Variant, delta As Double, ByRef par As Double, ByRef x As Variant, ByRef sdiag As Variant, ByRef wa1 As Variant, ByRef wa2 As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
n|Single|False||
ByRef|Variant|False||
ldr|Single|False||
ipvt|Variant|False||
diag|Variant|False||
qtb|Variant|False||
delta|Double|False||
ByRef|Double|False||
ByRef|Variant|False||
ByRef|Variant|False||
ByRef|Variant|False||
ByRef|Variant|False||


---
VBA Procedure: **qrsolv**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub qrsolv(n As Single, ByRef r As Variant, ldr As Single, ByRef ipvt As Variant, ByRef diag As Variant, ByRef qtb As Variant, ByRef x As Variant, ByRef sdiag As Variant, ByRef wa As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
n|Single|False||
ByRef|Variant|False||
ldr|Single|False||
ByRef|Variant|False||
ByRef|Variant|False||
ByRef|Variant|False||
ByRef|Variant|False||
ByRef|Variant|False||
ByRef|Variant|False||


---
VBA Procedure: **fdjac2**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub fdjac2(m As Single, n As Single, x As Variant, fvec As Variant, ByRef fjac As Variant, ldfjac As Single, iflag As Single, epsfcn As Double, wa As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
m|Single|False||
n|Single|False||
x|Variant|False||
fvec|Variant|False||
ByRef|Variant|False||
ldfjac|Single|False||
iflag|Single|False||
epsfcn|Double|False||
wa|Variant|False||


---
VBA Procedure: **LevenbergCostFunction**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub LevenbergCostFunction(m As Single, n As Single, x As Variant, ByRef fvec As Variant, ByRef iflag As Single)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
m|Single|False||
n|Single|False||
x|Variant|False||
ByRef|Variant|False||
ByRef|Single|False||


---
VBA Procedure: **CalibrateParameters**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function CalibrateParameters(FunctionName As String, xvec As Variant, yvec As Variant, params As Variant, Optional ftol As Double = 0.00000001, Optional xtol As Double = 0.00000001, Optional gtol As Double = 0.00000001, Optional maxfeval As Single = 400, Optional epsfcn As Double = 0.00000001) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
FunctionName|String|False||
xvec|Variant|False||
yvec|Variant|False||
params|Variant|False||
ftol|Double|True| 0.00000001|
xtol|Double|True| 0.00000001|
gtol|Double|True| 0.00000001|
maxfeval|Single|True| 400|
epsfcn|Double|True| 0.00000001|
