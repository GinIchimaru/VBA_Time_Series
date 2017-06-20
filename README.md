# Time Series
Unit root tests, ARIMAX, GARCH models in VBA for the time being.
Files with extension .vba contains the actual code. The repo was committed to git with [VbaGitBootStrap](https://github.com/brucemcpherson/VbaGit). For now it contains only:  
*   Augmented Dickey Fuller test
*   Phillips Perron test
*   KPSS test
*   ARMA(Lag<sub>ar</sub>,Lag<sub>ma</sub>) estimation function:
    *  uses *Levenberg–Marquardt* procedure to minimize non linear sum of squares (conditional sum of squares)
    *  lag coefficients can be constrained to any value (*read zero value*),
    *  uses numerical hessian to calculate standard errors so caution should be taken when dealing with inference.   
    
I will upload first version of GARCH as soon as I finish my master thesis. The general goal is to have a choice between likelihood estimators an non linear least squares estimator for both ARMA and GARCH. GARCH module should have more than one model preferably with the choice of, at least, two conditional distributions when dealing with ML estimation method. Depending on some life circumstances I may be more or less committed to this project so when will all this be done I cannot tell for certain.  
The actual comment is more of a reminder for me what I promised to myself to do than an info for you, you poor soul, just how did you get here.  
