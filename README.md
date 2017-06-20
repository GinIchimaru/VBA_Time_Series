# Time Series
Unit root tests, ARIMAX, GARCH models in VBA for the time being.
Files with extension .vba contains the actual code. The repo was commited to git with [VbaGitBootStrap](https://github.com/brucemcpherson/VbaGit). For now it contains only:  
*   Augmented Dickey Fuller test
*   Phillips Perron test
*   KPSS test
*   ARMA(Lag<sub>ar</sub>,Lag<sub>ma</sub>) estimation function:
    *  uses *Levenberg–Marquardt* procedure to minimize non linear sum of squares (conditional sum of squares)
    *  lag coefficients can be constrained to any value (*read zero value*),
    *  uses numerical hessian to calculate standard errors so caution should be taken when dealing with inference.
