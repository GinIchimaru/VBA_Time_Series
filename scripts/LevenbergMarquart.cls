'Minpack Copyright Notice (1999) University of Chicago.  All rights reserved
'
'Redistribution and use in source and binary forms, with or
'without modification, are permitted provided that the
'following conditions are met:
'
'1. Redistributions of source code must retain the above
'copyright notice, this list of conditions and the following
'disclaimer.
'
'2. Redistributions in binary form must reproduce the above
'copyright notice, this list of conditions and the following
'disclaimer in the documentation and/or other materials
'provided with the distribution.
'
'3. The end-user documentation included with the
'redistribution, if any, must include the following
'acknowledgment:
'
'   "This product includes software developed by the
'   University of Chicago, as Operator of Argonne National
'   Laboratory.
'
'Alternately, this acknowledgment may appear in the software
'itself, if and wherever such third-party acknowledgments
'normally appear.
'
'4. WARRANTY DISCLAIMER. THE SOFTWARE IS SUPPLIED "AS IS"
'WITHOUT WARRANTY OF ANY KIND. THE COPYRIGHT HOLDER, THE
'UNITED STATES, THE UNITED STATES DEPARTMENT OF ENERGY, AND
'THEIR EMPLOYEES: (1) DISCLAIM ANY WARRANTIES, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO ANY IMPLIED WARRANTIES
'OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, TITLE
'OR NON-INFRINGEMENT, (2) DO NOT ASSUME ANY LEGAL LIABILITY
'OR RESPONSIBILITY FOR THE ACCURACY, COMPLETENESS, OR
'USEFULNESS OF THE SOFTWARE, (3) DO NOT REPRESENT THAT USE OF
'THE SOFTWARE WOULD NOT INFRINGE PRIVATELY OWNED RIGHTS, (4)
'DO NOT WARRANT THAT THE SOFTWARE WILL FUNCTION
'UNINTERRUPTED, THAT IT IS ERROR-FREE OR THAT ANY ERRORS WILL
'BE CORRECTED.
'
'5. LIMITATION OF LIABILITY. IN NO EVENT WILL THE COPYRIGHT
'HOLDER, THE UNITED STATES, THE UNITED STATES DEPARTMENT OF
'ENERGY, OR THEIR EMPLOYEES: BE LIABLE FOR ANY INDIRECT,
'INCIDENTAL, CONSEQUENTIAL, SPECIAL OR PUNITIVE DAMAGES OF
'ANY KIND OR NATURE, INCLUDING BUT NOT LIMITED TO LOSS OF
'PROFITS OR LOSS OF DATA, FOR ANY REASON WHATSOEVER, WHETHER
'SUCH LIABILITY IS ASSERTED ON THE BASIS OF CONTRACT, TORT
'(INCLUDING NEGLIGENCE OR STRICT LIABILITY), OR OTHERWISE,
'EVEN IF ANY OF SAID PARTIES HAS BEEN WARNED OF THE
'POSSIBILITY OF SUCH LOSS OR DAMAGES.

'|--------------------------------------------------------------------------|'
'|--------------------------------------------------------------------------|'
'             Documentation for MINPACK subroutine LMDIF
'
'                        Double precision version
'
'                      Argonne National Laboratory
'
'         Burton S. Garbow, Kenneth E. Hillstrom, Jorge J. More
'
'                               March 1980
'
'
' 1. Purpose.
'
'       The purpose of LMDIF is to minimize the sum of the squares of M
'       nonlinear functions in N variables by a modification of the
'       Levenberg-Marquardt algorithm.  The user must provide a subrou-
'       tine which calculates the functions.  The Jacobian is then cal-
'       culated by a forward-difference approximation.
'
'
' 2. Subroutine and type statements.
'
'       SUBROUTINE LMDIF(FCN,M,N,X,FVEC,FTOL,XTOL,GTOL,MAXFEV,EPSFCN,
'      *                 DIAG,MODE,FACTOR,NPRINT,INFO,NFEV,FJAC,LDFJAC,
'      *                 IPVT,QTF,WA1,WA2,WA3,WA4)
'       INTEGER M,N,MAXFEV,MODE,NPRINT,INFO,NFEV,LDFJAC
'       INTEGER IPVT(N)
'       DOUBLE PRECISION FTOL,XTOL,GTOL,EPSFCN,FACTOR
'       DOUBLE PRECISION X(N),FVEC(M),DIAG(N),FJAC(LDFJAC,N),QTF(N),
'      *                 WA1(N),WA2(N),WA3(N),WA4(M)
'       EXTERNAL fcn
'
'
' 3. Parameters.
'
'       Parameters designated as input parameters must be specified on
'       entry to LMDIF and are not changed on exit, while parameters
'       designated as output parameters need not be specified on entry
'       and are set to appropriate values on exit from LMDIF.
'
'       FCN is the name of the user-supplied subroutine which calculates
'         the functions.  FCN must be declared in an EXTERNAL statement
'         in the user calling program, and should be written as follows.
'
'         SUBROUTINE fcn(m, n, x, fvec, iflag)
'         INTEGER M,N,IFLAG
'         DOUBLE PRECISION X(N),FVEC(M)
'         ----------
'         CALCULATE THE FUNCTIONS AT X AND
'         RETURN THIS VECTOR IN FVEC.
'         ----------
'         Return
'         End
'
'
'         The value of IFLAG should not be changed by FCN unless the
'         user wants to terminate execution of LMDIF.  In this case set
'         IFLAG to a negative integer.
'
'       M is a positive integer input variable set to the number of
'         functions.
'
'       N is a positive integer input variable set to the number of
'         variables.  N must not exceed M.
'
'       X is an array of length N.  On input X must contain an initial
'         estimate of the solution vector.  On output X contains the
'         final estimate of the solution vector.
'
'       FVEC is an output array of length M which contains the functions
'         evaluated at the output X.
'
'       FTOL is a nonnegative input variable.  Termination occurs when
'         both the actual and predicted relative reductions in the sum
'         of squares are at most FTOL.  Therefore, FTOL measures the
'         relative error desired in the sum of squares.  Section 4 con-
'         tains more details about FTOL.
'
'       XTOL is a nonnegative input variable.  Termination occurs when
'         the relative error between two consecutive iterates is at most
'         XTOL.  Therefore, XTOL measures the relative error desired in
'         the approximate solution.  Section 4 contains more details
'         about XTOL.
'
'       GTOL is a nonnegative input variable.  Termination occurs when
'         the cosine of the angle between FVEC and any column of the
'         Jacobian is at most GTOL in absolute value.  Therefore, GTOL
'         measures the orthogonality desired between the function vector
'         and the columns of the Jacobian.  Section 4 contains more
'         details about GTOL.
'
'       MAXFEV is a positive integer input variable.  Termination occurs
'         when the number of calls to FCN is at least MAXFEV by the end
'         of an iteration.
'
'       EPSFCN is an input variable used in determining a suitable step
'         for the forward-difference approximation.  This approximation
'         assumes that the relative errors in the functions are of the
'         order of EPSFCN.  If EPSFCN is less than the machine preci-
'         sion, it is assumed that the relative errors in the functions
'         are of the order of the machine precision.
'
'       DIAG is an array of length N.  If MODE = 1 (see below), DIAG is
'         internally set.  If MODE = 2, DIAG must contain positive
'         entries that serve as multiplicative scale factors for the
'         variables.
'
'       MODE is an integer input variable.  If MODE = 1, the variables
'         will be scaled internally.  If MODE = 2, the scaling is
'
'
'                                                                 Page 3
'
'         specified by the input DIAG.  Other values of MODE are equiva-
'         lent to MODE = 1.
'
'       FACTOR is a positive input variable used in determining the ini-
'         tial step bound.  This bound is set to the product of FACTOR
'         and the Euclidean norm of DIAG*X if nonzero, or else to FACTOR
'         itself.  In most cases FACTOR should lie in the interval
'         (.1,100.).  100. is a generally recommended value.
'
'       NPRINT is an integer input variable that enables controlled
'         printing of iterates if it is positive.  In this case, FCN is
'         called with IFLAG = 0 at the beginning of the first iteration
'         and every NPRINT iterations thereafter and immediately prior
'         to return, with X and FVEC available for printing.  If NPRINT
'         is not positive, no special calls of FCN with IFLAG = 0 are
'         made.
'
'       INFO is an integer output variable.  If the user has terminated
'         execution, INFO is set to the (negative) value of IFLAG.  See
'         description of FCN.  Otherwise, INFO is set as follows.
'
'         INFO = 0  Improper input parameters.
'
'         INFO = 1  Both actual and predicted relative reductions in the
'                   sum of squares are at most FTOL.
'
'         INFO = 2  Relative error between two consecutive iterates is
'                   at most XTOL.
'
'         INFO = 3  Conditions for INFO = 1 and INFO = 2 both hold.
'
'         INFO = 4  The cosine of the angle between FVEC and any column
'                   of the Jacobian is at most GTOL in absolute value.
'
'         INFO = 5  Number of calls to FCN has reached or exceeded
'                   MAXFEV.
'
'         INFO = 6  FTOL is too small.  No further reduction in the sum
'                   of squares is possible.
'
'         INFO = 7  XTOL is too small.  No further improvement in the
'                   approximate solution X is possible.
'
'         INFO = 8  GTOL is too small.  FVEC is orthogonal to the
'                   columns of the Jacobian to machine precision.
'
'         Sections 4 and 5 contain more details about INFO.
'
'       NFEV is an integer output variable set to the number of calls to
'         FCN.
'
'       FJAC is an output M by N array.  The upper N by N submatrix of
'         FJAC contains an upper triangular matrix R with diagonal ele-
'         ments of nonincreasing magnitude such that
'
'
'                                                                 Page 4
'
'                T     T           T
'               P *(JAC *JAC)*P = R *R,
'
'         where P is a permutation matrix and JAC is the final calcu-
'         lated j
'|-------------------------------------------------------------------------|
'|-------------------------------------------------------------------------|



Option Explicit
Private FunctionName_ As String
Public fvec As Variant
Private IndependentVariables_ As Variant
Private ObservedVariables_ As Variant
Private iflag_ As Variant
Public jacobian As Variant

Private Sub lmdif(fcn As String, m As Single, n As Single, x As Variant, fvec As Variant, _
      ftol As Double, xtol As Double, gtol As Double, maxfev As Single, epsfcn As Double, _
      diag As Variant, mode As Single, factor As Double, nprint As Single, _
      info As Variant, nfev As Variant, fjac As Variant, ldfjac As Single, _
      ipvt As Variant, qtf As Variant, wa1 As Variant, wa2 As Variant, _
      wa3 As Variant, wa4 As Variant)
Dim i As Single
Dim iflag As Single
Dim ij As Single
Dim jj As Single
Dim iter As Single
Dim j As Single
Dim l As Single
Dim actred As Double
Dim delta As Double
delta = 0
Dim dirder As Double
Dim fnorm As Double
Dim fnorm1 As Double
Dim gnorm As Double
Dim par As Double
Dim pnorm As Double
Dim prered As Double
Dim ratio As Double
Dim sum As Double
Dim temp As Double
Dim temp1 As Double
Dim temp2 As Double
Dim temp3 As Double
Dim xnorm As Double
xnorm = 0
Dim one As Double
Dim p1 As Double
Dim p5 As Double
Dim p25 As Double
Dim p75 As Double
Dim p0001 As Double
Dim zero As Double
one = 1
p1 = 0.1
p5 = 0.5
p25 = 0.25
p75 = 0.75
p0001 = 0.0001
zero = 0
Dim epsmch As Double
'     epsmch is the machine precision.
epsmch = 1E-16

      info = 0
      iflag = 0
      nfev = 0

'     check the input parameters for errors.
      If ((n <= 0) Or (m < n) Or (ldfjac < m) _
         Or (ftol < zero) Or (xtol < zero) Or (gtol < zero) _
         Or (maxfev <= 0) Or (factor <= zero)) Then GoTo L300
      If (mode <> 2) Then GoTo L20
      For j = 1 To n
         If (diag(j) <= zero) Then GoTo L300
      Next j
      
L20:
'     evaluate the function at the starting point
'     and calculate its norm.
      iflag = 1
      LevenbergCostFunction m, n, x, fvec, iflag
      nfev = 1
      If (iflag < 0) Then GoTo L300
      fnorm = myenorm(m, fvec)

'     initialize levenberg-marquardt parameter and iteration counter.

      par = zero
      iter = 1

'     beginning of the outer loop.
'
L30:

'        calculate the jacobian matrix.

         iflag = 2
         fdjac2 m, n, x, fvec, fjac, ldfjac, iflag, epsfcn, wa4
         nfev = nfev + n
         If (iflag < 0) Then GoTo L300

'        if requested, call fcn to enable printing of iterates.

         If (nprint <= 0) Then GoTo L40
         iflag = 0
         If (mymod(iter - 1, nprint) = 0) Then LevenbergCostFunction m, n, x, fvec, iflag
         If (iflag < 0) Then GoTo L300
L40:
'
'        compute the qr factorization of the jacobian.
'
         qrfac m, n, fjac, ldfjac, True, ipvt, n, wa1, wa2, wa3
'
'        on the first iteration and if mode is 1, scale according
'        to the norms of the columns of the initial jacobian.
'
         If (iter <> 1) Then GoTo L80
         If (mode = 2) Then GoTo L60
         For j = 1 To n
            diag(j) = wa2(j)
            If (wa2(j) = zero) Then diag(j) = one
         Next j
L60:
'
'        on the first iteration, calculate the norm of the scaled x
'        and initialize the step bound delta.
'
         For j = 1 To n
            wa3(j) = diag(j) * x(j)
         Next j
         xnorm = myenorm(n, wa3)
         delta = factor * xnorm
         If (delta = zero) Then delta = factor

L80:
'
'        form (q transpose)*fvec and store the first n components in
'        qtf.
'
         For i = 1 To m
            wa4(i) = fvec(i)
         Next i
         For j = 1 To n
            If (fjac(j, j) = zero) Then GoTo L120
            sum = zero
            For i = j To m
               sum = sum + fjac(i, j) * wa4(i)
            Next i
            temp = -sum / fjac(j, j)
            For i = j To m
               wa4(i) = wa4(i) + fjac(i, j) * temp
            Next i
L120:
            fjac(j, j) = wa1(j)
            qtf(j) = wa4(j)
         Next j
'
'        compute the norm of the scaled gradient.
'
         gnorm = zero
         If (fnorm = zero) Then GoTo L170
         For j = 1 To n
            l = ipvt(j)
            If (wa2(l) = zero) Then GoTo L150
            sum = zero
            For i = 1 To j
               sum = sum + fjac(i, j) * (qtf(i) / fnorm)
            Next i
            gnorm = dmax1(gnorm, dabs(sum / wa2(l)))
L150:
         Next j
L170:
'
'        test for convergence of the gradient norm.
'
         If (gnorm <= gtol) Then info = 4
         If (info <> 0) Then GoTo L300
'
'        rescale if necessary.
'
         If (mode = 2) Then GoTo L190
         For j = 1 To n
            diag(j) = dmax1(CDbl(diag(j)), CDbl(wa2(j)))
         Next j
L190:
'
'        beginning of the inner loop.
'
L200:
'
'           determine the levenberg-marquardt parameter.
'
            lmpar n, fjac, ldfjac, ipvt, diag, qtf, delta, par, wa1, wa2, wa3, wa4
'
'           store the direction p and x + p. calculate the norm of p.
'
            For j = 1 To n
               wa1(j) = -wa1(j)
               wa2(j) = x(j) + wa1(j)
               wa3(j) = diag(j) * wa1(j)
            Next j
            pnorm = myenorm(n, wa3)
'
'           on the first iteration, adjust the initial step bound.
'
            If (iter = 1) Then delta = dmin1(delta, pnorm)
'
'           evaluate the function at x + p and calculate its norm.
'
            iflag = 1
            LevenbergCostFunction m, n, wa2, wa4, iflag
            nfev = nfev + 1
            If (iflag < 0) Then GoTo L300
            fnorm1 = myenorm(m, wa4)
'
'           compute the scaled actual reduction.
'
            actred = -one
            If (p1 * fnorm1 < fnorm) Then actred = one - (fnorm1 / fnorm) * (fnorm1 / fnorm)
'
'           compute the scaled predicted reduction and
'           the scaled directional derivative.
'
            For j = 1 To n
               wa3(j) = zero
               l = ipvt(j)
               temp = wa1(l)
               For i = 1 To j
                  wa3(i) = wa3(i) + fjac(i, j) * temp
               Next i
            Next j
            temp1 = myenorm(n, wa3) / fnorm
            temp2 = (dsqrt(par) * pnorm) / fnorm
            prered = temp1 * temp1 + (temp2 * temp2) / p5
            dirder = dirder - (temp1 * temp1 + temp2 * temp2)
'
'           compute the ratio of the actual to the predicted
'           reduction.
'
            ratio = zero
            If (prered <> zero) Then ratio = actred / prered
'
'           update the step bound.
'
            If (ratio > p25) Then GoTo L240
               If (actred >= zero) Then temp = p5
               If (actred < zero) Then _
                 temp = p5 * dirder / (dirder + p5 * actred)
               If ((p1 * fnorm1 >= fnorm) Or (temp < p1)) Then temp = p1
               delta = temp * dmin1(delta, pnorm / p1)
               par = par / temp
               GoTo L260
L240:
               If ((par <> zero) And (ratio < p75)) Then GoTo L250
               delta = pnorm / p5
               par = p5 * par
L250:
L260:
'
'           test for successful iteration.
'
            If (ratio < p0001) Then GoTo L290
'
'           successful iteration. update x, fvec, and their norms.
'
            For j = 1 To n
               x(j) = wa2(j)
               wa2(j) = diag(j) * x(j)
            Next j
            For i = 1 To m
               fvec(i) = wa4(i)
            Next i
            xnorm = myenorm(n, wa2)
            fnorm = fnorm1
            iter = iter + 1
L290:
'
'           tests for convergence.
'
            If ((dabs(actred) <= ftol) And (prered <= ftol) _
               And (p5 * ratio <= one)) Then info = 1
            If (delta <= xtol * xnorm) Then info = 2
            If ((dabs(actred) <= ftol) And (prered <= ftol) _
               And (p5 * ratio <= one) And (info = 2)) Then info = 3
            If (info <> 0) Then GoTo L300
'
'           tests for termination and stringent tolerances.
'
            If (nfev >= maxfev) Then info = 5
            If ((dabs(actred) <= epsmch) And (prered <= epsmch) _
               And (p5 * ratio <= one)) Then info = 6
            If (delta <= epsmch * xnorm) Then info = 7
            If (gnorm <= epsmch) Then info = 8
            If (info <> 0) Then GoTo L300
'
'           end of the inner loop. repeat if iteration unsuccessful.
'
            If (ratio < p0001) Then GoTo L200
'
'        end of the outer loop.
'
         GoTo L30
L300:
'
'     termination, either normal or user imposed.
'
      If (iflag < 0) Then info = iflag
      iflag = 0
      If (nprint > 0) Then LevenbergCostFunction m, n, x, fvec, iflag
      'Return
'
'     last card of subroutine lmdif.
'
      
End Sub

Private Function dsqrt(x As Double) As Double
dsqrt = x ^ 0.5
End Function
Private Function dmax1(a As Double, b As Double)
If a >= b Then
  dmax1 = a
Else
  dmax1 = b
End If
End Function
Private Function dmin1(a As Double, b As Double)
If a <= b Then
  dmin1 = a
Else
  dmin1 = b
End If
End Function

Private Function min1single(a As Single, b As Single)
If a <= b Then
  min1single = a
Else
  min1single = b
End If
End Function

Private Function dabs(x As Double) As Double
dabs = Abs(x)
End Function

Private Sub testenorm(x As Variant)

End Sub
 Private Sub qrfac(m As Single, n As Single, ByRef a As Variant, lda As Single, _
   pivot As Boolean, ByRef ipvt As Variant, lipvt As Single, ByRef rdiag As Variant, _
    ByRef acnorm As Variant, ByRef wa As Variant)
    Dim i As Single
    Dim j As Single
    Dim k As Single
    Dim jp1 As Single
    Dim kmax As Single
    Dim minmn As Single
    Dim ajnorm As Double
    Dim epsmch As Double
    Dim one As Double
    Dim p05 As Double
    Dim sum As Double
    Dim temp As Double
    Dim zero As Double
    Dim dpmpar As Double
    Dim enorm As Double
    one = 1
    p05 = 0.05
    zero = 0
    
    ReDim tmpaarr(1 To m) As Variant
    Dim tmpj As Single
    
'
'     epsmch is the machine precision.
'
      epsmch = 1.2E-16
'
'     compute the initial column norms and initialize several arrays.
'
      For j = 1 To n
         For tmpj = 1 To m
           tmpaarr(tmpj) = a(tmpj, j)
         Next tmpj
         'acnorm(j) = enorm(m, a(1, j))
         
         'a3 = enorm(m, a2)
         'testenorm tmpaarr
         acnorm(j) = myenorm(m, tmpaarr)
         rdiag(j) = acnorm(j)
         wa(j) = rdiag(j)
         If (pivot) Then ipvt(j) = j
      Next j
'
'     reduce a to r with householder transformations.
'
      minmn = min1single(m, n)
      For j = 1 To minmn
         If (Not pivot) Then GoTo L40
'
'        bring the column of largest norm into the pivot position.
'
         kmax = j
         For k = j To n
            If (rdiag(k) > rdiag(kmax)) Then kmax = k
         Next k
         If (kmax = j) Then GoTo L40
         For i = 1 To m
            temp = a(i, j)
            a(i, j) = a(i, kmax)
            a(i, kmax) = temp
         Next i
         rdiag(kmax) = rdiag(j)
         wa(kmax) = wa(j)
         k = ipvt(j)
         ipvt(j) = ipvt(kmax)
         ipvt(kmax) = k
L40:
'
'        compute the householder transformation to reduce the
'        j-th column of a to a multiple of the j-th unit vector.
'
         ReDim tmpaarr(1 To m - j + 1) As Variant
         For tmpj = 1 To m - j + 1
           tmpaarr(tmpj) = a(tmpj + j - 1, j)
         Next tmpj
         'ajnorm = myenorm(m - j + 1, a(j, j))
         ajnorm = myenorm(m - j + 1, tmpaarr)
         If (ajnorm = zero) Then GoTo L100
         If (a(j, j) < zero) Then ajnorm = -ajnorm
         For i = j To m
            a(i, j) = a(i, j) / ajnorm
         Next i
         a(j, j) = a(j, j) + one
'
'        apply the transformation to the remaining columns
'        and update the norms.
'
         jp1 = j + 1
         If (n < jp1) Then GoTo L100
         For k = jp1 To n
            sum = zero
            For i = j To m
               sum = sum + a(i, j) * a(i, k)
            Next i
            temp = sum / a(j, j)
            For i = j To m
               a(i, k) = a(i, k) - temp * a(i, j)
            Next i
            If ((Not pivot) Or (rdiag(k) = zero)) Then GoTo L80
            temp = a(j, k) / rdiag(k)
            rdiag(k) = rdiag(k) * dsqrt(dmax1(zero, one - temp * temp))
            If (p05 * (rdiag(k) / wa(k)) ^ 2 > epsmch) Then GoTo L80
            rdiag(k) = myenorm(m - j, a(jp1, k))
            wa(k) = rdiag(k)
L80:
         Next k
L100:
         rdiag(j) = -ajnorm
      Next j
      'Return
'
'     last card of subroutine qrfac.
'
 End Sub
Private Function mymod(x As Single, y As Single) As Single
mymod = 1
End Function

Private Function myenorm(n As Single, x As Variant) As Double
      Dim i As Single
      Dim agiant As Double
      Dim floatn As Double
      Dim one As Double
      Dim rdwarf As Double
      Dim rgiant As Double
      Dim s1 As Double
      Dim s2 As Double
      Dim s3 As Double
      Dim xabs As Double
      Dim x1max As Double
      Dim x3max As Double
      Dim zero As Double
      
      one = 1
      zero = 0
      rdwarf = 3.834E-20
      rgiant = 1.304E+19
      s1 = zero
      s2 = zero
      s3 = zero
      x1max = zero
      x3max = zero
      floatn = n
      agiant = rgiant / floatn
      For i = 1 To n
         xabs = dabs(CDbl(x(i)))
         If ((xabs > rdwarf) And (xabs < agiant)) Then GoTo L70
            If (xabs <= rdwarf) Then GoTo L30
'
'              sum for large components.
'
               If (xabs <= x1max) Then GoTo L10
                  s1 = one + s1 * (x1max / xabs) ^ 2
                  x1max = xabs
                  GoTo L20
L10:
                  s1 = s1 + (xabs / x1max) ^ 2
L20:
               GoTo L60
L30:
'
'              sum for small components.
'
               If (xabs <= x3max) Then GoTo L40
                  s3 = one + s3 * (x3max / xabs) ^ 2
                  x3max = xabs
                  GoTo L50
L40:
                  If (xabs <> zero) Then s3 = s3 + (xabs / x3max) ^ 2
L50:
L60:
            GoTo L80
L70:
'
'           sum for intermediate components.
'
            s2 = s2 + xabs ^ 2
L80:
      Next i
'
'     calculation of norm.
'
      If (s1 = zero) Then GoTo L100
         myenorm = x1max * dsqrt(s1 + (s2 / x1max) / x1max)
         GoTo L130
L100:
         If (s2 = zero) Then GoTo L110
            If (s2 >= x3max) Then _
              myenorm = dsqrt(s2 * (one + (x3max / s2) * (x3max * s3)))
            If (s2 < x3max) Then _
              myenorm = dsqrt(x3max * ((s2 / x3max) + (x3max * s3)))
            GoTo L120
L110:
            myenorm = x3max * dsqrt(s3)
L120:
L130:
'
'     last card of function enorm.
'
End Function

Private Sub lmpar(n As Single, ByRef r As Variant, ldr As Single, ipvt As Variant, _
  diag As Variant, qtb As Variant, delta As Double, ByRef par As Double, _
  ByRef x As Variant, ByRef sdiag As Variant, ByRef wa1 As Variant, ByRef wa2 As Variant)
      Dim i As Single
      Dim iter As Single
      Dim j As Single
      Dim jm1 As Single
      Dim jp1 As Single
      Dim k As Single
      Dim l As Single
      Dim nsing As Single
      Dim dxnorm As Double
      Dim dwarf As Double
      Dim fp As Double
      Dim gnorm As Double
      Dim parc As Double
      Dim parl As Double
      Dim paru As Double
      Dim p1 As Double
      Dim p001 As Double
      Dim sum As Double
      Dim temp As Double
      Dim zero As Double
      Dim dmppar As Double
      Dim enorm As Double
      p1 = 0.1
      p001 = 0.001
      zero = 0
'
'     dwarf is the smallest positive magnitude.
'
      dwarf = 1E-38
'
'     compute and store in x the gauss-newton direction. if the
'     jacobian is rank-deficient, obtain a least squares solution.
'
      nsing = n
      For j = 1 To n
         wa1(j) = qtb(j)
         If ((r(j, j) = zero) And (nsing = n)) Then nsing = j - 1
         If (nsing < n) Then wa1(j) = zero
      Next j
      If (nsing < 1) Then GoTo L50
      For k = 1 To nsing
         j = nsing - k + 1
         wa1(j) = wa1(j) / r(j, j)
         temp = wa1(j)
         jm1 = j - 1
         If (jm1 < 1) Then GoTo L30
         For i = 1 To jm1
            wa1(i) = wa1(i) - r(i, j) * temp
         Next i
L30:
      Next k
L50:
      For j = 1 To n
         l = ipvt(j)
         x(l) = wa1(j)
      Next j
'
'     initialize the iteration counter.
'     evaluate the function at the origin, and test
'     for acceptance of the gauss-newton direction.
'
      iter = 0
      For j = 1 To n
         wa2(j) = diag(j) * x(j)
      Next j
      dxnorm = myenorm(n, wa2)
      fp = dxnorm - delta
      If (fp <= p1 * delta) Then GoTo L220
'
'     if the jacobian is not rank deficient, the newton
'     step provides a lower bound, parl, for the zero of
'     the function. otherwise set this bound to zero.
'
      parl = zero
      If (nsing < n) Then GoTo L120
      For j = 1 To n
         l = ipvt(j)
         wa1(j) = diag(l) * (wa2(l) / dxnorm)
      Next j
      For j = 1 To n
         sum = zero
         jm1 = j - 1
         If (jm1 < 1) Then GoTo L100
         For i = 1 To jm1
            sum = sum + r(i, j) * wa1(i)
         Next i
L100:
         wa1(j) = (wa1(j) - sum) / r(j, j)
      Next j
      temp = myenorm(n, wa1)
      parl = ((fp / delta) / temp) / temp
L120:
'
'     calculate an upper bound, paru, for the zero of the function.
'
      For j = 1 To n
         sum = zero
         For i = 1 To j
            sum = sum + r(i, j) * qtb(i)
         Next i
         l = ipvt(j)
         wa1(j) = sum / diag(l)
      Next j
      gnorm = myenorm(n, wa1)
      paru = gnorm / delta
      If (paru = zero) Then paru = dwarf / dmin1(delta, p1)
'
'     if the input par lies outside of the interval (parl,paru),
'     set par to the closer endpoint.
'
      par = dmax1(par, parl)
      par = dmin1(par, paru)
      If (par = zero) Then par = gnorm / dxnorm
'
'     beginning of an iteration.
'
L150:
         iter = iter + 1
'
'        evaluate the function at the current value of par.
'
         If (par = zero) Then par = dmax1(dwarf, p001 * paru)
         temp = dsqrt(par)
         For j = 1 To n
            wa1(j) = temp * diag(j)
         Next j
         qrsolv n, r, ldr, ipvt, wa1, qtb, x, sdiag, wa2
         For j = 1 To n
            wa2(j) = diag(j) * x(j)
         Next j
         dxnorm = myenorm(n, wa2)
         temp = fp
         fp = dxnorm - delta
'
'        if the function is small enough, accept the current value
'        of par. also test for the exceptional cases where parl
'        is zero or the number of iterations has reached 10.
'
         If ((dabs(fp) <= p1 * delta) _
            Or ((parl = zero) And (fp <= temp) And (temp < zero)) _
            Or (iter = 10)) Then GoTo L220
'
'        compute the newton correction.
'
         For j = 1 To n
            l = ipvt(j)
            wa1(j) = diag(l) * (wa2(l) / dxnorm)
         Next j
         For j = 1 To n
            wa1(j) = wa1(j) / sdiag(j)
            temp = wa1(j)
            jp1 = j + 1
            If (n < jp1) Then GoTo L200
            For i = jp1 To n
               wa1(i) = wa1(i) - r(i, j) * temp
            Next i
L200:
         Next j
         temp = myenorm(n, wa1)
         parc = ((fp / delta) / temp) / temp
'
'        depending on the sign of the function, update parl or paru.
'
         If (fp > zero) Then parl = dmax1(parl, par)
         If (fp < zero) Then paru = dmin1(paru, par)
'
'        compute an improved estimate for par.
'
         par = dmax1(parl, par + parc)
'
'        end of an iteration.
'
         GoTo L150
L220:
'
'     termination.
'
      If (iter = 0) Then par = zero
'
'     last card of subroutine lmpar.
'
End Sub
  
Private Sub qrsolv(n As Single, ByRef r As Variant, ldr As Single, ByRef ipvt As Variant, _
ByRef diag As Variant, ByRef qtb As Variant, ByRef x As Variant, ByRef sdiag As Variant, ByRef wa As Variant)
      Dim i As Single
      Dim j As Single
      Dim jp1 As Single
      Dim k As Single
      Dim kp1 As Single
      Dim l As Single
      Dim nsing As Single
      Dim mycos As Double
      Dim mycotan As Double
      Dim p5 As Double
      Dim p25 As Double
      Dim qtbpj As Double
      Dim mysin As Double
      Dim sum As Double
      Dim mytan As Double
      Dim temp As Double
      Dim zero As Double
      p5 = 0.5
      p25 = 0.25
      zero = 0

'
'     copy r and (q transpose)*b to preserve input and initialize s.
'     in particular, save the diagonal elements of r in x.
'
      For j = 1 To n
         For i = j To n
            r(i, j) = r(j, i)
         Next i
         x(j) = r(j, j)
         wa(j) = qtb(j)
      Next j
'
'     eliminate the diagonal matrix d using a givens rotation.
'
      For j = 1 To n
'
'        prepare the row of d to be eliminated, locating the
'        diagonal element using p from the qr factorization.
'
         l = ipvt(j)
         If (diag(l) = zero) Then GoTo L90
         For k = j To n
            sdiag(k) = zero
         Next k
         sdiag(j) = diag(l)
'
'        the transformations to eliminate the row of d
'        modify only a single element of (q transpose)*b
'        beyond the first n, which is initially zero.
'
         qtbpj = zero
         For k = j To n
'
'           determine a givens rotation which eliminates the
'           appropriate element in the current row of d.
'
            If (sdiag(k) = zero) Then GoTo L70
            If (dabs(CDbl(r(k, k))) >= dabs(CDbl(sdiag(k)))) Then GoTo L40
               mycotan = r(k, k) / sdiag(k)
               mysin = p5 / dsqrt(p25 + p25 * mycotan ^ 2)
               mycos = mysin * mycotan
               GoTo L50
L40:
               mytan = sdiag(k) / r(k, k)
               mycos = p5 / dsqrt(p25 + p25 * mytan ^ 2)
               mysin = mycos * mytan
L50:
'
'           compute the modified diagonal element of r and
'           the modified element of ((q transpose)*b,0).
'
            r(k, k) = mycos * r(k, k) + mysin * sdiag(k)
            temp = mycos * wa(k) + mysin * qtbpj
            qtbpj = -mysin * wa(k) + mycos * qtbpj
            wa(k) = temp
'
'           accumulate the tranformation in the row of s.
'
            kp1 = k + 1
            If (n < kp1) Then GoTo L70
            For i = kp1 To n
               temp = mycos * r(i, k) + mysin * sdiag(i)
               sdiag(i) = -mysin * r(i, k) + mycos * sdiag(i)
               r(i, k) = temp
            Next i
L70:
L80:
          Next k
L90:
'
'        store the diagonal element of s and restore
'        the corresponding diagonal element of r.
'
         sdiag(j) = r(j, j)
         r(j, j) = x(j)
      Next j
'
'     solve the triangular system for z. if the system is
'     singular, then obtain a least squares solution.
'
      nsing = n
      For j = 1 To n
         If (sdiag(j) = zero And nsing = n) Then nsing = j - 1
         If (nsing < n) Then wa(j) = zero
      Next j
      If (nsing < 1) Then GoTo L150
      For k = 1 To nsing
         j = nsing - k + 1
         sum = zero
         jp1 = j + 1
         If (nsing < jp1) Then GoTo L130
         For i = jp1 To nsing
            sum = sum + r(i, j) * wa(i)
         Next i
L130:
         wa(j) = (wa(j) - sum) / sdiag(j)
      Next k
L150:
'
'     permute the components of z back to components of x.
'
      For j = 1 To n
         l = ipvt(j)
         x(l) = wa(j)
      Next j
      
'
'     last card of subroutine qrsolv.
'
 

End Sub

Private Sub fdjac2(m As Single, n As Single, x As Variant, _
  fvec As Variant, ByRef fjac As Variant, ldfjac As Single, iflag As Single, _
  epsfcn As Double, wa As Variant)
  Dim i As Single
  Dim j As Single
  Dim eps As Double
  Dim epsmch As Double
  Dim h As Double
  Dim temp As Double
  Dim zero As Double
  zero = 0
  jacobian = fjac
  '     epsmch is the machine precision.
  epsmch = 1E-16
'
      eps = dsqrt(dmax1(epsfcn, epsmch))
      For j = 1 To n
         temp = x(j)
         h = eps * dabs(temp)
         If (h = zero) Then h = eps
         x(j) = temp + h
         LevenbergCostFunction m, n, x, wa, iflag
         If (iflag = 0) Then GoTo L30
         x(j) = temp
         For i = 1 To m
            fjac(i, j) = (wa(i) - fvec(i)) / h
            'added in order to record jacobian at the
            'extremum since fjac is further edited
            jacobian(i, j) = Sqr(2) * fjac(i, j)
         Next i
      Next j
L30:
'     last card of subroutine fdjac2.
  End Sub

Private Sub LevenbergCostFunction(m As Single, n As Single, x As Variant, _
              ByRef fvec As Variant, ByRef iflag As Single)
  fvec = Application.Run(FunctionName_, x, IndependentVariables_)
  Dim i As Single
  For i = 1 To m
    fvec(i) = fvec(i) - ObservedVariables_(i)
  Next i
End Sub

Public Function CalibrateParameters(FunctionName As String, _
                                    xvec As Variant, _
                                    yvec As Variant, _
                                    params As Variant, _
                                    Optional ftol As Double = 0.00001, _
                                    Optional xtol As Double = 0.00001, _
                                    Optional gtol As Double = 0.00001, _
                                    Optional maxfeval As Single = 700, _
                                    Optional epsfcn As Double = 0.000000001) As Variant
  FunctionName_ = FunctionName
  Dim m As Single
  Dim n As Single
  m = UBound(yvec)
  n = UBound(params)
  'm = 10
  'n = 3
  IndependentVariables_ = xvec
  ObservedVariables_ = yvec
  Dim x As Variant
  x = params
  'ReDim x(1 To n)
  'x(1) = 0.1
  'x(2) = 0.1
  'x(3) = 0.1
  
  
  ReDim fvec(1 To m)
  ReDim diag(1 To n)
  Dim mode As Single
  mode = 1
  Dim factor As Double
  factor = 1
  Dim nprint As Single
  nprint = 0
  
  Dim info As Single
  Dim nfeval As Single
  Dim ldfjac As Single
  ldfjac = m
  ReDim fjac(1 To ldfjac, 1 To n)
  ReDim ipvt(1 To n)
  ReDim qtf(1 To n)
  ReDim wa1(1 To n)
  ReDim wa2(1 To n)
  ReDim wa3(1 To n)
  ReDim wa4(1 To m)
  
  lmdif "MyFunction", m, n, x, fvec, ftol, xtol, gtol, maxfeval, epsfcn, diag, mode, factor, nprint, info, nfeval _
        , fjac, ldfjac, ipvt, qtf, wa1, wa2, wa3, wa4
  
  CalibrateParameters = x

End Function




