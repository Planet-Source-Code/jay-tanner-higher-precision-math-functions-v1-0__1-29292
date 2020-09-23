Attribute VB_Name = "Hi_Precision_Math_Module"
  Option Explicit
  
' Higher Precision Advanced Math Functions for Special Applications
' Author: Jay Tanner
'
' For MS Visual BASIC v5/6
'
' Any bugs or suggestions
' E-Mail:
' Jay@NeoProgrammics.com
'
' ====================================================================
' These functions use the 29 digit precision decimal data type.
'
' Potential accuracy: Up to ± 0.000000000000000000000000001
' The potential accuracy may sometimes be a tiny bit less
' due to the inexact nature of trancendental functions and the
' internal rounding errors inherent in digital computations.
'
' -----------------------------------------------------------
' The higher precision functions in this implementation are:
'
' Sine(x)    - Circular sine for degree arguments
' Cosine(x)  - Circular cosine for degree arguments
' Tangent(x) - Circular tangent for degree arguments
'
' ExpF(x)    - Natural exponential function
' LogE(x)      - Natural logarithm function
' Log10(x)   - Base 10 logarithm function
' AntiLog    - Base 10 antilog function

' Sinh(x)    - Hyperbolic sine function
' Cosh(x)    - Hyperbolic cosine function
' Tanh(x)    - Hyperbolic tangent function
' ArcSinh(x) - Hyperbolic arc sine function
' ArcCosh(x) - Hyperbolic arc cosine function
' ArcTanh(x) - Hyperbolic arc tangent function
'
' Square_Root(x) - High precision square root function
' Cube_Root(x)   - High precision cube root function
'
' More higher precision functions are planned for the future.
'
'
' ====================================================================
' High precision square root function
'
' Valid (X) argument range: > 0
'
' If an error is detected, then an empty result is returned.
'
  Public Function Square_Root(X_Arg)
  
  On Error GoTo ERROR_HANDLER
  
  Dim X    As Variant  ' The input argument
  Dim A    As Variant  ' Current approximation
  Dim B    As Variant  ' Previous approximation
  
  Dim k    As Integer  ' Safety counter to prevent infinite loop
      k = 0
        
' Read input argument
  X = Trim(X_Arg): If X = "" Or X = "-" Then X = 0
  
' Return empty string if non-numeric argument.
  If Not IsNumeric(X) Then Square_Root = "": Exit Function
    
' Convert argument to decimal data type
  X = CDec(X)
  
' Return zero if X = 0
  If X = 0 Then Square_Root = 0: Exit Function
  
' Return empty string if negative argument.
  If X < 0 Then Square_Root = "": Exit Function
 
' Initialize loop variables
  A = Sqr(X)  ' 1st approximation = Normal VB Sqr(x) function
  A = CDec(A)
  B = CDec(1)
  
ITERATE:
  B = (A + X / A) / 2 ' Compute next approx (B) from (A)
  
' Check if finished
  If (B = A) Or (k >= 20) Then Square_Root = B: Exit Function

' Update approximation and iteration loop counter
  A = B
  k = k + 1
  GoTo ITERATE

  Exit Function
  
ERROR_HANDLER:
  Square_Root = ""
  MsgBox Error$, vbCritical, " PROGRAM ERROR"

  End Function

' ====================================================================
' High precision cube root function
'
' Valid (X) argument range: ± X
'
' If an error is detected, then an empty result is returned.
'
  Public Function Cube_Root(X_Arg)
  
  On Error GoTo ERROR_HANDLER
  
  Dim X       As Variant ' The input argument
  Dim A       As Variant ' Current approximation
  Dim B       As Variant ' Previous approximation
  Dim NegFlag As Boolean ' Negative argument flag
  
  Dim k    As Integer    ' Safety counter to prevent infinite loop
      k = 0
        
' Read input argument
  X = Trim(X_Arg): If X = "" Or X = "-" Then X = 0
  
' Return empty string if non-numeric argument.
  If Not IsNumeric(X) Then Cube_Root = "": Exit Function
    
' Convert argument to decimal data type
  X = CDec(X)
  
' Return zero if X = 0
  If X = 0 Then Cube_Root = 0: Exit Function
  
' Account for negative argument, if indicated.
  If X < 0 Then
     NegFlag = True
     X = Abs(X)
  Else
     NegFlag = False
  End If
  
' Initialize loop variables
  A = X ^ (1 / 3) ' 1st approximation = Normal VB  X^(1/3) function
  A = CDec(A)
  B = CDec(1)
  
ITERATE:
  B = ((2 * A) + X / (A * A)) / 3  ' Compute next approx (B) from (A)
  
' Check if finished
  If (B = A) Or (k >= 20) Then
      If NegFlag = True Then B = -B
      Cube_Root = B
      Exit Function
  End If
  
' Update approximation and iteration loop counter
  A = B
  k = k + 1
  GoTo ITERATE

  Exit Function
  
ERROR_HANDLER:
  Cube_Root = ""
  MsgBox Error$, vbCritical, " PROGRAM ERROR"
  
  End Function
 
' ====================================================================
' High precision exponential function
'
' Valid (X) argument range: 0 to ±65
'
' If an error is detected, then an empty result is returned.
'
  Public Function ExpF(X_Arg)

  On Error GoTo ERROR_HANDLER
  
  Dim X        As Variant ' The input argument
  Dim FactX    As Variant ' Factorial seed
  Dim Term     As Variant ' Current series term value
  Dim PwrX     As Variant ' Power of X
  Dim S        As Variant ' Series summation accumulator
  Dim NegFlag  As Boolean ' Negative argument flag
  
' Error tolerance limit for series
  Dim ET   As Variant
      ET = 1E-29
' 65.37052415368304665919611913
' Read input argument
  X = Trim(X_Arg): If X = "" Then X = 0
  
' Return empty string for non-numeric argument or if numeric
' argument is not in the valid range.
  If Not IsNumeric(X) Then ExpF = "": Exit Function
  If Abs(X) > CDec("65.37052415368304665919611913") Then
     ExpF = ""
     Exit Function
  End If
  
  If X < 0 Then NegFlag = True Else NegFlag = False
  X = CDec(X)
  X = Abs(X)

' Initialize series variables
  FactX = CDec(1)
   PwrX = X
   Term = CDec(1)
      S = CDec(0)
   
' Exponential series summation loop
  While Term > ET
  
  Term = Term * X / FactX
  S = S + Term
  FactX = FactX + 1
  
  Wend

' Add 1 to summation result to finish
  S = 1 + S
  
' Return exponential function value of series
  If NegFlag = True Then ExpF = 1 / S Else ExpF = S
  
  Exit Function
  
ERROR_HANDLER:
  ExpF = ""
  MsgBox Error$, vbCritical, " PROGRAM ERROR"
  
  End Function

' ====================================================================
' High precision natural logarithm function.
'
' This is a core routine used by the Ln(x) and Log10(x) functions.
' The Ln(x) function computes Log10(x) and converts it into the
' natural logarithm equivalent to speed up the computation process,
' especially for larger arguments.
'
' Valid (X) argument range: > 0
'
' If an error is detected, then an empty result is returned.
'
  Public Function LogE(X_Arg)
  
  On Error GoTo ERROR_HANDLER
  
  Dim X         As Variant ' The input argument
  Dim FactX     As Variant ' Factorial seed value
  Dim Term      As Variant ' Value of current series term
  Dim PwrX      As Variant ' Power of X argument
  Dim S         As Variant ' Series summation accumulator
  Dim W         As Variant ' Temp work
  Dim FracFlag  As Boolean ' Flag for fractional value (0 < X < 1)
  
' Error tolerance limit for series
  Dim ET   As Variant
      ET = 1E-29
    
' Read input argument
  X = Trim(X_Arg)
  
' Return empty string if non-numeric argument.
  If X = "" Or X = "-" Then LogE = "": Exit Function
  If Not IsNumeric(X) Then LogE = "": Exit Function
  
' Convert argument to decimal data type
  X = CDec(X)
    
' Return empty string if negative or zero argument.
  If X <= 0 Then LogE = "": Exit Function
  
' Return zero if X = 1
  If X = 1 Then LogE = 0: Exit Function
  
' Account for fractional X (0 < X < 1), if indicated.
  If X < 1 Then
     FracFlag = True
     X = 1 / X
  Else
     FracFlag = False
  End If
  
' Initialize series variables
  FactX = CDec(1)
   Term = CDec(1)
   PwrX = CDec(1)
      S = CDec(0)
      W = (X - 1) / X
  
' Natural logarithm series summation loop
  While Abs(Term) > ET
  
  Term = PwrX * W / FactX
  S = S + Term
  PwrX = PwrX * W
  FactX = FactX + 1
  
  Wend
  
' Account for fractional X value, if indicated.
  If FracFlag = True Then S = -S
  
' Return computed LogE(X) value
  LogE = S
  
  Exit Function
  
ERROR_HANDLER:
  LogE = ""
  MsgBox Error$, vbCritical, " PROGRAM ERROR"
  
  End Function

' ====================================================================
' High precision hyperbolic sine function
'
' Valid (X) argument range: 0 to ±65
'
' If an error is detected, then an empty result is returned.
'
  Public Function Sinh(X_Arg)
  
  On Error GoTo ERROR_HANDLER
  
  Dim X       As Variant ' The input argument
  Dim W       As Variant ' Temp work
  Dim NegFlag As Boolean ' Negative argument flag
  
' Read input argument
  X = Trim(X_Arg): If X = "" Then X = 0
  
' Return empty string if non-numeric argument
  If Not IsNumeric(X) Then Sinh = "": Exit Function

' Convert argument to decimal data type
  X = CDec(X)
  
' Check sign of argument
  If X < 0 Then NegFlag = True Else NegFlag = False
  X = Abs(X)
  
' Return empty string if numeric argument outside valid range
  If Abs(X) > CDec("65.37052415368304665919611913") Then
     Sinh = ""
     Exit Function
  End If
  
' Call high-precision exponential function as first step
  W = ExpF(X)
  
' Return empty string if error detected
  If Not IsNumeric(W) Then Sinh = "": Exit Function
  
' Complete the Sinh(X) computation
  W = (W - (1 / W)) / 2
  
' Account for sign
  If NegFlag = True Then W = -W
  
' Return the Sinh(X) value
  Sinh = W
  
  Exit Function
  
ERROR_HANDLER:
  Sinh = ""
  MsgBox Error$, vbCritical, " PROGRAM ERROR"
  
  End Function

' ====================================================================
' High precision hyperbolic cosine function
'
' Valid (X) argument range: 0 to ±65
'
' If an error is detected, then an empty result is returned.
'
  Public Function Cosh(X_Arg)
  
  On Error GoTo ERROR_HANDLER
  
  Dim X       As Variant ' The input argument
  Dim W       As Variant ' Temp work
  
' Read input argument
  X = Trim(X_Arg): If X = "" Then X = 0
  X = Abs(CDec(X))
  
' Return empty string for non-numeric argument or if numeric
' argument is not in the valid range.
  If Not IsNumeric(X) Then Cosh = "": Exit Function
  If Abs(X) > CDec("65.37052415368304665919611913") Then
     Cosh = ""
     Exit Function
  End If
  
' Call exponential function as first step
  W = ExpF(X)
  
' Return an empty string if error detected
  If Not IsNumeric(W) Then Cosh = "": Exit Function
  
' Return the computed cosh(x) value
  Cosh = (W + (1 / W)) / 2
  
  Exit Function
  
ERROR_HANDLER:
  Cosh = ""
  MsgBox Error$, vbCritical, " PROGRAM ERROR"
  
  End Function
  
' ====================================================================
' High precision hyperbolic tangent function.  This function calls
' the Sinh() and Cosh() functions.
'
' Valid (X) argument range: 0 to ±65
'
' If an error is detected, then an empty result is returned.
'
  Public Function Tanh(X_Arg)
  
  On Error GoTo ERROR_HANDLER
  
  Dim X       As Variant ' The input argument
  Dim W       As Variant ' Temp work
  
' Read input argument
  X = Trim(X_Arg): If X = "" Then X = 0
  
' Return empty string for non-numeric argument.
  If Not IsNumeric(X) Then Tanh = "": Exit Function
  
' Convert argument to decimal data type
  X = CDec(X)
  
' Return empty string if numeric argument out of valid range.
  If Abs(X) > CDec("65.37052415368304665919611913") Then
     Tanh = ""
     Exit Function
  End If
     
' Return computed Tanh(x) value.
  Tanh = Sinh(X) / Cosh(X)
  
  Exit Function
  
ERROR_HANDLER:
  Tanh = ""
  MsgBox Error$, vbCritical, " PROGRAM ERROR"
  
  End Function

' ====================================================================
' High precision inverse hyperbolic sine function
'
' Valid (X) argument range: ± X
'
' If an error is detected, then an empty result is returned.
'
  Public Function ArcSinh(X_Arg)

  On Error GoTo ERROR_HANDLER
  
  Dim X As Variant  ' The input argument
  Dim W As Variant  ' Temp work
  
' Read input argument
  X = Trim(X_Arg): If X = "" Or X = "-" Then X = 0
  
' Return empty string if non-numeric argument.
  If Not IsNumeric(X) Then ArcSinh = "": Exit Function
  
' Convert argument to decimal data type.
  X = CDec(X)
      
' Compute the ArcSinh(x) value
  W = X + Square_Root(X * X + 1)
  W = Ln(W)
  
' Return empty string if error detected.
  If Not IsNumeric(W) Then ArcSinh = "": Exit Function
  
' Return computed ArcSinh value
  ArcSinh = W

  Exit Function
  
ERROR_HANDLER:
  ArcSinh = ""
  MsgBox Error$, vbCritical, " PROGRAM ERROR"
  
  End Function
  
' ====================================================================
' High precision inverse hyperbolic cosine function
'
' Valid (X) argument range: X >= 1
'
' If an error is detected, then an empty result is returned.
'
  Public Function ArcCosh(X_Arg)

  On Error GoTo ERROR_HANDLER
  
  Dim X As Variant  ' The input argument
  Dim W As Variant  ' Temp work
  
' Read input argument
  X = Trim(X_Arg): If X = "" Or X = "-" Then X = 0
  
' Return empty string if non-numeric argument.
  If Not IsNumeric(X) Then ArcCosh = "": Exit Function
  
' Convert argument to decimal data type.
  X = CDec(X)
  
' Return empty string if numeric argument is not in the valid range.
  If X < 1 Then ArcCosh = "": Exit Function
        
' Compute the ArcCosh(x) value
  W = X + Square_Root(X * X - 1)
  W = Ln(W)
  
' Return empty string if error detected.
  If Not IsNumeric(W) Then ArcCosh = "": Exit Function
  
' Return computed ArcSinh value
  ArcCosh = W

  Exit Function
  
ERROR_HANDLER:
  ArcCosh = ""
  MsgBox Error$, vbCritical, " PROGRAM ERROR"
  
  End Function
  
' ====================================================================
' High precision inverse hyperbolic tangent function
'
' Valid (X) argument range: -1 < X < +1
'
' If an error is detected, then an empty result is returned.
'
  Public Function ArcTanh(X_Arg)

  On Error GoTo ERROR_HANDLER
  
  Dim X        As Variant  ' The input argument
  Dim W        As Variant  ' Temp work
  Dim SignFlag As Boolean       ' Negative argument flag
  
' Read input argument
  X = Trim(X_Arg): If X = "" Or X = "-" Then X = 0
  
' Return empty string if non-numeric argument.
  If Not IsNumeric(X) Then ArcTanh = "": Exit Function
  
' Convert argument to decimal data type.
  X = CDec(X)
  
' Account for sign of argument
  If X < 0 Then SignFlag = True Else SignFlag = False
  X = Abs(X)
  
' Return empty string if numeric argument is not in the valid range.
  If Abs(X) >= 1 Then ArcTanh = "": Exit Function
        
' Compute the ArcTanh(x) value
  W = (1 + X) / (1 - X)
  W = Ln(W)
  
' Return empty string if error detected.
  If Not IsNumeric(W) Then ArcTanh = "": Exit Function
  
' Compute ArcTanh(x) value
  W = W / 2
  
' Account for sign
  If SignFlag = True Then W = -W
  
' Return computed ArcSinh value
  ArcTanh = W

  Exit Function
  
ERROR_HANDLER:
  ArcTanh = ""
  MsgBox Error$, vbCritical, " PROGRAM ERROR"
  
  End Function
  
' ====================================================================
' High precision circular sine function for degree arguments

  Public Function Sine(Deg_Arg)
  
  On Error GoTo ERROR_HANDLER
  
  Dim X         As Variant ' The input argument
  Dim FactX     As Variant ' Factorial seed value
  Dim Term      As Variant ' Value of current series term
  Dim PwrX      As Variant ' Power of X argument
  Dim S         As Variant ' Series summation accumulator
  Dim W         As Variant ' Temp work
  Dim NegFlag   As Boolean ' Flag for negative argument
  
  Dim i         As Variant ' Sign control
      i = -1
      
  Dim Pi        As Variant ' Value of Pi constant
      Pi = CDec("3.14159265358979323846264338327950288")
  
' Error tolerance limit for series
  Dim ET   As Variant
      ET = 1E-29
      
' Read input argument
  X = Trim(Deg_Arg)
  
' Return empty string if non-numeric argument.
  If X = "" Or X = "-" Then Sine = "": Exit Function
  If Not IsNumeric(X) Then Sine = "": Exit Function
  
' Convert argument to decimal data type
  X = CDec(X)
          
' Account for sign
  If X < 0 Then NegFlag = True Else NegFlag = False
  X = Abs(X)
          
' If >= 360 degrees, then subtract 360
  If X >= 360 Then X = X - 360
          
' Account for special exact values.
  If X = 0 Then Sine = 0: Exit Function
  If X = 90 Then Sine = 1: Exit Function
  If X = 180 Then Sine = 0: Exit Function
  If X = 270 Then Sine = -1: Exit Function
          
' Convert X degrees to radians
  X = X * Pi / 180
          
' Initialize series variables
  FactX = CDec(3)
   Term = CDec(X)
      S = X
  
' Circular sine series summation loop
  While Abs(Term) > ET
    
  Term = Term * X / FactX * X / (FactX - 1)
  S = S + Term * i
  FactX = FactX + 2
  i = -i
  
  Wend

' Account for sign, as indicated.
  If NegFlag = True Then S = -S
  
' Return computed Sine(X) value
  Sine = S
  
  Exit Function
  
ERROR_HANDLER:
  Sine = ""
  MsgBox Error$, vbCritical, " PROGRAM ERROR"
  
  End Function

' ====================================================================
' High precision circular cosine function for degree arguments

  Public Function Cosine(Deg_Arg)
  
  On Error GoTo ERROR_HANDLER
  
  Dim X         As Variant ' The input argument
  Dim FactX     As Variant ' Factorial seed value
  Dim Term      As Variant ' Value of current series term
  Dim PwrX      As Variant ' Power of X argument
  Dim S         As Variant ' Series summation accumulator
  Dim W         As Variant ' Temp work
  Dim NegFlag   As Boolean ' Flag for negative argument
  
  Dim i         As Variant ' Sign control
      i = -1
      
  Dim Pi        As Variant ' Value of Pi constant
      Pi = CDec("3.14159265358979323846264338327950288")
  
' Error tolerance limit for series
  Dim ET   As Variant
      ET = 1E-29
    
' Read input argument
  X = Trim(Deg_Arg)
  
' Return empty string if non-numeric argument.
  If X = "" Or X = "-" Then Cosine = "": Exit Function
  If Not IsNumeric(X) Then Cosine = "": Exit Function
  
' Convert argument to decimal data type
  X = CDec(X)
  X = Abs(X)
                    
' If x >= 360 degrees, then subtract 360
  If X >= 360 Then X = X - 360
          
' Account for special exact values.
  If X = 0 Then Cosine = 1: Exit Function
  If X = 90 Or X = 270 Then Cosine = 0: Exit Function
  If X = 180 Then Cosine = -1: Exit Function

' Convert X degrees to radians
  X = X * Pi / 180
          
' Initialize series variables
  FactX = CDec(2)
   Term = CDec(1)
      S = 1
  
' Circular cosine series summation loop
  While Abs(Term) > ET
    
  Term = Term * X / FactX * X / (FactX - 1)
  S = S + Term * i
  FactX = FactX + 2
  i = -i
  
  Wend
  
' Return computed Cosine(X) value
  Cosine = S
  
  Exit Function
  
ERROR_HANDLER:
  Cosine = ""
  MsgBox Error$, vbCritical, " PROGRAM ERROR"
  End Function

' ====================================================================
' High precision circular Tangent(x) function for degree arguments.
' This function calls the Sine() and Cosine() functions.
  
  Public Function Tangent(X_Arg)
  
  On Error GoTo ERROR_HANDLER
  
  Dim V As Variant ' Temp work
  Dim W As Variant ' Temp work
  Dim X As Variant ' The input argument
  
  X = Trim(X_Arg)
  
' Return empty string if non-numeric argument.
  If X = "" Or X = "-" Then Tangent = "": Exit Function
  If Not IsNumeric(X) Then Tangent = "": Exit Function
  
' Return empty string if error detected
  V = Sine(X)
  W = Cosine(X)
  If V = "" Or W = "" Then Tangent = "": Exit Function
  
' Return empty string if infinite result
  If W = 0 Then Tangent = "": Exit Function
  
' Return computed Tangent(x) value
  Tangent = V / W
  
  Exit Function
  
ERROR_HANDLER:
  Tangent = ""
  MsgBox Error$, vbCritical, " PROGRAM ERROR"
  
  End Function
  
' ====================================================================
' High precision base 10 logarithm function

' If an error is detected, then an empty result is returned.

  Public Function Log10(X_Arg)
  
  On Error GoTo ERROR_HANDLER
  
  Dim X As Variant  ' The input argument
  Dim W As Variant  ' Temp work
  Dim Ip As Variant ' Integer part of X argument
  Dim Dp As Variant ' Decimal part of X argument
  
  Dim LowFlag As Boolean ' Special low value flag for X < 1
  
  Dim V As Variant ' Natural logarithm of 10
      V = CDec("2.30258509299404568401799145468")
  
' Read the input argument
  X = Trim(X_Arg)
    
' Return empty result if non-numeric argument
  If X = "" Or X = "-" Then Log10 = "": Exit Function
  If Not IsNumeric(X) Then Log10 = "": Exit Function
    
  X = CDec(X)
  
' Account for special case where 0 < X < 1
  If X > 0 And X < 1 Then LowFlag = True: X = 1 / X
  
  Ip = 0
  
' For large numbers, split X argument into integer and fractional
' parts.  This algorithm was included to speed up the process based
' on the rules of base 10 logarithms.  Without this adjustment, the
' computation of logarithms of large numbers would take much longer.
  If X >= 10 Then
     If Right(X, 1) = "." Then X = X & "0"
     W = InStr(X, ".")
     If W = 0 Then X = X & ".0"
     W = InStr(X, ".")
     Ip = Left(X, W - 1)
     Dp = Mid(X, W + 1, Len(X))
     X = Left(Ip, 1) & "." & Mid(Ip, 2, Len(Ip)) & Dp
     Ip = CDec(Len(Ip)) - 1
  End If
  
' Compute the natural logarithm of the X argument
  W = LogE(X): If W = "" Then Log10 = "": Exit Function
  
' Compute Log10(x) value
  W = Ip + (W / V)
  If LowFlag = True Then W = -W
  
' Return computed Log10(x) value
  Log10 = W
  
  Exit Function
  
ERROR_HANDLER:
  Log10 = ""
  MsgBox Error$, vbCritical, " PROGRAM ERROR"
  
  End Function
  
' ==============================================================================
' Base 10 antilog function.  This function is the opposite of the
' base 10 logarithm function.

  Public Function AntiLog(X_Arg)
  
  Dim X As Variant ' The input argument
  Dim W As Variant ' Temp work
  
  Dim V As Variant ' Natural logarithm of 10
      V = CDec("2.30258509299404568401799145468")
    
' Read input argument
  X = Trim(X_Arg)
  
' Return empty result if non-numeric argument
  If X = "" Or X = "-" Then AntiLog = "": Exit Function
  If Not IsNumeric(X) Then AntiLog = "": Exit Function

' Convert argument to decimal data type
  X = CDec(X)

' Handle special case of integer arguments from ±1 to ±28
  If Right(X, 1) = "." Then X = Left(X, Len(X) - 1)
  If InStr(X, ".") = 0 And Abs(X) < 29 Then
     W = 1 & String(Abs(X), "0")
     If X < 0 Then AntiLog = 1 / W Else AntiLog = W
     Exit Function
  End If
  
' Compute antilog value
  W = ExpF(X * V)
  
' Return empty result if error detected.
  If W = "" Then AntiLog = "": Exit Function
  
' Return computed antilog value
  AntiLog = W
  
  End Function


' ====================================================================
' Natural logarithm function shell for the LogE(x) function.
'
' This function computes the base 10 logarithm to speed up the process
' of computing the natural logarithms of larger arguments.  It first
' computes the base 10 value and then converts it into the natural
' logarithm equivalent by multiplying by a conversion constant which
' is simply the natural logarithm of 10.
'
' If an error is detected, then an empty result is returned.

  Public Function Ln(X_Arg)
  
  On Error GoTo ERROR_HANDLER
  
  Dim X As Variant  ' The input argument
  Dim W As Variant  ' Temp work
  Dim Ip As Variant ' Integer part of X argument
  Dim Dp As Variant ' Decimal part of X argument
  
  Dim LowFlag As Boolean ' Low value flag for arguments < 1
  
  Dim V As Variant ' Natural logarithm of 10
      V = CDec("2.30258509299404568401799145468")
  
' Read the input argument
  X = Trim(X_Arg)
    
' Return empty result if non-numeric argument
  If X = "" Or X = "-" Then Ln = "": Exit Function
  If Not IsNumeric(X) Then Ln = "": Exit Function
    
  X = CDec(X)
  
' Account for special case where 0 < X < 1
  If X > 0 And X < 1 Then LowFlag = True: X = 1 / X
  
  Ip = 0
  
' For large numbers, split X argument into integer and fractional
' parts.  This algorithm was included to speed up the process based
' on the rules of base 10 logarithms.  Without this adjustment, the
' computation of logarithms of large numbers would take much longer.
  If X >= 10 Then
     If Right(X, 1) = "." Then X = X & "0"
     W = InStr(X, ".")
     If W = 0 Then X = X & ".0"
     W = InStr(X, ".")
     Ip = Left(X, W - 1)
     Dp = Mid(X, W + 1, Len(X))
     X = Left(Ip, 1) & "." & Mid(Ip, 2, Len(Ip)) & Dp
     Ip = CDec(Len(Ip)) - 1
  End If
  
' Compute the natural logarithm of the X argument
  W = LogE(X): If W = "" Then Ln = "": Exit Function
  
' Compute natural logarithm value
  W = Ip * V + W
  
' Return computed LogE(x) value
  If LowFlag = True Then W = -W
  Ln = W
  
  Exit Function
  
ERROR_HANDLER:
  Ln = ""
  MsgBox Error$, vbCritical, " PROGRAM ERROR"
  
  End Function
  
