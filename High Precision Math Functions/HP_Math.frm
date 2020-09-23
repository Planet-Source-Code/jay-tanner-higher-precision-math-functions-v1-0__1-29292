VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Higher Precision Math Functions Module Testing Interface"
   ClientHeight    =   5668
   ClientLeft      =   39
   ClientTop       =   312
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5668
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cube_Root_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cube Root"
      Height          =   273
      Left            =   3588
      TabIndex        =   17
      Top             =   1508
      Width           =   949
   End
   Begin VB.CommandButton ArcTanh_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ArcTanh"
      Height          =   273
      Left            =   5460
      TabIndex        =   11
      Top             =   936
      Width           =   741
   End
   Begin VB.CommandButton ArcCosh_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ArcCosh"
      Height          =   273
      Left            =   4732
      TabIndex        =   10
      Top             =   936
      Width           =   741
   End
   Begin VB.CommandButton Tanh_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tanh"
      Height          =   273
      Left            =   3276
      TabIndex        =   8
      Top             =   936
      Width           =   585
   End
   Begin VB.CommandButton Cosh_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cosh"
      Height          =   273
      Left            =   2704
      TabIndex        =   7
      Top             =   936
      Width           =   585
   End
   Begin VB.CommandButton AntiLog_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "AntiLog10"
      Height          =   273
      Left            =   1612
      TabIndex        =   15
      Top             =   1508
      Width           =   897
   End
   Begin VB.CommandButton Square_Root_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Square Root"
      Height          =   273
      Left            =   2600
      TabIndex        =   16
      Top             =   1508
      Width           =   1001
   End
   Begin VB.CommandButton Clear_X_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clear  X Arg"
      Height          =   273
      Left            =   3432
      TabIndex        =   2
      Top             =   260
      Width           =   1001
   End
   Begin VB.CommandButton Log10_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Log10"
      Height          =   273
      Left            =   988
      TabIndex        =   14
      Top             =   1508
      Width           =   637
   End
   Begin VB.CommandButton LN_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ln"
      Height          =   273
      Left            =   572
      TabIndex        =   13
      Top             =   1508
      Width           =   429
   End
   Begin VB.CommandButton ArcSinh_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ArcSinh"
      Height          =   273
      Left            =   4004
      TabIndex        =   9
      Top             =   936
      Width           =   741
   End
   Begin VB.CommandButton Sinh_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sinh"
      Height          =   273
      Left            =   2132
      TabIndex        =   6
      Top             =   936
      Width           =   585
   End
   Begin VB.CommandButton Clear_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clear Work Area"
      Height          =   273
      Left            =   4992
      TabIndex        =   18
      Top             =   1872
      Width           =   1209
   End
   Begin VB.CommandButton Tangent_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tangent"
      Height          =   273
      Left            =   1300
      TabIndex        =   5
      Top             =   936
      Width           =   741
   End
   Begin VB.CommandButton Cosine_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cosine"
      Height          =   273
      Left            =   624
      TabIndex        =   4
      Top             =   936
      Width           =   689
   End
   Begin VB.CommandButton Sine_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sine"
      Height          =   273
      Left            =   52
      TabIndex        =   3
      Top             =   936
      Width           =   585
   End
   Begin VB.TextBox Work 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.34
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3445
      Left            =   52
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2184
      Width           =   6149
   End
   Begin VB.TextBox Arg 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.34
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   273
      Left            =   52
      MaxLength       =   31
      TabIndex        =   0
      Text            =   "1"
      Top             =   260
      Width           =   3341
   End
   Begin VB.CommandButton Exp_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Exp"
      Height          =   273
      Left            =   52
      TabIndex        =   12
      Top             =   1508
      Width           =   533
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   " Computed Work Output"
      ForeColor       =   &H00000000&
      Height          =   221
      Left            =   52
      TabIndex        =   24
      Top             =   1976
      Width           =   1677
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   " Square and Cube Roots"
      ForeColor       =   &H00000000&
      Height          =   221
      Left            =   2600
      TabIndex        =   23
      Top             =   1300
      Width           =   1937
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   2548
      X2              =   2548
      Y1              =   1248
      Y2              =   1820
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   2080
      X2              =   2080
      Y1              =   676
      Y2              =   1248
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   52
      X2              =   6188
      Y1              =   1820
      Y2              =   1820
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   " Exponential and Logarithmic Functions"
      ForeColor       =   &H00000000&
      Height          =   221
      Left            =   52
      TabIndex        =   22
      Top             =   1300
      Width           =   2457
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   52
      X2              =   6188
      Y1              =   1248
      Y2              =   1248
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   52
      X2              =   6188
      Y1              =   676
      Y2              =   676
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   " Hyperbolic Functions"
      ForeColor       =   &H00000000&
      Height          =   221
      Left            =   2132
      TabIndex        =   21
      Top             =   728
      Width           =   4069
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   " Trig Functions for Degrees"
      ForeColor       =   &H00000000&
      Height          =   221
      Left            =   52
      TabIndex        =   20
      Top             =   728
      Width           =   1989
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   " X  Argument Input"
      ForeColor       =   &H00000000&
      Height          =   221
      Left            =   52
      TabIndex        =   19
      Top             =   52
      Width           =   3341
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit

  Public Out  As Variant ' Holds computed output of function

' Higher precision advanced scientific math functions for VB 5/6
'
' v1.0 - Written in VB 5.0, but should work for VB 6 also.
'
' Author: Jay Tanner
'
' Any bugs or suggestions
' E-Mail:
' Jay@NeoProgrammics.com
'
' Higher precision math module testing interface
'
' This program consists of an interface to demonstrate the use of the VB decimal
' data type to compute some advanced math functions beyond the normal 16 digit
' limitations of VB.
'
' These functions were developed for special scientific purposes where a higher
' precision was necessary for comparing certain aspects of relativity theory
' with classical mechanics computations at very low relativistic velocities.
' They can also be useful for other purposes as well.
'
' Since some of these functions use the infinite Taylor series summations,
' extremely large values could take a long time to compute, or result in an
' overflow error, so care should be taken.  These functions were not intended
' to take extremely large arguments outside the range of general practicality.
'
' In the case of detected errors, the functions return an empty string instead
' of a specific error message.  Most errors are caught, but there are still some
' that may get by.  These errors can be reduced by allowing only values within
' a certain range to be presented to a function by filtering input arguments
' prior to calling the given function.
'
' The functions are contained in a separate external module which can be
' attached to a program as needed.
'
' They are limited by the scope of the decimal data type, but are suitible for
' most practical computations.
'
' The decimal data type can display values up to 29 digits long.  This 29 digit
' span includes both the integer and decimal parts combined.
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
' Log10(x)   - Base 10  logarithm function

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
' Some of the functions are used to build other dependent functions.
'
' ===========================================================
' More higher precision functions are planned for the future.
'
' -----------------------------------------------------------
' INFO ABOUT THE DECIMAL DATA TYPE FROM THE VB HELP FILES
'
' Decimal variables are stored as 96-bit (12-byte) unsigned integers scaled by
' a variable power of 10. The power of 10 scaling factor specifies the number
' of digits to the right of the decimal point, and ranges from 0 to 28.
'
' With a scale of 0 (no decimal places), the largest possible value is
' +/-79,228,162,514,264,337,593,543,950,335. With 28 decimal places, the
' largest value is +/-7.9228162514264337593543950335 and the smallest,
' non-zero value is +/-0.0000000000000000000000000001
'
' ==============================================================================
' CIRCULAR TRIGONOMETRIC FUNCTIONS

' Circular sine trigonometric function for degree arguments.
  Private Sub Sine_Button_Click()
    
  Out = Sine(Arg)
  
  If Out = "" Then
     PRT "Invalid Sine(x) argument."
     PRT "X degrees argument should be in the range 0 to ±360"
     PRT String(56, "-")
     Beep
     Exit Sub
  End If
  
' Print result
  PRT "Sine(" & Arg & ") deg"
  PRT Out
  PRT String(56, "-")
  
  End Sub
  
' ==============================================================================
' Circular cosine trigonometric function for degree arguments.

  Private Sub Cosine_Button_Click()
  
  Out = Cosine(Arg)
  
  If Out = "" Then
     PRT "Invalid Cosine(x) argument."
     PRT "X degrees argument should be in the range 0 to ±360"
     PRT String(56, "-")
     Beep
     Exit Sub
  End If
  
' Print result
  PRT "Cosine(" & Arg & ") deg"
  PRT Out
  PRT String(56, "-")
  
  End Sub
  
' ==============================================================================
' Circular tangent trigonometric function for degree arguments.

  Private Sub Tangent_Button_Click()

  Out = Tangent(Arg)
  
  If Out = "" Then
     PRT "Invalid Tangent(x) argument."
     PRT "±90 or ±270 degrees will cause an infinity error."
     PRT "It is best to use angles in the range 0 to ± 360 degrees."
     PRT String(56, "-")
     Beep
     Exit Sub
  End If
  
' Print result
  PRT "Tangent(" & Arg & ") deg"
  PRT Out
  PRT String(56, "-")

  End Sub
  
' ==============================================================================
' HYPERBOLIC FUNCTIONS


' Hyperbolic Sinh(x) function

  Private Sub Sinh_Button_Click()

  Out = Sinh(Arg)
  
  If Out = "" Then
     PRT "Invalid Sinh(x) argument."
     PRT "X should be in the range from"
     PRT "0 to ±65.37052415368304665919611913"
     PRT String(56, "-")
     Beep
     Exit Sub
  End If
  
' Print result
  PRT "Sinh(" & Arg & ")"
  PRT Out
  PRT String(56, "-")

  End Sub

' ==============================================================================
' Hyperbolic ArcSinh(x) function

  Private Sub ArcSinh_Button_Click()
  
  Out = ArcSinh(Arg)
  
  If Out = "" Then
     PRT "Invalid ArcSinh(x) argument."
     PRT String(56, "-")
     Beep
     Exit Sub
  End If
  
' Print result
  PRT "ArcSinh(" & Arg & ")"
  PRT Out
  PRT String(56, "-")
  
  End Sub
  
' ==============================================================================
' Hyperbolic Cosh(x) function

  Private Sub Cosh_Button_Click()
  
  Out = Cosh(Arg)
  
  If Out = "" Then
     PRT "Invalid Cosh(x) argument."
     PRT "X should be in the range from"
     PRT "0 to ±65.37052415368304665919611913"
     PRT String(56, "-")
     Beep
     Exit Sub
  End If
  
' Print result
  PRT "Cosh(" & Arg & ")"
  PRT Out
  PRT String(56, "-")
  
  End Sub
  
' ==============================================================================
' Hyperbolic ArcCosh(x) function.

  Private Sub ArcCosh_Button_Click()
  
  Out = ArcCosh(Arg)
  
  If Out = "" Then
     PRT "Invalid ArcCosh(x) argument."
     PRT "X must be at least 1 or greater."
     PRT String(56, "-")
     Beep
     Exit Sub
  End If
  
' Print result
  PRT "ArcCosh(" & Arg & ")"
  PRT Out
  PRT String(56, "-")
  
  End Sub
  
' ==============================================================================
' Hyperbolic Tanh(x) function

  Private Sub Tanh_Button_Click()
  
  Out = Tanh(Arg)
  
  If Out = "" Then
     PRT "Invalid Tanh(x) argument."
     PRT "X should be in the range from"
     PRT "0 to ±65.37052415368304665919611913"
     PRT String(56, "-")
     Beep
     Exit Sub
  End If
  
' Print result
  PRT "Tanh(" & Arg & ")"
  PRT Out
  PRT String(56, "-")
  
  End Sub
  
' ==============================================================================
' Hyperbolic ArcTanh(x) function

  Private Sub ArcTanh_Button_Click()
  
  Out = ArcTanh(Arg)
  
  If Out = "" Then
     PRT "Invalid ArcTanh(x) argument."
     PRT "X should be between -1 and +1 exclusive."
     PRT String(56, "-")
     Beep
     Exit Sub
  End If
  
' Print result
  PRT "ArcTanh(" & Arg & ")"
  PRT Out
  PRT String(56, "-")
  
  End Sub
  
' ==============================================================================
' Natural logarithm function

  Private Sub Ln_Button_Click()

    Out = Ln(Arg)
  
  If Out = "" Then
     PRT "Invalid Ln(x) argument."
     PRT "X should be > 0"
     PRT String(56, "-")
     Beep
     Exit Sub
  End If
  
' Print result
  PRT "Ln(" & Arg & ")"
  PRT Out
  PRT String(56, "-")

  End Sub

' ==============================================================================
' Natural exponential function.  This is the inverse of the natural LogE(x)
' function.

  Private Sub Exp_Button_Click()

  Out = ExpF(Arg)
  
  If Out = "" Then
     PRT "Invalid Exp(x) argument."
     PRT "X must be in the range from"
     PRT "0 to ±65.37052415368304665919611913"
     PRT String(56, "-")
     Beep
     Exit Sub
  End If
  
' Print result
  PRT "Exp(" & Arg & ")"
  PRT Out
  PRT String(56, "-")

  End Sub
  
' ==============================================================================
' Base 10 logarithm function.

  Private Sub Log10_Button_Click()

  Out = Log10(Arg)
  
  If Out = "" Then
     PRT "Invalid Log10(x) argument."
     PRT "X should be > 0"
     PRT String(56, "-")
     Beep
     Exit Sub
  End If
  
' Print result
  PRT "Log10(" & Arg & ")"
  PRT Out
  PRT String(56, "-")

  End Sub
  
' ==============================================================================
' Square root function.

  Private Sub Square_Root_Button_Click()

  Out = Square_Root(Arg)
  
  If Out = "" Then
     PRT "Invalid Square Root(x) argument."
     PRT "X should be a positive value."
     PRT String(56, "-")
     Beep
     Exit Sub
  End If
  
' Print result
  PRT "Square Root(" & Arg & ")"
  PRT Out
  PRT String(56, "-")

  End Sub
  
' ==============================================================================
' Cube Root function.

  Private Sub Cube_Root_Button_Click()

  Out = Cube_Root(Arg)
  
  If Out = "" Then
     PRT "Invalid Cube Root(x) argument."
     PRT String(56, "-")
     Beep
     Exit Sub
  End If
  
' Print result
  PRT "Cube Root(" & Arg & ")"
  PRT Out
  PRT String(56, "-")

  End Sub
  
' ==============================================================================
' Base 10 AntiLog(x) function
' 28.229141323711371885157146061
  Private Sub AntiLog_Button_Click()

  Out = AntiLog(Arg)
  
  If Out = "" Then
     PRT "Invalid AntiLog(x) argument."
     PRT "Maximum legal value is: ±28.229141323711371885157146061"
     PRT String(56, "-")
     Beep
     Exit Sub
  End If
  
' Print result
  PRT "AntiLog(" & Arg & ")"
  PRT Out
  PRT String(56, "-")

  End Sub
  
' ==============================================================================
' ==============================================================================
'  Output control commands for printing to yellow work output area

' Clear the work area
  Private Sub CLEAR()
  Work.Text = ""
  End Sub
' ==============================================================================
' Print out value of (Expression)
  Private Sub PRT(Expression)
  Work.Text = Work.Text & Expression & vbCrLf
  End Sub

' ==============================================================================
' Print a blank line
  Private Sub BLINE()
  Work.Text = Work.Text & vbCrLf
  End Sub

' ==============================================================================
' Button to clear the yellow work display

  Private Sub Clear_Button_Click()
  CLEAR
  End Sub

' ==============================================================================
' Button to clear X argument to zero

  Private Sub Clear_X_Button_Click()
  Arg.Text = 0
  End Sub

' ==============================================================================
' Routine to substitute zero for an empty argument when the X argument input
' text box loses focus.
  
  Private Sub Arg_LostFocus()
  If Trim(Arg) = "" Then Arg = 0
  End Sub

' ==============================================================================

