VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spherical mapping in pure VB, by KACI Lounes"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   483
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Render"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   10935
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5070
      Left            =   5760
      Picture         =   "Main.frx":0000
      ScaleHeight     =   338
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   1
      Top             =   120
      Width           =   9600
   End
   Begin VB.PictureBox D 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   120
      ScaleHeight     =   377
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   369
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'####################################################################
'##                  Author: Mr KACI Lounes                        ##
'##             Spherical Mapping in *Pure* VB Code!               ##
'##    Compile for more speed ! Mail me at KLKEANO@CARAMAIL.COM    ##
'##               Copyright © 2006 - KACI Lounes                   ##
'####################################################################

Option Explicit
Sub ClipLine(RX1!, RY1!, RX2!, RY2!, X1!, Y1!, X2!, Y2!, OutX1!, OutY1!, OutX2!, OutY2!)

 'A Liang-Barsky line clipping algorithm

 Dim PX1!, PY1!, PX2!, PY2!, U1!, U2!, Dx!, Dy!, P!, Q!, R!, Temp!, CT As Byte

 If (RX1 > RX2) Then Temp = RX1: RX1 = RX2: RX2 = Temp
 If (RY1 > RY2) Then Temp = RY1: RY1 = RY2: RY2 = Temp

 U1 = 0: U2 = 1: PX1 = X1: PY1 = Y1: PX2 = X2: PY2 = Y2
 Dx = (PX2 - PX1): Dy = (PY2 - PY1)

 P = -Dx: Q = (PX1 - RX1)
 If (P < 0) Then
  R = (Q / P): If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
 ElseIf (P > 0) Then
  R = (Q / P): If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
 ElseIf (Q < 0) Then
  CT = 1
 End If
 If CT = 0 Then
  P = Dx: Q = (RX2 - PX1)
  If (P < 0) Then
   R = (Q / P): If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
  ElseIf (P > 0) Then
   R = (Q / P): If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
  ElseIf (Q < 0) Then
   CT = 1
  End If
  If CT = 0 Then
   P = -Dy: Q = (PY1 - RY1)
   If (P < 0) Then
    R = (Q / P): If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
   ElseIf (P > 0) Then
    R = (Q / P): If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
   ElseIf (Q < 0) Then
    CT = 1
   End If
   If CT = 0 Then
    P = Dy: Q = (RY2 - PY1)
    If (P < 0) Then
     R = (Q / P): If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
    ElseIf (P > 0) Then
     R = (Q / P): If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
    ElseIf (Q < 0) Then
     CT = 1
    End If
    If CT = 0 Then
     If (U2 < 1) Then PX2 = (PX1 + (U2 * Dx)): PY2 = (PY1 + (U2 * Dy))
     If (U1 > 0) Then PX1 = (PX1 + (U1 * Dx)): PY1 = (PY1 + (U1 * Dy))
     OutX1 = PX1: OutY1 = PY1: OutX2 = PX2: OutY2 = PY2
    End If
   End If
  End If
 End If

End Sub
Function Distance(X1!, Y1!, X2!, Y2!) As Single

 'Compute the distance between any 2D vectors

 Distance = (((X1 - X2) ^ 2) + ((Y1 - Y2) ^ 2)) ^ 0.5

End Function
Function IsInCircle(CX!, CY!, Ray!, PX!, PY!) As Boolean

 'A simple point/circle intersection, based on distances

 If Distance(CX, CY, PX, PY) < Ray Then IsInCircle = True

End Function
Private Sub Command1_Click()

 'Some serval constants:
 Const Ratio45! = 1.414214  '1/COS(45°)
 Const Ray! = 180           '(The circle ray, can change)

 Command1.Enabled = False

 Dim CX!, CY!, X!, Y!, t!, U!, V!, Dist!
 Dim X1!, Y1!, OX1!, OY1!, OX2!, OY2!

 'Define the center in screen space:
 CX = (D.ScaleWidth * 0.5)
 CY = (D.ScaleHeight * 0.5)

 'Scan conversion
 For Y = 0 To (D.ScaleHeight - 1)
  For X = 0 To (D.ScaleWidth - 1)

   If IsInCircle(CX, CY, Ray, X, Y) = True Then

    X1 = (X - CX): Y1 = (Y - CY)
    Dist = 1 / (((X1 * X1) + (Y1 * Y1)) ^ 0.5)
    X1 = (X1 * Dist): X1 = X1 * (Ray * Ratio45)
    Y1 = (Y1 * Dist): Y1 = Y1 * (Ray * Ratio45)
    X1 = (X1 + CX): Y1 = (Y1 + CY)

    ClipLine (CX - Ray), (CY - Ray), (CX + Ray), (CY + Ray), CX, CY, X1, Y1, OX1, OY1, OX2, OY2

    'Find the linearly t:
    t = (Distance(CX, CY, X, Y) / Ray)

    X1 = CX + ((OX2 - CX) * t)
    Y1 = CY + ((OY2 - CY) * t)

    'Find U&V coordinates in texture space (range between 0...1)
    U = (Ray + (X1 - CX)) / (Ray * 2)
    V = (Ray + (Y1 - CY)) / (Ray * 2)

    'And scale these parametric coordinates
    ' by texture scales to find the texel coords,
    '  then you can choose a filtering methode:

    'Nearest nightbor:
    D.PSet (X, Y), S.Point(U * S.ScaleWidth, V * S.ScaleHeight)

    'Bilinear 1
    'D.PSet (X, Y), Bilinear(S, U * S.ScaleWidth, V * S.ScaleHeight)

    'Bilinear 2
    'D.PSet (X, Y), Bilinear2(S, U * S.ScaleWidth, V * S.ScaleHeight, 1)

    'Bell
    'D.PSet (X, Y), Bell(S, U * S.ScaleWidth, V * S.ScaleHeight)

    'Gaussian
    'D.PSet (X, Y), Gaussian(S, U * S.ScaleWidth, V * S.ScaleHeight, 1)

    'Bicubic B spline
    'D.PSet (X, Y), BicubicBSpline(S, U * S.ScaleWidth, V * S.ScaleHeight)

    'Bicubic BC spline
    'D.PSet (X, Y), BicubicBCSpline(S, U * S.ScaleWidth, V * S.ScaleHeight, 0.5, 0.5)

    'Bicubic cardinal spline
    'D.PSet (X, Y), BicubicCardinal(S, U * S.ScaleWidth, V * S.ScaleHeight, 0.8)

   End If

  Next X
  DoEvents
 Next Y

 Command1.Enabled = True

End Sub

Private Sub Form_Load()

 MsgBox "This program work perfectly with squared textures !", vbInformation, "Info"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

 Dim StrMsg As String

 Unload Me

 StrMsg = "Spherical mapping, Pure VB ! & FREE, but give me a little credit !" & vbNewLine & vbNewLine & _
          "                       KACI Lounes - Mai 2006"

 MsgBox StrMsg, vbInformation, "By !"

 End

End Sub
