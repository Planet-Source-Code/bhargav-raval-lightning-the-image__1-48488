VERSION 5.00
Begin VB.Form frmLigntning 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lightning the Picture"
   ClientHeight    =   6915
   ClientLeft      =   300
   ClientTop       =   1380
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   10440
   Begin VB.PictureBox Picture1 
      Height          =   6255
      Left            =   120
      Picture         =   "frmLigntning.frx":0000
      ScaleHeight     =   413
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   333
      TabIndex        =   3
      Top             =   0
      Width           =   5055
   End
   Begin VB.PictureBox Picture2 
      Height          =   6255
      Left            =   5280
      ScaleHeight     =   413
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   333
      TabIndex        =   2
      Top             =   0
      Width           =   5055
   End
   Begin VB.TextBox txtVal 
      Height          =   285
      Left            =   4200
      TabIndex        =   1
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Re&lode"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   6480
      Width           =   2175
   End
End
Attribute VB_Name = "frmLigntning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intUpperboundX As Integer
Dim intUpperboundY As Integer
Dim Pixels()
Private Sub Command1_Click()
On Error Resume Next
    Dim X, Y, intAddOn As Integer
    Dim R As Integer, G As Integer, B As Integer
    intUpperboundX = Picture1.ScaleWidth
    intUpperboundY = Picture1.ScaleHeight
    ReDim Pixels(1 To intUpperboundX, 1 To intUpperboundY)
    intAddOn = Val(txtVal)
    For X = 1 To intUpperboundX
        For Y = 1 To intUpperboundY
            Pixels(X, Y) = Picture1.Point(X, Y)
        Next Y
    Next X
    For X = 1 To intUpperboundX
        For Y = 1 To intUpperboundY
            R = Pixels(X, Y) And &HFF
            G = ((Pixels(X, Y) And &HFF00) / &H100) Mod &H100
            B = ((Pixels(X, Y) And &HFF0000) / &H10000) Mod &H100
            R = R + intAddOn
            If R > 255 Then R = 255
            G = G + intAddOn
            If G > 255 Then G = 255
            B = B + intAddOn
            If B > 255 Then B = 255
            Pixels(X, Y) = RGB(R, G, B)
        Next Y
    Next X
    For X = 1 To intUpperboundX
        For Y = 1 To intUpperboundY
            Picture2.PSet (X, Y), Pixels(X, Y)
        Next Y
    Next X
End Sub

