VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9540
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14820
   LinkTopic       =   "Form1"
   ScaleHeight     =   9540
   ScaleWidth      =   14820
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrStep 
      Interval        =   20
      Left            =   6840
      Top             =   4560
   End
   Begin VB.Shape shpChase 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mousex As Single
Dim mousey As Single

Private Sub Form_Load()
mousex = 1
mousey = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mousex = X
mousey = Y
End Sub

Private Sub tmrStep_Timer()
Dim diffX As Single
Dim diffY As Single
    If (shpChase.Left + (shpChase.Width / 2)) <> (mousex) Then
        diffX = (shpChase.Left + (shpChase.Width / 2)) Mod 10
        
        If (mousex) - (shpChase.Left + (shpChase.Width / 2)) > 0 Then
            shpChase.Left = shpChase.Left + 10 * diffX
        Else
            shpChase.Left = shpChase.Left - 10 * diffX
        End If
    ElseIf (shpChase.Left + (shpChase.Width / 2)) <> Int(mousex) Then
        shpChase.BackColor = vbRed
    End If
    If (shpChase.Top + (shpChase.Height / 2)) <> (mousey) Then
        diffY = (shpChase.Top + (shpChase.Height / 2)) Mod 10
        If (mousey) - (shpChase.Top + (shpChase.Height / 2)) > 0 Then
            shpChase.Top = shpChase.Top + 10 * diffY
        Else
            shpChase.Top = shpChase.Top - 10 * diffY
        End If
        shpChase.BackColor = vbBlue
    ElseIf Int(shpChase.Top + (shpChase.Height / 2)) <> Int(mousey) Then
        shpChase.BackColor = vbRed
    End If
End Sub
