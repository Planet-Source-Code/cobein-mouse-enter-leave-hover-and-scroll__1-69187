VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Index           =   1
      Left            =   1260
      ScaleHeight     =   2475
      ScaleWidth      =   3015
      TabIndex        =   1
      Top             =   300
      Width           =   3075
      Begin VB.PictureBox Picture1 
         Height          =   915
         Index           =   2
         Left            =   1320
         ScaleHeight     =   855
         ScaleWidth      =   975
         TabIndex        =   2
         Top             =   780
         Width           =   1035
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   915
      Index           =   0
      Left            =   240
      ScaleHeight     =   855
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   300
      Width           =   1035
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents f_cMO As cMouseOver
Attribute f_cMO.VB_VarHelpID = -1

' The mouse enter event can be configured to be fired X millisecs after the mouse
' move event, in that way if you are not targetting the control... just playing arround
' the control will not change the state.
Private Sub f_cMO_MouseEnter(ByVal lhWnd As Long, ByVal vExtra As Variant)
    Dim i As Long
    For i = 0 To 2
        If Picture1(i).hwnd = lhWnd Then
            Picture1(i).BackColor = vbBlue
            Exit Sub
        End If
    Next
End Sub

Private Sub f_cMO_MouseHover(ByVal lhWnd As Long, ByVal vExtra As Variant)
    
    Dim i As Long
    For i = 0 To 2
        If Picture1(i).hwnd = lhWnd Then
            Picture1(i).BackColor = vbRed
            Exit Sub
        End If
    Next
End Sub

Private Sub f_cMO_MouseLeave(ByVal lhWnd As Long, ByVal vExtra As Variant)
    Dim i As Long
    For i = 0 To 2
        If Picture1(i).hwnd = lhWnd Then
            Picture1(i).BackColor = vbButtonFace
            Exit Sub
        End If
    Next
End Sub

Private Sub f_cMO_MouseScroll(ByVal lhWnd As Long, ByVal vExtra As Variant, ByVal lLines As Long)
    Debug.Print lhWnd, lLines
End Sub

Private Sub Form_Load()
    Set f_cMO = New cMouseOver
    Dim i As Long
    For i = 0 To 2
        f_cMO.AttachObject Picture1(i).hwnd 'you have some optional parameters
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set f_cMO = Nothing
End Sub
