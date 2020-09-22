VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   480
      ScaleHeight     =   2355
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_X As Integer
Public m_Y As Integer
Public m_Len As Integer
Public R As Long
Public G As Long
Public B As Long

Private Sub Form_Load()
m_X = 1500
m_Y = 1000
m_Len = 500
Timer1_Timer
End Sub

Private Sub Timer1_Timer()
If m_Len = 0 Then
Timer1.Enabled = False
Else
Picture1.Line (m_X, m_Y)-(m_X + m_Len, m_Y), RGB(R, G, B)
Picture1.Line (m_X, m_Y)-(m_X, m_Y + m_Len), RGB(R, G, B)
Picture1.Line (m_X, m_Y + m_Len)-(m_X + m_Len, m_Y + m_Len), RGB(R, G, B)
Picture1.Line (m_X + m_Len, m_Y)-(m_X + m_Len, m_Y + m_Len), RGB(R, G, B)

m_X = m_X + 1
m_Y = m_Y + 1
m_Len = m_Len - 2

R = R + 1
G = G + 1
B = B + 1
End If
End Sub
