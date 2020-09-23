VERSION 5.00
Begin VB.Form LOG_FORM 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Log Window"
   ClientHeight    =   1500
   ClientLeft      =   795
   ClientTop       =   840
   ClientWidth     =   3225
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton LOG_CLEAR 
      Caption         =   "Clear Log"
      Height          =   240
      Left            =   -15
      TabIndex        =   1
      Top             =   6630
      Width           =   4620
   End
   Begin VB.TextBox LOG_TEXT 
      Height          =   6615
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "LOG_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LOG_CLEAR_Click()
    LOG_TEXT.Text = ""
End Sub

Private Sub LOG_TEXT_Change()

End Sub
