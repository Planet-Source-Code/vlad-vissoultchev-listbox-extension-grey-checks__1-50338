VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4860
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5472
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   5472
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check3 
      Caption         =   "Large font"
      Height          =   264
      Left            =   3024
      TabIndex        =   5
      Top             =   1176
      Width           =   2112
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Enabled"
      Height          =   264
      Left            =   3024
      TabIndex        =   4
      Top             =   840
      Value           =   1  'Checked
      Width           =   2196
   End
   Begin VB.ListBox List1 
      Height          =   1944
      IntegralHeight  =   0   'False
      Left            =   84
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   84
      Width           =   2868
   End
   Begin VB.TextBox Text1 
      Height          =   2532
      Left            =   84
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2184
      Width           =   5304
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Draw hollow greyed"
      Height          =   264
      Left            =   3024
      TabIndex        =   1
      Top             =   504
      Width           =   2196
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset greyed"
      Height          =   348
      Left            =   3024
      TabIndex        =   0
      Top             =   84
      Width           =   2112
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_oListBoxExt As cListBoxExt
Attribute m_oListBoxExt.VB_VarHelpID = -1

Private Sub Check1_Click()
    m_oListBoxExt.GreyedNotChecked = (Check1.Value = vbChecked)
End Sub

Private Sub Check2_Click()
    List1.Enabled = (Check2.Value = vbChecked)
End Sub

Private Sub Check3_Click()
    List1.Font.Size = 8 + Check3.Value * 16
End Sub

Private Sub Command1_Click()
    m_oListBoxExt.Selected(1) = vbGrayed
    m_oListBoxExt.Selected(3) = vbGrayed
End Sub

Private Sub Form_Load()
    List1.AddItem "aaa"
    List1.AddItem "bbb (can grey)"
    List1.AddItem "cccc"
    List1.AddItem Timer & "  (can grey)"
    List1.AddItem "test (can grey)"
    List1.AddItem "proba"
    Set m_oListBoxExt = New cListBoxExt
    m_oListBoxExt.Init List1
    m_oListBoxExt.Selected(1) = vbGrayed
    m_oListBoxExt.Selected(3) = vbGrayed
    m_oListBoxExt.CanGrey(4) = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_oListBoxExt.Terminate
End Sub

Private Sub m_oListBoxExt_ItemCheck(Item As Integer)
    Text1.Text = Text1.Text & "ItemCheck: Item =" & Item & ", Selected =" & m_oListBoxExt.Selected(Item) & vbCrLf
    Text1.SelStart = Len(Text1.Text)
End Sub
