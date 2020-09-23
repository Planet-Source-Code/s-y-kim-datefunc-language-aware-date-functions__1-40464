VERSION 5.00
Begin VB.Form frmDatesFunc 
   Caption         =   "Language-Aware DateFunc Test"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   9540
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin VB.Frame fraCtrl 
      BorderStyle     =   0  '¾øÀ½
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   4260
      Width           =   9195
      Begin VB.ComboBox cboLang 
         Height          =   300
         Left            =   1020
         TabIndex        =   4
         Text            =   "cboLang"
         Top             =   60
         Width           =   2295
      End
      Begin VB.CommandButton cmdGetInfo 
         Caption         =   "Get Date Func Infos"
         Height          =   315
         Left            =   6960
         TabIndex        =   3
         Top             =   0
         Width           =   2175
      End
      Begin VB.TextBox txtDateTime 
         Height          =   270
         Left            =   4620
         TabIndex        =   2
         Text            =   "txtDateTime"
         Top             =   60
         Width           =   2235
      End
      Begin VB.Label lblLang 
         Caption         =   "Language:"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   60
         Width           =   2295
      End
      Begin VB.Label lblDate 
         Caption         =   "Date/Time:"
         Height          =   195
         Left            =   3480
         TabIndex        =   5
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.ListBox lstDateInfo 
      Height          =   3660
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   9255
   End
End
Attribute VB_Name = "frmDatesFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessageArray Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Function SetTapStops(lstHwnd As Long, Tabs() As Long) As Long

   Const LB_SETTABSTOPS = &H192  'set the tab-stop positions
   'Private Declare Function SendMessageArray Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
   
   'Clear any existing tabs.
   Call SendMessageArray(lstHwnd, LB_SETTABSTOPS, 0&, 0&)
   'Set the tabs.
   Call SendMessageArray(lstHwnd, LB_SETTABSTOPS, UBound(Tabs) + 1, Tabs(0))

End Function

Private Sub cmdGetInfo_Click()
   
   Dim i As eDateFunc
   With Me.lstDateInfo
      .Clear
      For i = dtDateFunc_MIN To dtDateFunc_MAX
         lstDateInfo.AddItem eDateFuncDesc(i) & vbTab & DateFunc(i, CDate(txtDateTime.Text), , , , cboLang.ListIndex)
      Next i
      .AddItem DateMatchDesc(2002, vbNovember, vbTuesday, vbFourth, , cboLang.ListIndex) & vbTab & _
                     DateMatch(2002, vbNovember, vbTuesday, vbFourth)
   End With

End Sub

Private Sub Form_DblClick()
   Dim frm As frmDatesFunc
   Set frm = New frmDatesFunc
   frm.Show
End Sub

Private Sub Form_Load()

   Dim i As eLanguage
   Dim lstTabs(0) As Long
   
   
   Me.cboLang.Clear
   For i = eLanguage_MIN To eLanguage_MAX
      With Me.cboLang
         'If IsIneLanguage(i) Then
            .AddItem eLanguageDesc(i)
         'End If
      End With
   Next i
   
   lstTabs(0) = 150
   Call SetTapStops(Me.lstDateInfo.HWnd, lstTabs)
   Me.cboLang.ListIndex = eLanguage.English
   'Me.txtDateTime.Text = DateFunc(dtShortDateLongTime, Now(), , , , English)
   Me.txtDateTime.Text = Now()
   Debug.Print Now()
   Debug.Print DateFunc(dtShortDateLongTime, Now(), , , , English)
   
End Sub


Private Sub Form_Resize()
   On Error Resume Next
   If Me.Width < fraCtrl.Width Then
      Me.Width = fraCtrl.Width + 150
   End If
   Me.fraCtrl.Move 50, Me.ScaleHeight - fraCtrl.Height
   Me.lstDateInfo.Move 50, 50, Me.ScaleWidth - 50, fraCtrl.Top - 50
   On Error GoTo 0
End Sub
