VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   675
   ClientLeft      =   6165
   ClientTop       =   3855
   ClientWidth     =   2025
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   2025
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "Keep Transparent"
      Height          =   255
      Left            =   20
      TabIndex        =   2
      Top             =   410
      Width           =   1680
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2280
      Top             =   0
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2760
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Color           =   1
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Make Transparent"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdCol 
      BackColor       =   &H000000C0&
      Height          =   375
      Left            =   1640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COLOR_BACKGROUND = 1
Private Const LVM_FIRST = &H1000 ' ListView messages
Private Const LVM_GETTEXTCOLOR = (LVM_FIRST + 35)
Private Const LVM_SETTEXTCOLOR = (LVM_FIRST + 36)
Private Const LVM_GETTEXTBKCOLOR = (LVM_FIRST + 37)
Private Const LVM_SETTEXTBKCOLOR = (LVM_FIRST + 38)

Private Const CLR_NONE = &HFFFFFFFF

Private Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
    
Private Declare Function FindWindowEx Lib "user32" _
    Alias "FindWindowExA" (ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long
    
Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    
Private Declare Function GetSysColor Lib "user32" _
    (ByVal nIndex%) As Long
    
Private Declare Function InvalidateRect Lib "user32" _
    (ByVal hwnd As Long, lpRect As Any, _
    ByVal bErase As Long) As Long
    
Private Declare Function UpdateWindow Lib "user32" _
    (ByVal hwnd As Long) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpINIPath As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpString As Any, ByVal lpINIPath As String) As Long

Dim lProgman         As Long
Dim lSHELLDLLDefView As Long
Dim lSysListView32   As Long
Dim INIfPath         As String

Private Sub Check1_Click()
If Check1 Then
  Timer1.Enabled = True
Else
  Timer1.Enabled = False
End If
End Sub

Private Sub cmdChange_Click()
Dim bRet             As Boolean
' Get the current background color. If it is not CLR_NONE
' (no background color) set it so. If it is, set the current
' background color.
'
If (ListView_GetTextBkColor(lSysListView32) <> CLR_NONE) Then
    bRet = ListView_SetTextBkColor(lSysListView32, CLR_NONE)
Else
    Call ListView_SetTextBkColor(lSysListView32, GetSysColor(COLOR_BACKGROUND))
End If
'
' Add a rectangle to the listview's update region. This is the portion of
' the window's client area that must be redrawn. The 0 parameters tells
' it to redraw the entire client area.
'
Call InvalidateRect(lSysListView32, ByVal 0&, True)
'
' Send a WM_PAINT message to the listview to force
' it to redraw itself.
'
Call UpdateWindow(lSysListView32)

If bRet Then
    cmdChange.Caption = "Make Coloured"
Else
    cmdChange.Caption = "Make Transparent"
End If

End Sub

Private Sub cmdCol_Click()
 CD1.CancelError = False
 CD1.ShowColor
 cmdCol.BackColor = CD1.Color
 cmdChange.SetFocus
 Call ListView_SetTextColor(lSysListView32, cmdCol.BackColor)
 Call InvalidateRect(lSysListView32, ByVal 0&, True)
 Call UpdateWindow(lSysListView32)
End Sub

Private Sub Form_Load()
Dim tmpCol As String

If Right$(App.Path, 1) <> "\" Then
 INIfPath = App.Path & "\tdi.ini"
Else
 INIfPath = App.Path & "tdi.ini"
End If

If UCase$(Command$) = "-ON-H" Then
 Me.Visible = False
 Check1.Value = 1 ' fires the Check1_Click sub
End If

tmpCol = ReadINI("TXTCOL", "COL", INIfPath)
If tmpCol <> "" Then cmdCol.BackColor = CLng(tmpCol)

' Get the handle to the top level window with a class name of
' "Progman" and a caption of "Program Manager".
'
lProgman = FindWindow("Progman", "Program Manager")
If lProgman = 0 Then Exit Sub
'
' Get Program Manager's child window which has
' a class name of "SHELLDLL_DefView".
'
lSHELLDLLDefView = FindWindowEx(lProgman, 0&, "SHELLDLL_DefView", vbNullString)
If lSHELLDLLDefView = 0 Then Exit Sub
'
' Now get this window's child.
'
lSysListView32 = FindWindowEx(lSHELLDLLDefView, 0&, "SysListView32", vbNullString)
If lSysListView32 = 0 Then Exit Sub
'

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim OkY As Integer

OkY = WriteINI("TXTCOL", "COL", CStr(cmdCol.BackColor), INIfPath)
If OkY = 0 Then MsgBox "Error writing to INI file : settings not saved", vbCritical, "doh"

End Sub

Private Function ListView_SetTextBkColor(hwnd As Long, clrTextBk As Long) As Boolean
Dim lRet As Long

lRet = SendMessage((hwnd), LVM_SETTEXTBKCOLOR, 0&, clrTextBk)

If lRet = 0 Then
    ListView_SetTextBkColor = False
Else
    ListView_SetTextBkColor = True
End If
End Function

Private Function ListView_GetTextBkColor(hwnd As Long) As Long
 ListView_GetTextBkColor = SendMessage((hwnd), LVM_GETTEXTBKCOLOR, 0, 0)
End Function

Private Function ListView_SetTextColor(hwnd As Long, Colour As Long)
Dim lRet As Long
 
 lRet = SendMessage((hwnd), LVM_SETTEXTCOLOR, 0, Colour)
End Function

Private Sub Timer1_Timer()
If lSysListView32 <> 0 Then
 If (ListView_GetTextBkColor(lSysListView32) <> CLR_NONE) Then
  Call ListView_SetTextBkColor(lSysListView32, CLR_NONE)
  Call ListView_SetTextColor(lSysListView32, cmdCol.BackColor)
  Call InvalidateRect(lSysListView32, ByVal 0&, True)
  Call UpdateWindow(lSysListView32)
 End If
Else
 Form_Load
End If
End Sub

Private Function ReadINI(ByVal Section As String, ByVal Key As String, ByVal INIPath As String) As String
Dim RetStr As String

If Dir(INIPath) = "" Then Exit Function
RetStr = String(255, Chr(0))
ReadINI = Left(RetStr, GetPrivateProfileString(Section, ByVal Key, "", RetStr, Len(RetStr), INIPath))
End Function

Private Function WriteINI(ByVal Section As String, ByVal Key As String, ByVal KeyValue As String, ByVal INIPath As String) As Integer
'Function returns 1 if successful and 0 if unsuccessful

If Dir(INIPath) = "" Then Exit Function
WritePrivateProfileString Section, Key, KeyValue, INIPath
WriteINI = 1
End Function
