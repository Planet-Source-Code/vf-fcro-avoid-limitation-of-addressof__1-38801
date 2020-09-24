VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Subclassed And Timer 4"
      Height          =   495
      Left            =   6600
      TabIndex        =   5
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Subclassed And Timer 3"
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Subclassed And Timer 2"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Subclassed And Timer 1"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton CommandX 
      Caption         =   "End"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   3480
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "MULTI DELEGATOR FOR DIRECT SUBCLASSING AND TIMERS V2 (USING INTERFACES) BY VANJA FUCKAR,EMAIL:INGA@VIP.HR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetTimer& Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
Private Declare Function KillTimer& Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long)


Private OLDPROCEDURE1 As Long
Private OLDPROCEDURE2 As Long
Private OLDPROCEDURE3 As Long
Private OLDPROCEDURE4 As Long

Private IMSI As MultiSubTmr  'declare Interface
Implements MultiSubTmr  'and Implement interface to form


'**********************REQUIRED BY DELEGATON PER METHOD!************************
Private ObjectProcedures() As FunctionSPointerS '<------Class Procedure pointers and addresses!
'*******************************************************************************

'Subclassing
Private HMEM1 As Long   '<--------------Memory Handle1
Private HWNDPROC1 As Long  '<---------AddressOf Delegator1

Private HMEM2 As Long   '<--------------Memory Handle2
Private HWNDPROC2 As Long  '<---------AddressOf Delegator2

Private HMEM3 As Long   '<--------------Memory Handle3
Private HWNDPROC3 As Long  '<---------AddressOf Delegator3

Private HMEM4 As Long   '<--------------Memory Handle4
Private HWNDPROC4 As Long  '<---------AddressOf Delegator4

'Timers
Private HMEM5 As Long   '<--------------Memory Handle
Private TIMERPROC1 As Long  '<---------AddressOf Delegator
Private HTIMER1 As Long '<------Timer Handle

Private HMEM6 As Long   '<--------------Memory Handle
Private TIMERPROC2 As Long  '<---------AddressOf Delegator
Private HTIMER2 As Long '<------Timer Handle

Private HMEM7 As Long   '<--------------Memory Handle
Private TIMERPROC3 As Long  '<---------AddressOf Delegator
Private HTIMER3 As Long '<------Timer Handle

Private HMEM8 As Long   '<--------------Memory Handle
Private TIMERPROC4 As Long  '<---------AddressOf Delegator
Private HTIMER4 As Long '<------Timer Handle





Private Sub CommandX_Click()
Unload Me
End Sub

Private Sub Form_Load()
'******************
'REQUIRED! Set Interface Procedure Pointer on IMSI to Our Form Implemented Interface!!!
Set IMSI = Me
'******************

ObjectProcedures = GetObjectFunctionsPointers(IMSI, 8) '<----Get Pointers For 8 Methods

'DELGATE 8 METHODS FOR SUBCLASSING & TIMER PROCEDURES


'Delegate Subclass Method (1st)
HMEM1 = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, CalculateSpaceForDelegation(4))
HWNDPROC1 = GlobalLock(HMEM1) 'Get Address
DelegateFunction HWNDPROC1, IMSI, ObjectProcedures(0).FunctionAddress, 4

'Delegate Subclass Method (2nd)
HMEM2 = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, CalculateSpaceForDelegation(4))
HWNDPROC2 = GlobalLock(HMEM2) 'Get Address
DelegateFunction HWNDPROC2, IMSI, ObjectProcedures(1).FunctionAddress, 4

'Delegate Subclass Method (3rd)
HMEM3 = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, CalculateSpaceForDelegation(4))
HWNDPROC3 = GlobalLock(HMEM3) 'Get Address
DelegateFunction HWNDPROC3, IMSI, ObjectProcedures(2).FunctionAddress, 4

'Delegate Subclass Method (4th)
HMEM4 = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, CalculateSpaceForDelegation(4))
HWNDPROC4 = GlobalLock(HMEM4) 'Get Address
DelegateFunction HWNDPROC4, IMSI, ObjectProcedures(3).FunctionAddress, 4


'Delegate Timer Method (5th)
HMEM5 = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, CalculateSpaceForDelegation(4))
TIMERPROC1 = GlobalLock(HMEM5) 'Get Address
DelegateFunction TIMERPROC1, IMSI, ObjectProcedures(4).FunctionAddress, 4

'Delegate Timer Method (6th)
HMEM6 = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, CalculateSpaceForDelegation(4))
TIMERPROC2 = GlobalLock(HMEM6) 'Get Address
DelegateFunction TIMERPROC2, IMSI, ObjectProcedures(5).FunctionAddress, 4

'Delegate Timer Method (7th)
HMEM7 = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, CalculateSpaceForDelegation(4))
TIMERPROC3 = GlobalLock(HMEM7) 'Get Address
DelegateFunction TIMERPROC3, IMSI, ObjectProcedures(6).FunctionAddress, 4

'Delegate Timer Method (8th)
HMEM8 = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, CalculateSpaceForDelegation(4))
TIMERPROC4 = GlobalLock(HMEM8) 'Get Address
DelegateFunction TIMERPROC4, IMSI, ObjectProcedures(7).FunctionAddress, 4

'**********************************************************************************





'LET's GO Subclassing!
OLDPROCEDURE1 = SetWindowLong(Command1.hwnd, -4, HWNDPROC1)
OLDPROCEDURE2 = SetWindowLong(Command2.hwnd, -4, HWNDPROC2)
OLDPROCEDURE3 = SetWindowLong(Command3.hwnd, -4, HWNDPROC3)
OLDPROCEDURE4 = SetWindowLong(Command4.hwnd, -4, HWNDPROC4)

'Set Timers!
HTIMER1 = SetTimer(Command1.hwnd, 0, 0, TIMERPROC1)
HTIMER2 = SetTimer(Command2.hwnd, 0, 0, TIMERPROC2)
HTIMER3 = SetTimer(Command3.hwnd, 0, 0, TIMERPROC3)
HTIMER4 = SetTimer(Command4.hwnd, 0, 0, TIMERPROC4)


End Sub

Private Sub Form_Unload(Cancel As Integer)
'Unsubclass!
SetWindowLong Command1.hwnd, -4, OLDPROCEDURE1
SetWindowLong Command2.hwnd, -4, OLDPROCEDURE2
SetWindowLong Command3.hwnd, -4, OLDPROCEDURE3
SetWindowLong Command4.hwnd, -4, OLDPROCEDURE4

'Kill Timers!
KillTimer Command1.hwnd, HTIMER1
KillTimer Command2.hwnd, HTIMER2
KillTimer Command3.hwnd, HTIMER3
KillTimer Command4.hwnd, HTIMER4

'Free Memory Allocated For Delegation
Call GlobalUnlock(HMEM1)
Call GlobalFree(HMEM1)

Call GlobalUnlock(HMEM2)
Call GlobalFree(HMEM2)

Call GlobalUnlock(HMEM3)
Call GlobalFree(HMEM3)

Call GlobalUnlock(HMEM4)
Call GlobalFree(HMEM4)

Call GlobalUnlock(HMEM5)
Call GlobalFree(HMEM5)

Call GlobalUnlock(HMEM6)
Call GlobalFree(HMEM6)

Call GlobalUnlock(HMEM7)
Call GlobalFree(HMEM7)

Call GlobalUnlock(HMEM8)
Call GlobalFree(HMEM8)
End Sub


Private Sub MultiSubTmr_TIMERPROCEDURE(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
Debug.Print "TIMER ON HWND:" & hwnd & ",MSG:" & uMsg & ",IDEVENT:" & idEvent & ",DWTIME:" & dwTime
End Sub

Private Sub MultiSubTmr_TIMERPROCEDURE2(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
Debug.Print "TIMER ON HWND:" & hwnd & ",MSG:" & uMsg & ",IDEVENT:" & idEvent & ",DWTIME:" & dwTime
End Sub

Private Sub MultiSubTmr_TIMERPROCEDURE3(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
Debug.Print "TIMER ON HWND:" & hwnd & ",MSG:" & uMsg & ",IDEVENT:" & idEvent & ",DWTIME:" & dwTime
End Sub

Private Sub MultiSubTmr_TIMERPROCEDURE4(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
Debug.Print "TIMER ON HWND:" & hwnd & ",MSG:" & uMsg & ",IDEVENT:" & idEvent & ",DWTIME:" & dwTime
End Sub

Private Function MultiSubTmr_WndProc1(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Debug.Print "HWND:" & hwnd & ",MSG:" & uMsg & ",WPARAM:" & wParam & ",LPARAM:" & lParam

MultiSubTmr_WndProc1 = CallWindowProc(OLDPROCEDURE1, hwnd, uMsg, wParam, lParam)

End Function

Private Function MultiSubTmr_WndProc2(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Debug.Print "HWND:" & hwnd & ",MSG:" & uMsg & ",WPARAM:" & wParam & ",LPARAM:" & lParam

MultiSubTmr_WndProc2 = CallWindowProc(OLDPROCEDURE2, hwnd, uMsg, wParam, lParam)

End Function

Private Function MultiSubTmr_WndProc3(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Debug.Print "HWND:" & hwnd & ",MSG:" & uMsg & ",WPARAM:" & wParam & ",LPARAM:" & lParam

MultiSubTmr_WndProc3 = CallWindowProc(OLDPROCEDURE3, hwnd, uMsg, wParam, lParam)

End Function

Private Function MultiSubTmr_WndProc4(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Debug.Print "HWND:" & hwnd & ",MSG:" & uMsg & ",WPARAM:" & wParam & ",LPARAM:" & lParam

MultiSubTmr_WndProc4 = CallWindowProc(OLDPROCEDURE4, hwnd, uMsg, wParam, lParam)

End Function
