VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "End"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   2880
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "DELEGATOR FOR DIRECT TIMER V2 (USING INTERFACES) BY VANJA FUCKAR,EMAIL:INGA@VIP.HR"
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
Private Declare Function SetTimer& Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
Private Declare Function KillTimer& Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long)



Private TIS As TimerInterface 'declare Interface
Implements TimerInterface 'and Implement interface to form


'**********************REQUIRED BY DELEGATON PER METHOD!************************
Private ObjectProcedures() As FunctionSPointerS '<------Class Procedure pointers and addresses!
Private HMEM As Long   '<--------------Memory Handle
Private TIMERPROC As Long  '<---------AddressOf Delegator (who call our Class Method!)
Private HTIMER As Long '<--------Timer Handle
'*******************************************************************************


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
'******************
'REQUIRED! Set Interface Procedure Pointer on TIS to Our Form Implemented Interface!!!
Set TIS = Me
'******************


'REQUIRED BY DELEGATION PER METHOD!**********************************************
ObjectProcedures = GetObjectFunctionsPointers(TIS, 1)
HMEM = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, CalculateSpaceForDelegation(4))
TIMERPROC = GlobalLock(HMEM) 'Get Address
DelegateFunction TIMERPROC, TIS, ObjectProcedures(0).FunctionAddress, 4
'USE DELEGATION For 1st Method In SubclassingInterface Class!!! -->ObjectProcedures(0)
'********************************************************************************


'Set Timer!
HTIMER = SetTimer(Me.hwnd, 0, 0, TIMERPROC) '<<---Call ASM Code!

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Kill Timer!
KillTimer Me.hwnd, HTIMER

'Free Memory Allocated For Delegation
Call GlobalUnlock(HMEM)
Call GlobalFree(HMEM)
End Sub


Private Sub TimerInterface_TIMERPROCEDURE(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
Debug.Print "TIMER ON HWND:" & hwnd & ",MSG:" & uMsg & ",IDEVENT:" & idEvent & ",DWTIME:" & dwTime
End Sub
