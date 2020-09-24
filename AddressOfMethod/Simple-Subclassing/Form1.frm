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
      Caption         =   "DELEGATOR FOR DIRECT SUBCLASSING V2 (USING INTERFACES) BY VANJA FUCKAR,EMAIL:INGA@VIP.HR"
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
Private OLDPROCEDURE As Long


Private IMSI As SubclassingInterface 'declare Interface
Implements SubclassingInterface 'and Implement interface to form


'**********************REQUIRED BY DELEGATON PER METHOD!************************
Private ObjectProcedures() As FunctionSPointerS '<------Class Procedure pointers and addresses!
Private HMEM As Long   '<--------------Memory Handle
Private HWNDPROC As Long  '<---------AddressOf Delegator (who call our Class Method!)
'*******************************************************************************


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
'******************
'REQUIRED! Set Interface Procedure Pointer on IMSI to Our Form Implemented Interface!!!
Set IMSI = Me
'******************


'REQUIRED BY DELEGATION PER METHOD!**********************************************
ObjectProcedures = GetObjectFunctionsPointers(IMSI, 1)
HMEM = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, CalculateSpaceForDelegation(4))
HWNDPROC = GlobalLock(HMEM) 'Get Address
DelegateFunction HWNDPROC, IMSI, ObjectProcedures(0).FunctionAddress, 4
'USE DELEGATION For 1st Method In SubclassingInterface Class!!! -->ObjectProcedures(0)
'********************************************************************************


'LET's GO Subclassing!
OLDPROCEDURE = SetWindowLong(Me.hwnd, -4, HWNDPROC) '<<---Call ASM Code!

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Unsubclass!
SetWindowLong Me.hwnd, -4, OLDPROCEDURE

'Free Memory Allocated For Delegation
Call GlobalUnlock(HMEM)
Call GlobalFree(HMEM)
End Sub

Private Function SubclassingInterface_WndProc(ByVal hwnd As Long, ByVal umsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'DO NOT USE BREAKPOINT!!!!!
Debug.Print "HWND:" & hwnd & ",MSG:" & umsg & ",WPARAM:" & wParam & ",LPARAM:" & lParam

SubclassingInterface_WndProc = CallWindowProc(OLDPROCEDURE, hwnd, umsg, wParam, lParam)
End Function
