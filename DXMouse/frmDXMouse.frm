VERSION 5.00
Begin VB.Form frmDXMouse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DXMouse "
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2265
   ClipControls    =   0   'False
   Icon            =   "frmDXMouse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   2265
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Left            =   1560
      Top             =   1320
   End
   Begin VB.Label Label5 
      Caption         =   "Button 3"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Button 2"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Button 1"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Button 0"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Z"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Y"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "X"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmDXMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code was made by Sauron (sauron@mtdoom.cjb.net)
'If you like this code please vote for it on
'www.planetsourcecode.com

'Thanks to Jack Hoxley for the tutorial on how to use
'DirectInput with the keyboard because it showed me how to
'use Direct Input. If I hadn't seen it I wouldn't have
'knowen how to make this.

'This is the main DirectX Object which you must create or
'you will not be able to use any other DirectX Objects
Dim DX7 As New DirectX7

'This is the main direct Input object. It is used to create
'other Direct Input objects and a few other functions, but
'in this tutorial we are just going to use it to create
'a direct input device.
Dim DI As DirectInput

'This is the Direct Input Device. This is used to
'control and get information from a input device which
'in this case is the mouse.
Dim DIDev As DirectInputDevice

'This is the Direct Input Mouse State. It is used to hold
'information on the mouse such as what keys are down or
'how far the mouse has moved along the X and Y axis since
'it was last check.
Dim DIState As DIMOUSESTATE

'This is the main code that is needed to setup DirectInput
Private Sub Form_Load()

'Making the direct input device. This must be done before
'you can use any direct input stuff.
Set DI = DX7.DirectInputCreate

'Creating the DirectInputDevice which in this case is a
'mouse. You must set the Guide (the string) to GUID_SysMouse
'or it won't work (unless you are using a different device
'which you would use its Guide instead).
Set DIDev = DI.CreateDevice("GUID_SysMouse")

'This is telling Direct Input to use the mouse communication
'method when dealing with the device.
Call DIDev.SetCommonDataFormat(DIFORMAT_MOUSE)

'Setting cooperativelevel to the window and alowing other
'programs to use the mouse too. (if you want to stop this
'just change the last parameter to DISCL_EXCLUSIVE
Call DIDev.SetCooperativeLevel(Me.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE)

'This is need to tell directx to start controlling the
'device.
DIDev.Acquire

'Setting timer which is used to check the mouse
Timer1.Interval = 1
Timer1.Enabled = True

End Sub

'Code to be run when closing direct input
Private Sub Form_Unload(Cancel As Integer)

'This tells directx to stop using the mouse.
DIDev.Unacquire
End Sub

'The timer is used to check the mouse
Private Sub Timer1_Timer()

'This is done to let windows do things it needs too. If you
'don't have this the program may crash. (its always good
'practice to include this).
DoEvents

'Refreshing the mouse status
DIDev.GetDeviceStateMouse DIState

'Drawing changes in position to the textboxs
Text1(0).Text = DIState.x
Text1(1).Text = DIState.y
Text1(2).Text = DIState.z

'Checking if any mouse buttons have been pressed
'If they have been pressed the value is not 0.
'I have made it that when a button is down the colour of
'the label that represents it will change.

If DIState.buttons(0) <> 0 Then Label2.ForeColor = RGB(255, 0, 0) Else Label2.ForeColor = RGB(0, 0, 255)
If DIState.buttons(1) <> 0 Then Label3.ForeColor = RGB(255, 0, 0) Else Label3.ForeColor = RGB(0, 0, 255)
If DIState.buttons(2) <> 0 Then Label4.ForeColor = RGB(255, 0, 0) Else Label4.ForeColor = RGB(0, 0, 255)
If DIState.buttons(3) <> 0 Then Label5.ForeColor = RGB(255, 0, 0) Else Label5.ForeColor = RGB(0, 0, 255)

End Sub
