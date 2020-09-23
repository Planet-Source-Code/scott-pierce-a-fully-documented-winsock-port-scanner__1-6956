VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Port Scanner"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClearList 
      Caption         =   "&Clear List"
      Height          =   495
      Left            =   2400
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   4350
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3651
            Text            =   "Current Port: "
            TextSave        =   "Current Port: "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "4/2/00"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "4:37 PM"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraMaxConnections 
      Caption         =   "&Max Connections"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1455
      Begin VB.TextBox txtMaxConnections 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Text            =   "1"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdStop 
      Cancel          =   -1  'True
      Caption         =   "S&top"
      Height          =   495
      Left            =   3720
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "&Scan"
      Default         =   -1  'True
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame fraOpenPorts 
      Caption         =   "Open Ports"
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   4815
      Begin VB.ListBox lstOpenPorts 
         Height          =   2400
         ItemData        =   "frmMain.frx":0000
         Left            =   120
         List            =   "frmMain.frx":0002
         TabIndex        =   7
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Timer timTimer 
      Interval        =   100
      Left            =   1920
      Top             =   1080
   End
   Begin MSWinsockLib.Winsock wskSocket 
      Index           =   0
      Left            =   1680
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraScanPorts 
      Caption         =   "Scan &Ports"
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1935
      Begin VB.TextBox txtUpperBound 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Text            =   "32676"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtLowerBound 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "1"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblTo 
         Caption         =   "To"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame fraRemoteIP 
      Caption         =   "&Remote IP"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "000.000.000.000"
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================='
' Original Source By: Scott Pierce (webmaster@calclinks.net)    '
' April 02, 2000                                                '
' See the Readme.txt file for more information                  '
'==============================================================='

Option Explicit

Private Sub cmdClearList_Click()
'This Sub is ran when user clicks the Clear button
'=================================================
   
   
   'Clear contents of lstOpenPorts
   '==============================
   Me.lstOpenPorts.Clear
   
'End cmdClearList Sub
'====================
End Sub

Private Sub cmdScan_Click()
'This Sub is ran when user clicks the Start button
'=================================================
   
   'For loop variable
   '=================
   Dim intI As Integer
   
   
   'Sets first lowerbound (first port to scan) to lngNextPort
   '=========================================================
   lngNextPort = Val(Me.txtLowerBound)
   
   'Load (txtMaxConnections) winsock controls
   '=========================================
   For intI = 1 To Val(Me.txtMaxConnections)
   
      'Load new winsock control
      '========================
      Load Me.wskSocket(intI)
      
      'Increment lngNextPort by 1
      '==========================
      lngNextPort = lngNextPort + 1
      
      'Connect new winsock control to IP address and next port
      '=======================================================
      Me.wskSocket(intI).Connect Me.txtIP, lngNextPort
   
   'Next for loop
   '=============
   Next intI
   
'End cmdScan_Click Sub
'=====================
End Sub

Private Sub cmdStop_Click()
'This Sub is ran when user clicks the Stop button
'================================================
   
   'For loop variable
   '=================
   Dim intI As Integer
   
   
   'Loop through open sockets
   '=======================
   For intI = 1 To Val(Me.txtMaxConnections)
   
      'Close current winsock connection
      '================================
      Me.wskSocket(intI).Close
      
      'Unload current winsock control
      '==============================
      Unload Me.wskSocket(intI)
   
   'Next for loop
   '=============
   Next intI
   
'End cmdStop_Click Sub
'=====================
End Sub

Private Sub timTimer_Timer()
'This Sub is ran when the timer interval is reached (default 1000ms)
'===================================================================
   
   
   'Update statusbar with next port to be scanned
   '=============================================
   Me.sbMain.Panels(1).Text = "Current Port: " + Str(lngNextPort)
   
'End timTimer_Timer Sub
'======================
End Sub

Private Sub wskSocket_Connect(Index As Integer)
'When winsock connection received this Sub is ran
'================================================
   
   
   'Add open port to lstOpenPorts
   '=============================
   Me.lstOpenPorts.AddItem "Port: " + Str(Me.wskSocket(Index).RemotePort)
   
   'Scan next available port
   '========================
   Try_Next_Port (Index)
      
'End wskSocket_Connect Sub
'=========================
End Sub

Private Sub wskSocket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'If an error is returned from current winsock control (current port not listening)
'=================================================================================
   
   
   'Scan next available port
   '========================
   Try_Next_Port (Index)
   
'End wskSocket_Error Sub
'=======================
End Sub

Private Sub Try_Next_Port(Index As Integer)
'This function will close the current winsock connection, check to see if
'upper port bound has been reached, then try a connection to the next
'available port or unload the current winsock control depending on if there
'are more ports to scan or not.
'==========================================================================


   'Close current wisock connection
   '===============================
   Me.wskSocket(Index).Close
   
   'If there are still ports to be scanned
   '======================================
   If lngNextPort < Val(Me.txtUpperBound) Then
      
      'Connect current winsock control to next port
      '============================================
      Me.wskSocket(Index).Connect , lngNextPort
      
      'Increment lngNextPort by 1
      '==========================
      lngNextPort = lngNextPort + 1
         
   'Else, if no other ports to scan
   '===============================
   Else
      
      'Unload current winsock control
      '==============================
      Unload Me.wskSocket(Index)
      
   'End if lngNextPort < Val(Me.txtUpperBound)
   End If

'End Try_Next_Port sub
'====================
End Sub
