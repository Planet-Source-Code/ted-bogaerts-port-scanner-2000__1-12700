VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000002&
   Caption         =   "Port Scanner 2000"
   ClientHeight    =   5550
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7335
   Icon            =   "PortScanner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "PortScanner.frx":27A2
   ScaleHeight     =   5550
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLocalHost 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5040
      TabIndex        =   13
      ToolTipText     =   "Local Host Name (Computer Name)"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtLocalIP 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5040
      TabIndex        =   10
      ToolTipText     =   "Local I.P. of this computer"
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdStopScan 
      Caption         =   "Stop Scan"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   9720
      Top             =   840
   End
   Begin VB.TextBox txtEndingPort 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3360
      TabIndex        =   4
      ToolTipText     =   "Enter Ending Port Number (Last port scanned will be one less than this number)"
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtStartingPort 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      ToolTipText     =   "Enter Starting Port Number"
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Enter I.P. or Host Name (i.e. www.microsoft.com)"
      Top             =   720
      Width           =   1695
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FFFF&
      Height          =   3765
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   4095
   End
   Begin VB.CommandButton cmdStartScan 
      Caption         =   "Start Scan"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   4560
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   9720
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   7320
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Host Name"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local I.P."
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   11
      Top             =   360
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   360
      X2              =   4440
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   360
      X2              =   4440
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ending Port"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Starting Port"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I.P. or Host Name"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   360
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000012&
      Height          =   5175
      Left            =   240
      Top             =   240
      Width           =   4335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   1935
      Left            =   4920
      Top             =   240
      Width           =   2175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuStartScanning 
         Caption         =   "Start Scanning"
      End
      Begin VB.Menu mnuStopScanning 
         Caption         =   "Stop Scanning"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'Program: Port Scanner 2000
'Scans Remote ports as well as local ports
'Date: 10/05/2000
'******************************************************************************************

Option Explicit 'Declare variables before they are used

'**************************General Variable Declarations***********************************
Dim pstrIp As String 'Stores the value of the remote I.P.
Dim msngPort As Single 'Stores the value of the starting remote port
Dim msngEndingPort As Single 'Stores the value of the ending remote port
Dim msngStop As Single 'Variable used to check if stop button has been clicked
Dim x As Single 'Used as counter for loops to slow program down (prevents buffer errors)

'*******************************************************************************************
'Exits program
'*******************************************************************************************
Private Sub cmdExit_Click()
    Dim pstrmessage As String
    pstrmessage = MsgBox("Are you sure you want to quit?", vbYesNo)
        If pstrmessage = vbYes Then
            Unload Form1
        End If
End Sub

'*******************************************************************************************
'Starts scanning selected ports
'*******************************************************************************************
Private Sub CmdStartScan_Click()
    msngPort = Val(txtStartingPort.Text) 'assigns starting port value
    msngEndingPort = Val(txtEndingPort.Text) 'assigns ending port value
    msngStop = 0 'Sets initial value of msngStop to 0 (1 = Break the following loop)
    List1.Clear 'clears list box before starting scan and adds banner
    List1.AddItem ("         ************** Port Scanner 2000 **************")
    tedlock 'locks controls until scan is terminated or finished
    Dim msngTotalPorts As Single
    
'******************************************************************************************
'Loop routine - continues to loop through following code until msngPort = End port
'******************************************************************************************
    Do Until msngPort = msngEndingPort
    msngTotalPorts = msngTotalPorts + 1
    'Checks to see if cmdStopScanning has been clicked, if so, breaks out of loop
        If msngStop = 1 Then
            Exit Sub
        End If
    
        'Slows program down - prevents buffer errors
            For x = 1 To 1500000
            Next

    DoEvents 'slows down program
    pstrIp = txtIP.Text 'sets value of pstrIp to txtIP.text
    DoEvents 'slows down program
    Winsock1.Close
    Winsock1.Connect pstrIp, msngPort 'connects to I.P. and Port specified in txtIp & txtStartingPort
        
        'Slows program down - prevents buffer errors
            For x = 1 To 1500000
            Next
    
    msngPort = msngPort + 1 'increases value of port being scanned by 1
    txtEndingPort.Text = msngPort 'increments txtEndingPort for visual aid in current port
    
        'Slows program down - prevents buffer errors
            For x = 1 To 1500000
            Next

    DoEvents
            'Unlocks controls if scan makes it to ending port number
            If msngPort = msngEndingPort Then
                UnlockControls
                List1.AddItem (" ")
                List1.AddItem (" ")
                List1.AddItem ("Completed Scanning " & msngTotalPorts & " Ports" & " !!")
            End If
            
    Loop 'returns program to beginning of loop
End Sub

'*******************************************************************************************
'Command button to stop scanning before ending port has been reached
'*******************************************************************************************
Private Sub cmdStopScan_Click()
'Assigns value of 1 to msngStop - when returning to main loop, program will break based on this value
    UnlockControls 'Unlocks controls when stop has been clicked
    msngStop = 1
    List1.AddItem (" ")
    List1.AddItem (" ")
    List1.AddItem ("Scan Terminated!")
End Sub

'*******************************************************************************************
'Misc routines to run when form loads
'*******************************************************************************************
Private Sub Form_Load()
    txtIP.Text = Winsock1.LocalIP 'assigns local i.p. as first i.p to scan
    txtStartingPort.Text = 1
    txtEndingPort.Text = 65530
    List1.AddItem ("         ************** Port Scanner 2000 **************")
    txtLocalIP.Text = Winsock1.LocalIP 'displays local ip in local info data area
    txtLocalHost.Text = Winsock1.LocalHostName 'displays local host name in local info data area
    txtLocalIP.Locked = True
    txtLocalHost.Locked = True
    UnlockControls
End Sub

Private Sub mnuAbout_Click()
List1.AddItem (" ")
List1.AddItem (" ")
List1.AddItem (" *************************************************************")
List1.AddItem ("  It scans Ports!!! What else do you need to know?!?         *")
List1.AddItem (" *************************************************************")
End Sub

Private Sub mnuExit_Click()
    cmdExit_Click
End Sub

Private Sub mnuStartScanning_Click()
    CmdStartScan_Click
End Sub

Private Sub mnuStopScanning_Click()
    cmdStopScan_Click
End Sub

'*******************************************************************************************
'Checks to see if msngport has made a connection, if so, adds item to listbox
'*******************************************************************************************
Private Sub Winsock1_Connect()
    List1.AddItem ("Port " & Winsock1.RemotePort & " Connected")
End Sub

'*******************************************************************************************
'Locks controls while scan is running
'*******************************************************************************************
Private Sub tedlock()
    cmdExit.Enabled = False
    cmdStartScan.Enabled = False
    mnuFile.Enabled = False
    cmdStopScan.Enabled = True
    txtIP.Locked = True
    txtStartingPort.Locked = True
    txtEndingPort.Locked = True
    mnuTools.Enabled = False
    mnuHelp.Enabled = False
End Sub

'*******************************************************************************************
'Unlocks controls when scan is finished or has been stopped
'*******************************************************************************************
Private Sub UnlockControls()
    cmdExit.Enabled = True
    cmdStartScan.Enabled = True
    cmdStopScan.Enabled = False
    mnuFile.Enabled = True
    txtIP.Locked = False
    txtStartingPort.Locked = False
    txtEndingPort.Locked = False
    mnuTools.Enabled = True
    mnuHelp.Enabled = True
End Sub

