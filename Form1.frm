VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ping Utility V1.0"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInfo 
      Height          =   690
      Left            =   7470
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   135
      Width           =   645
   End
   Begin VB.PictureBox picStatus 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   3105
      ScaleHeight     =   420
      ScaleWidth      =   5055
      TabIndex        =   20
      Top             =   135
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Label lblTimeLeft 
         AutoSize        =   -1  'True
         Caption         =   "00:25:23"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3510
         TabIndex        =   26
         Top             =   225
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Est. Time Left:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   2205
         TabIndex        =   25
         Top             =   225
         Width           =   1260
      End
      Begin VB.Label lblTimeElapsed 
         AutoSize        =   -1  'True
         Caption         =   "00:25:23"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   1260
         TabIndex        =   24
         Top             =   225
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Time Elapsed: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   23
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "255.255.255.255"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   21
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Pinging                          ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   2385
      End
   End
   Begin VB.CheckBox chkOnlyFirst 
      Caption         =   "&Only Ping First IP"
      Height          =   285
      Left            =   195
      TabIndex        =   9
      Top             =   945
      Width           =   1995
   End
   Begin VB.PictureBox picResults 
      BorderStyle     =   0  'None
      Height          =   3840
      Left            =   135
      ScaleHeight     =   3840
      ScaleWidth      =   7980
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1350
      Visible         =   0   'False
      Width           =   7980
      Begin MSComctlLib.ListView lsvResults 
         Height          =   3540
         Left            =   0
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   270
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   6244
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15136759
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "IP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Status Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Dest.Address"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "RoundTrip"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Data Size"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PING RESULTS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   7980
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   420
      Left            =   3150
      TabIndex        =   11
      Top             =   5265
      Width           =   2025
   End
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H0000FFFF&
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   3105
      ScaleHeight     =   210
      ScaleWidth      =   4260
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   585
      Visible         =   0   'False
      Width           =   4290
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100 %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   1665
         TabIndex        =   22
         Top             =   0
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6075
      TabIndex        =   12
      Top             =   5265
      Width           =   2025
   End
   Begin VB.CommandButton cmdPing 
      Caption         =   "&Ping"
      Default         =   -1  'True
      Height          =   420
      Left            =   135
      TabIndex        =   10
      Top             =   5265
      Width           =   2025
   End
   Begin VB.TextBox txtIP2 
      Height          =   285
      Index           =   4
      Left            =   2475
      MaxLength       =   3
      TabIndex        =   7
      Top             =   540
      Width           =   375
   End
   Begin VB.TextBox txtIP1 
      Height          =   285
      Index           =   4
      Left            =   2475
      MaxLength       =   3
      TabIndex        =   3
      Top             =   150
      Width           =   375
   End
   Begin VB.TextBox txtIP2 
      Height          =   285
      Index           =   3
      Left            =   1980
      MaxLength       =   3
      TabIndex        =   6
      Top             =   540
      Width           =   375
   End
   Begin VB.TextBox txtIP1 
      Height          =   285
      Index           =   3
      Left            =   1965
      MaxLength       =   3
      TabIndex        =   2
      Top             =   150
      Width           =   375
   End
   Begin VB.TextBox txtIP2 
      Height          =   285
      Index           =   2
      Left            =   1455
      MaxLength       =   3
      TabIndex        =   5
      Top             =   540
      Width           =   375
   End
   Begin VB.TextBox txtIP1 
      Height          =   285
      Index           =   2
      Left            =   1455
      MaxLength       =   3
      TabIndex        =   1
      Top             =   150
      Width           =   375
   End
   Begin VB.TextBox txtIP2 
      Height          =   285
      Index           =   1
      Left            =   945
      MaxLength       =   3
      TabIndex        =   4
      Top             =   540
      Width           =   375
   End
   Begin VB.TextBox txtIP1 
      Height          =   285
      Index           =   1
      Left            =   945
      MaxLength       =   3
      TabIndex        =   0
      Top             =   135
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   75
      Index           =   5
      Left            =   2355
      Shape           =   3  'Circle
      Top             =   750
      Width           =   105
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   75
      Index           =   4
      Left            =   1845
      Shape           =   3  'Circle
      Top             =   750
      Width           =   105
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   75
      Index           =   3
      Left            =   1335
      Shape           =   3  'Circle
      Top             =   720
      Width           =   105
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   75
      Index           =   2
      Left            =   2355
      Shape           =   3  'Circle
      Top             =   360
      Width           =   105
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   75
      Index           =   1
      Left            =   1845
      Shape           =   3  'Circle
      Top             =   330
      Width           =   105
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   75
      Index           =   0
      Left            =   1335
      Shape           =   3  'Circle
      Top             =   330
      Width           =   105
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   915
      TabIndex        =   15
      Top             =   210
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "To IP:"
      Height          =   195
      Index           =   1
      Left            =   195
      TabIndex        =   14
      Top             =   600
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "From IP:"
      Height          =   195
      Index           =   0
      Left            =   195
      TabIndex        =   13
      Top             =   210
      Width           =   585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bCancel As Boolean

Private Sub cmdCancel_Click()
   bCancel = True
End Sub

Private Sub cmdExit_Click()
   Screen.MousePointer = vbHourglass
   bCancel = True
   Dim t As Single
   Me.Refresh
   t = Timer
   Do Until Timer > t + 0.5
   Loop
   Unload Me
   End
End Sub

Private Function TimeFromSeconds(ByVal Seconds As Single) As String
   Dim Hours As Single, Mins As Single, Secs As Single
   
   Hours = Seconds \ (60 * 60)
   Mins = (Seconds - Hours * (60 * 60)) \ 60
   Secs = Seconds Mod 60
   
   TimeFromSeconds = Format(Hours, "00") & ":" & Format(Mins, "00") & ":" & Format(Secs, "00")
   
End Function

Private Sub cmdInfo_Click()
   Dim s As String
   s = "By" & vbCrLf & vbCrLf & _
   "Theo 'THE_TROOPER' Kandiliotis" & vbCrLf & _
   "theok@oneworld.gr" & vbCrLf & vbCrLf & _
   "Double Click on an IP in the results pane, to ping it again." & vbCrLf & _
   "Sort on any column of the results pane by clicking on the column title." & vbCrLf & _
   "You may interrupt the pinging process by clicking Cancel or Exit. "
   
   MsgBox s, vbInformation, App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub cmdPing_Click()

   Dim NStart As Long, NEnd As Long
   Dim CStart As Long, CEnd As Long
   Dim X, y, i As Long
   Dim IPs As Long, IPIndex As Long
   Dim CurIP As String
   Dim TimeStart As Single
   
   Screen.MousePointer = vbHourglass
   picResults.Visible = False
   DoEvents
   lsvResults.ListItems.Clear
   
   
   If chkOnlyFirst.Value = vbChecked Then
      EchoIt txtIP1(1) & "." & txtIP1(2) & "." & txtIP1(3) & "." & txtIP1(4)
      GoTo CleanExit
   Else
      
      For i = 1 To 4
         If Trim(txtIP1(i)) = "" Or Trim(txtIP2(i)) = "" Then GoTo CleanExit
      Next
      NStart = txtIP1(3): NEnd = txtIP2(3)
      CStart = txtIP1(4): CEnd = txtIP2(4)
      
      IPs = Abs(NEnd - NStart + 1) * Abs(CEnd - CStart + 1)
      IPIndex = 0
      
      picProgress.Cls
      picProgress.Visible = True
      bCancel = False
      cmdCancel.Enabled = True
      picStatus.Visible = True
      
      TimeStart = Timer
      
      For X = NStart To NEnd
         For y = CStart To CEnd
            IPIndex = IPIndex + 1
            With picProgress
               picProgress.Line (0, 0)-(IPIndex * .ScaleWidth / IPs, .ScaleHeight), , BF
               lblProgress = FormatNumber(IPIndex * 100 / IPs, 1) & " %"
               lblTimeElapsed = TimeFromSeconds(Timer - TimeStart)
               lblTimeLeft = TimeFromSeconds((Timer - TimeStart) * (IPs - IPIndex) / IPIndex)
               CurIP = txtIP1(1) & "." & txtIP1(2) & "." & X & "." & y
               EchoIt CurIP
            End With
            DoEvents
            If bCancel Then GoTo CleanExit
         Next
      Next
      
   End If
   
   'EchoIt "193.150.173.1"
   
CleanExit:
   Screen.MousePointer = vbDefault
   picResults.Visible = True
   picProgress.Visible = False
   cmdCancel.Enabled = False
   picStatus.Visible = False
End Sub

Private Sub Form_Load()
   
   Caption = App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision
   Dim i As Long
   With lsvResults
      .ColumnHeaders(1).Width = 1750
      .ColumnHeaders(2).Width = 2700
      .ColumnHeaders(3).Width = 1250
      .ColumnHeaders(4).Width = 1000
      .ColumnHeaders(5).Width = 900
   
   End With
   
   
   For i = 1 To 4
      txtIP1(i) = GetSetting(App.Title, "IP1", "Part" & CStr(i), "")
      txtIP2(i) = GetSetting(App.Title, "IP2", "Part" & CStr(i), "")
   Next
   
   Show
   

End Sub

Private Sub EchoIt(ByVal IP As String, Optional ListPos)
   
   If Len(IP) < 7 Or InStr(1, IP, ".") = 0 Then Exit Sub
   
   lblStatus = IP
   
   Dim ECHO As ICMP_ECHO_REPLY
   Dim Pos As Integer, StatusCode As String
   Dim AnItem As ListItem, ASubItem As ListSubItem
   
   Call Ping(Trim(IP), ECHO)
   
   If Not IsEmpty(ListPos) Then
      Set AnItem = lsvResults.ListItems.Add(ListPos, , IP)
   Else
      Set AnItem = lsvResults.ListItems.Add(, , IP)
   End If
   
   StatusCode = GetStatusCode(ECHO.status)
   
   With AnItem
      .ListSubItems.Add , , StatusCode
      .ListSubItems.Add , , ECHO.Address
      .ListSubItems.Add , , ECHO.RoundTripTime & " ms"
      .ListSubItems.Add , , ECHO.DataSize
   
      .Bold = True
      If InStr(1, StatusCode, "success") Then
         .ForeColor = &H4000&
      Else
         .ForeColor = vbRed
      End If
   
   End With
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim i As Long
   For i = 1 To 4
      SaveSetting App.Title, "IP1", "Part" & CStr(i), txtIP1(i)
      SaveSetting App.Title, "IP2", "Part" & CStr(i), txtIP2(i)
   Next
End Sub

Private Sub lsvResults_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   
      
      Static Sort(0 To 10) As Boolean
      
      Screen.MousePointer = vbHourglass
      DoEvents
      lsvResults.SortKey = ColumnHeader.Index - 1
      Sort(ColumnHeader.Index) = Not Sort(ColumnHeader.Index)
      lsvResults.SortOrder = IIf(Sort(ColumnHeader.Index), lvwAscending, lvwDescending)
      lsvResults.Sorted = True
      Screen.MousePointer = vbDefault
   
   
'   Static Sort0 As Boolean, Sort1 As Boolean, Sort3 As Boolean, Sort4 As Boolean
'
'   Select Case ColumnHeader.Text
'      Case "IP"
'         Screen.MousePointer = vbHourglass
'         lsvResults.SortKey = 0
'         Sort0 = Not Sort0
'         lsvResults.SortOrder = IIf(Sort0, lvwAscending, lvwDescending)
'         lsvResults.Sorted = True
'         Screen.MousePointer = vbDefault
'      Case "Status Code"
'         Screen.MousePointer = vbHourglass
'         lsvResults.SortKey = 1
'         Sort1 = Not Sort1
'         lsvResults.SortOrder = IIf(Sort1, lvwAscending, lvwDescending)
'         lsvResults.Sorted = True
'         Screen.MousePointer = vbDefault
'      Case "RoundTrip"
'         Screen.MousePointer = vbHourglass
'         lsvResults.SortKey = 3
'         Sort3 = Not Sort3
'         lsvResults.SortOrder = IIf(Sort3, lvwAscending, lvwDescending)
'         lsvResults.Sorted = True
'         Screen.MousePointer = vbDefault
'      Case "Data Size"
'         Screen.MousePointer = vbHourglass
'         lsvResults.SortKey = 4
'         Sort4 = Not Sort4
'         lsvResults.SortOrder = IIf(Sort4, lvwAscending, lvwDescending)
'         lsvResults.Sorted = True
'         Screen.MousePointer = vbDefault
'   End Select

End Sub

Private Sub lsvResults_DblClick()
   Dim IP As String
   Dim Pos As Long
   With lsvResults
      Pos = .SelectedItem.Index
      IP = .SelectedItem.Text
      .ListItems.Remove Pos
      EchoIt IP, Pos
      .ListItems(Pos).Selected = True
      .SetFocus
   End With
End Sub

Private Sub txtIP1_GotFocus(Index As Integer)
   txtIP1(Index).SelStart = 0
   txtIP1(Index).SelLength = Len(txtIP1(Index))
End Sub

'Private Sub txtIP1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If Len(txtIP1(Index)) = 3 Then
'      Select Case Index
'         Case 1, 2, 3
'            txtIP1(Index + 1).SetFocus
'         Case 4
'            txtIP2(1).SetFocus
'      End Select
'   End If
'End Sub

'Private Sub txtIP2_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If Len(txtIP2(Index)) = 3 Then
'      Select Case Index
'         Case 1, 2, 3
'            txtIP2(Index + 1).SetFocus
'         Case 4
'            cmdPing.SetFocus
'      End Select
'   End If
'End Sub

Private Sub txtIP2_GotFocus(Index As Integer)
   txtIP2(Index).SelStart = 0
   txtIP2(Index).SelLength = Len(txtIP2(Index))
End Sub


Private Sub txtIP1_Validate(Index As Integer, Cancel As Boolean)
   Dim Value As String
   Value = txtIP1(Index)
   If Trim(Value) <> "" Then
      If Not IsNumeric(Value) Then Cancel = True: Exit Sub
      If Int(Value) <> Value Then Cancel = True: Exit Sub
      If Int(Value) < 0 Or Int(Value) > 255 Then Cancel = True: Exit Sub
   End If
End Sub

Private Sub txtIP2_Validate(Index As Integer, Cancel As Boolean)
   Dim Value As String
   Value = txtIP2(Index)
   If Trim(Value) <> "" Then
      If Not IsNumeric(Value) Then Cancel = True: Exit Sub
      If Int(Value) <> Value Then Cancel = True: Exit Sub
      If Int(Value) < 0 Or Int(Value) > 255 Then Cancel = True: Exit Sub
   End If

End Sub
