VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{9A03465D-3CA7-4DAA-9024-7738C42A706E}#1.1#0"; "IMPULSEMP3.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MP3 Player"
   ClientHeight    =   5460
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5460
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Hide Playlist Editor"
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open Playlist"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   5040
      Width           =   1095
   End
   Begin ImpulseMP3.ISMP3PlayList MP3PlayList 
      Left            =   2280
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   0
      TabIndex        =   6
      Top             =   1800
      Width           =   4335
   End
   Begin ImpulseMP3.ISMP3Player MP3Engine 
      Left            =   1680
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      FillColor       =   &H80000012&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   480
      Width           =   1575
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   555
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   960
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3360
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   1
      ImageHeight     =   1
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0000
            Key             =   "Dot"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   688
      BandCount       =   2
      FixedOrder      =   -1  'True
      _CBWidth        =   4335
      _CBHeight       =   390
      _Version        =   "6.0.8169"
      Child1          =   "Toolbar2"
      MinHeight1      =   315
      Width1          =   1800
      NewRow1         =   0   'False
      BandBackColor2  =   -2147483648
      Child2          =   "Toolbar1"
      MinHeight2      =   330
      Width2          =   3210
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   315
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   556
         ButtonWidth     =   820
         ButtonHeight    =   556
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&File"
               Key             =   "File"
               ImageKey        =   "Dot"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Edit"
               Key             =   "Edit"
               ImageKey        =   "Dot"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   1995
         TabIndex        =   2
         Top             =   30
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Back"
               ImageKey        =   "Back"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Play"
               ImageKey        =   "Play"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Pause"
               ImageKey        =   "Pause"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Stop"
               ImageKey        =   "Stop"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Forward"
               ImageKey        =   "Forward"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               ImageKey        =   "Open"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2760
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":005C
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":03B0
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0704
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0A58
            Key             =   "Pause"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0DAC
            Key             =   "Play"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1100
            Key             =   "Back"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar ProgressBar1 
      Height          =   255
      Left            =   -240
      TabIndex        =   0
      Top             =   1440
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Appearance      =   2
      Arrows          =   65536
      Orientation     =   8323073
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3960
      Top             =   360
   End
   Begin MSComDlg.CommonDialog cdbDialogBox 
      Left            =   4440
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuPlayList 
         Caption         =   "Show Playlist"
      End
      Begin VB.Menu mnuMinimize 
         Caption         =   "&Minimize"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Player Controls"
      Visible         =   0   'False
      Begin VB.Menu mnuPlay 
         Caption         =   "&Play"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "&Stop"
      End
      Begin VB.Menu mnuBack 
         Caption         =   "&Back"
      End
      Begin VB.Menu mnuForward 
         Caption         =   "&Forward"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MP3FileName As String, DisplayTime As Boolean, TimeDuration As String
Dim TimeReversed As String
Public playlistindex As Integer
Dim sArtist As String
Dim sTitle As String
Dim sEntryCaption As String
Dim iEntry As Long
Dim Hiney As Long, ShowingPlayList As Boolean
Dim SFileName As String

Private Sub Command1_Click()
    ShowingPlayList = True
    mnuPlayList_Click
End Sub

Private Sub Command2_Click()
    mnuOpen_Click
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.Files(1) <> "" Then
        If Right(Data.Files(1), 3) = "mp3" Then
            OpenSingleFile (Data.Files(1))
        ElseIf Right(Data.Files(1), 3) = "m3u" Then
            OpenPlayListFile (Data.Files(1))
        End If
    Else
        Exit Sub
    End If
End Sub

Private Sub List1_DblClick()
    MP3Engine.FileName = MP3PlayList(List1.ListIndex + 1).FileName
    If MP3Engine.FileName <> "" Then
        If MP3Engine.IsPlaying = False Then
            ProgressBar1.Value = 0
            ProgressBar1.Max = Int(MP3Engine.TrackLength)
            ProgressBar1.Enabled = True
            Timer1.Enabled = True
            MP3Engine.PlayStart
        ElseIf MP3Engine.IsPlaying = True Then
            Exit Sub
        Else
            ProgressBar1.Max = Int(MP3Engine.TrackLength)
            MP3Engine.PlayStart
        End If
    End If
End Sub

Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
    Select Case NewHeight
        Case Is >= 730
            frmMain.Height = frmMain.Height + 345
            Picture1.Top = Picture1.Top + 345
            Command1.Top = Command1.Top + 345
            ProgressBar1.Top = ProgressBar1.Top + 345
        Case Is <= 395
            frmMain.Height = frmMain.Height - 345
            Picture1.Top = Picture1.Top - 345
            Command1.Top = Command1.Top - 345
            ProgressBar1.Top = ProgressBar1.Top - 345
    End Select
End Sub

Private Sub Form_Load()
On Error Resume Next
    DisplayTime = True
    ShowingPlayList = False
    frmMain.Height = 2160
    If Command <> "" Then
        If Right(Command, 3) = "mp3" Then
            OpenSingleFile (Command)
        ElseIf Right(Command, 3) = "m3u" Then
            OpenPlayListFile (Command)
        End If
    Else
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MP3Engine.PlayStop
End Sub

Private Sub Label1_Click()
    DisplayTime = Not DisplayTime
    If MP3Engine.IsPlaying = True Then
        If DisplayTime = True Then
            Label1.Caption = TimeDuration
        ElseIf DisplayTime = False Then
            Label1.Caption = TimeReversed
        End If
    Else
        If DisplayTime = True Then
            Label1.Caption = "0:00"
        ElseIf DisplayTime = False Then
            Label1.Caption = "-" & SecondsToTime(MP3Engine.TrackLength)
        End If
    End If
End Sub

Private Sub mnuBack_Click()
    MP3Engine.TrackPosition(impMP3SeekFromBeginning) = MP3Engine.TrackPosition(impMP3SeekFromBeginning) - 1
End Sub

Private Sub mnuMinimize_Click()
    frmMain.WindowState = vbMinimized
End Sub

Private Sub mnuPause_Click()
    If (Not MP3Engine.IsPlaying) Or MP3Engine.IsPaused Then Exit Sub
    
    MP3Engine.PlayPause
End Sub

Public Sub mnuPlay_Click()
    If MP3Engine.FileName <> "" Then
        If MP3Engine.IsPaused = True Then
            ProgressBar1.Value = 0
            ProgressBar1.Enabled = True
            Timer1.Enabled = True
            MP3Engine.PlayStart
        ElseIf MP3Engine.IsPlaying = True Then
            Exit Sub
        Else
            MP3Engine.PlayStart
        End If
    End If
End Sub

Private Sub mnuPlayList_Click()
    If ShowingPlayList = False Then
        Do
            frmMain.Height = frmMain.Height + 2
        Loop Until frmMain.Height >= 5835
        frmMain.Height = 5835
        ShowingPlayList = True
        mnuPlayList.Caption = "Hide Playlist"
    ElseIf ShowingPlayList = True Then
        Do
            frmMain.Height = frmMain.Height - 2
        Loop Until frmMain.Height <= 2160
        frmMain.Height = 2160
        ShowingPlayList = False
        mnuPlayList.Caption = "Show Playlist"
    End If
End Sub

Private Sub mnuStop_Click()
    MP3Engine.PlayStop
    Timer1.Enabled = False
    If DisplayTime = True Then
        Label1.Caption = "0:00"
    ElseIf DisplayTime = False Then
        Label1.Caption = "-" & SecondsToTime(MP3Engine.TrackLength)
    End If
    ProgressBar1.Value = 0
    ProgressBar1.Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyL Or Button = vbLeftButton Then
        MP3Engine.TrackPosition(impMP3SeekFromBeginning) = MP3Engine.TrackPosition(impMP3SeekFromBeginning) + 1
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub ProgressBar1_Scroll()
    MP3Engine.TrackPosition(impMP3SeekFromBeginning) = ProgressBar1.Value
    TimeDuration = SecondsToTime(ProgressBar1.Value)
    TimeReversed = "-" & SecondsToTime(MP3Engine.TrackLength - ProgressBar1.Value)
    If DisplayTime = True Then
        Label1.Caption = TimeDuration
    ElseIf DisplayTime = False Then
        Label1.Caption = TimeReversed
    End If
End Sub

Public Sub Timer1_Timer()
    ProgressBar1.Value = MP3Engine.TrackPosition(impMP3SeekFromBeginning)
    TimeDuration = SecondsToTime(ProgressBar1.Value)
    TimeReversed = "-" & SecondsToTime(MP3Engine.TrackLength - ProgressBar1.Value)
    If DisplayTime = True Then
        Label1.Caption = TimeDuration
    ElseIf DisplayTime = False Then
        Label1.Caption = TimeReversed
    End If
    If ProgressBar1.Value = ProgressBar1.Max Then
        Timer1.Enabled = False
        ProgressBar1.Enabled = False
    End If
    If MP3Engine.IsPaused = True Then
        Label1.Visible = Not Label1.Visible
    Else
        Label1.Visible = True
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Play"
            mnuPlay_Click
        Case "Open"
            mnuOpen_Click
        Case "Pause"
            mnuPause_Click
        Case "Stop"
            mnuStop_Click
    End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "File"
            PopupMenu mnuFile, , (Toolbar2.Left + Toolbar2.Buttons(1).Left), (CoolBar1.Bands(1).Height - (Toolbar2.Top + Toolbar2.Buttons(1).Top) + 20)
        Case "Edit"
            PopupMenu mnuEdit, , (Toolbar2.Left + Toolbar2.Buttons(2).Left), (CoolBar1.Bands(1).Height - (Toolbar2.Top + Toolbar2.Buttons(2).Top) + 20)
    End Select
End Sub

Private Function SecondsToTime(lSeconds As Double) As String
    Dim sTime As String
    Dim iSeconds As Integer
    Dim iMinutes As Integer
    
    iSeconds = Abs(Fix(lSeconds)) Mod 60
    iMinutes = Fix(Abs(Fix(lSeconds)) / 60)
    
    sTime = iMinutes & ":" & IIf(iSeconds < 10, "0", "") & iSeconds
    
    SecondsToTime = sTime
End Function

Private Sub RefreshPlayListDisplay()
    List1.Visible = False
    Screen.MousePointer = vbHourglass
    Me.Refresh
    List1.Clear

    For iEntry = 1 To MP3PlayList.Count
        With MP3PlayList.PlayListEntries(iEntry)
            sArtist = MP3PlayList.PlayListEntries(iEntry).ID3Artist
            sTitle = MP3PlayList.PlayListEntries(iEntry).ID3Title
            If sTitle = "" And sArtist = "" Then
                sEntryCaption = MP3PlayList.PlayListEntries(iEntry).FileName
            ElseIf sTitle = "" Then
                sEntryCaption = sArtist & " - Untitled"
            ElseIf sArtist = "" Then
                sEntryCaption = "Unknown Artist - " & sTitle
            Else
                sEntryCaption = sArtist & " - " & sTitle
            End If
            List1.AddItem iEntry & ". " & sEntryCaption
            List1.ItemData(List1.NewIndex) = iEntry
        End With
    Next iEntry
    
    List1.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuOpen_Click()
    On Error GoTo EndOfSub
    cdbDialogBox.FileName = ""
    cdbDialogBox.DialogTitle = "Add File To Playlist"
    cdbDialogBox.DefaultExt = "mp3"
    cdbDialogBox.CancelError = True
    cdbDialogBox.Filter = "Compatible Audio Files|*.mp3;*.m3u|MP3 Audio (*.mp3)|*.mp3|PlayList (*.m3u)|*.m3u"
    cdbDialogBox.ShowOpen
    SFileName = cdbDialogBox.FileName
    If Right(SFileName, 3) = "mp3" Then
        OpenSingleFile (SFileName)
    ElseIf Right(SFileName, 3) = "m3u" Then
        OpenPlayListFile (SFileName)
    End If
EndOfSub:
    Exit Sub
End Sub

Private Sub OpenSingleFile(SFileName As String)
    If SFileName <> "" Then
        MP3PlayList.AddEntry SFileName
        sArtist = MP3PlayList.PlayListEntries(MP3PlayList.Count).ID3Artist
        sTitle = MP3PlayList.PlayListEntries(MP3PlayList.Count).ID3Title
        If sTitle = "" And sArtist = "" Then
           sEntryCaption = MP3PlayList.PlayListEntries(MP3PlayList.Count).FileName
        ElseIf sTitle = "" Then
            sEntryCaption = sArtist & " - Untitled"
        ElseIf sArtist = "" Then
            sEntryCaption = "Unknown Artist - " & sTitle
        Else
            sEntryCaption = sArtist & " - " & sTitle
        End If
        List1.AddItem MP3PlayList.Count & ". " & sEntryCaption
        ProgressBar1.Value = 0
        MP3Engine.FileName = SFileName
        ProgressBar1.Max = Int(MP3Engine.TrackLength)
        ProgressBar1.Enabled = True
        Timer1.Enabled = True
        frmMain.Caption = sEntryCaption
        MP3Engine.PlayStart
    Else
        Exit Sub
    End If
End Sub

Private Sub OpenPlayListFile(SFileName As String)
    If SFileName <> "" Then
        MP3PlayList.LoadPlayList SFileName
        RefreshPlayListDisplay
    End If
End Sub
