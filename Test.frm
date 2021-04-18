VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "URLHistory Sample"
   ClientHeight    =   2820
   ClientLeft      =   1875
   ClientTop       =   3300
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   6570
   Begin ComctlLib.ListView URLs 
      Height          =   2730
      Left            =   15
      TabIndex        =   0
      Top             =   45
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   4815
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "URL"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Last Visited"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Last Updated"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Expires"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Flags"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuRefreshTop 
      Caption         =   "&View"
      Begin VB.Menu mnuRefresh 
         Caption         =   "Show &all"
         Index           =   0
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Show &ftp"
         Index           =   1
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Show http"
         Index           =   2
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Show fi&le"
         Index           =   3
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Show &HTML Help"
         Index           =   4
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToday 
         Caption         =   "Today only"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim History As CURLHistory

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
    
    Set History = New CURLHistory
    
    mnuRefresh_Click 0
    
End Sub


Private Sub Form_Resize()

    URLs.Move 0, 0, ScaleWidth, ScaleHeight
    
End Sub


Private Sub mnuOpen_Click()
    
    ShellExecute Me.hwnd, "open", URLs.SelectedItem.SubItems(1), "", "", 1
    
End Sub

Private Sub mnuRefresh_Click(Index As Integer)
Dim URL As URLHistoryItem, Itm As ListItem
    
    Select Case Index
        Case 1
            History.Refresh "ftp"
        Case 2
            History.Refresh "http"
        Case 3
            History.Refresh "file"
        Case 4
            History.Refresh "mk"
        Case Else
            History.Refresh
    End Select
    
    URLs.ListItems.Clear
    
    For Each URL In History
    
        If (mnuToday.Checked And DateValue(URL.LastVisited) = DateValue(Now())) Or mnuToday.Checked = False Then
            
            Set Itm = URLs.ListItems.Add(, , URL.Title)
            
            Itm.SubItems(1) = URL.URL
            Itm.SubItems(2) = URL.LastVisited
            Itm.SubItems(3) = URL.LastUpdated
            Itm.SubItems(4) = URL.Expires
            Itm.SubItems(5) = URL.Flags
            
        End If
        
    Next

End Sub

Private Sub mnuToday_Click()

    mnuToday.Checked = Not mnuToday.Checked
    
End Sub

Private Sub URLs_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)

    URLs.SortKey = ColumnHeader.Index - 1
    
End Sub


