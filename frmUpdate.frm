VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update..."
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4095
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3480
      Top             =   720
   End
   Begin VB.PictureBox picStatus 
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   3555
      TabIndex        =   1
      Top             =   1200
      Width           =   3615
      Begin VB.PictureBox picBar 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   1215
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3360
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmUpdate.frx":1272
         Top             =   360
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    tmrStart.Enabled = True
    SetStatus 0
End Sub

Private Sub DownloadToFile(Location As String, Destination As String)
Dim strURL As String
Dim bData() As Byte      ' Data variable
Dim intfile As Integer   ' FreeFile variable
    On Error Resume Next
    strURL = Location
    intfile = FreeFile()      ' Set intFile to an unused file.
    
    ' The result of the OpenURL method goes into the Byte array, and the Byte array is then saved to disk.
    Open Destination For Binary Access Write As #intfile
        If Err.Number > 0 Then
            MsgBox "Error: " & Err.Number & " " & Err.Description & vbCrLf & "Unable to save file please close all programs and run the Update again."
        Else
            bData() = Inet1.OpenURL(strURL, icByteArray)
            For i = 0 To UBound(bData)
                If bData(i) = 10 Then bData(i) = 13
            Next i
            Put #intfile, , bData()
        End If
    Close #intfile
End Sub

Private Sub SetStatus(Percentage As Integer)
Dim Temp As Single
    Temp = Percentage * 0.01
    picBar.Width = Temp * picStatus.Width
End Sub

Private Sub tmrStart_Timer()
Dim DownloadPath As String, RemotePath As String
    
    tmrStart.Enabled = False
    DownloadPath = App.Path & "\data\"
    RemotePath = "http://www.ups.com/using/software/currentrates/rate-csv/"
    
    If Len(Dir(DownloadPath)) = 0 Then
        MkDir DownloadPath
    End If
    
    DownloadToFile "http://www.ups.com/using/software/currentrates/zone-csv/" & Left(LocalZip, 3) & ".csv", DownloadPath & "Zones.csv"
    SetStatus 13
    DownloadToFile RemotePath & "gndcomm.csv", DownloadPath & "gndcomm.csv"
    SetStatus 25
    DownloadToFile RemotePath & "3dscomm.csv", DownloadPath & "3dscomm.csv"
    SetStatus 38
    DownloadToFile RemotePath & "2da.csv", DownloadPath & "2da.csv"
    SetStatus 50
    DownloadToFile RemotePath & "2dam.csv", DownloadPath & "2dam.csv"
    SetStatus 63
    DownloadToFile RemotePath & "1dasaver.csv", DownloadPath & "1dasaver.csv"
    SetStatus 75
    DownloadToFile RemotePath & "1da.csv", DownloadPath & "1da.csv"
    SetStatus 88
    DownloadToFile RemotePath & "accessorials.csv", DownloadPath & "accessorials.csv"
    SetStatus 100
    Unload Me
End Sub
