VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "UPS Rate Calc"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4095
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   3855
      Begin VB.TextBox txtZip 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtWeight 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox cmbService 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.CheckBox chkLetter 
         Caption         =   "Letter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox chkCOD 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   1305
         Width           =   375
      End
      Begin VB.TextBox txtValue 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox txtCost 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txtCharge 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ZIP CODE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "SERVICE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "WEIGHT:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "COD:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "INSUR VAL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1740
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "COST:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2220
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "CHARGE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2580
         Width           =   975
      End
   End
   Begin VB.Menu MnuOption 
      Caption         =   "&Options"
   End
   Begin VB.Menu MnuUpdate 
      Caption         =   "&Update Data"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Zones.csv can be gotten http://www.ups.com/using/software/currentrates/zone-csv/117.csv

Private Sub chkLetter_Click()
    If chkLetter.Value Then
        txtWeight.Locked = True
        txtWeight.Text = "Letter"
    Else
        txtWeight.Locked = False
        txtWeight.Text = ""
    End If
End Sub

Private Sub Form_Load()
Dim ct As Integer
    cmbService.AddItem "Ground"
    cmbService.AddItem "3 Day Select"
    cmbService.AddItem "BLUE - 2nd Day Air"
    cmbService.AddItem "BLUE - 2nd Day Air A.M."
    cmbService.AddItem "RED - Next Day Air Saver"
    cmbService.AddItem "RED - Next Day Air"
    cmbService.AddItem "RED - Next Day Air Early AM"
    cmbService.AddItem "RED - Next Day Air Saturday"
    cmbService.ListIndex = 0
    
    LocalZip = GetSetting(App.ProductName, "Settings", "LocalZip", 117)
    CODCharge = GetSetting(App.ProductName, "Settings", "CODCharge", 7)
    EarlyCharge = GetSetting(App.ProductName, "Settings", "EarlyCharge", 28.5)
    SaturdayCharge = GetSetting(App.ProductName, "Settings", "SaturdayCharge", 12.5)
    Handling = GetSetting(App.ProductName, "Settings", "Handling", 15)
End Sub

Private Sub MnuOption_Click()
    frmOptions.Show 1, Me
End Sub

Private Sub MnuUpdate_Click()
    frmUpdate.Show 1, Me
End Sub

Private Sub txtValue_Change()
    CalcTotal
End Sub

Private Sub txtWeight_Change()
    CalcTotal
End Sub

Private Sub chkCOD_Click()
    CalcTotal
End Sub

Private Sub txtZip_Change()
    If Len(txtZip.Text) >= 3 Then
        Zone = FindZone(txtZip.Text, cmbService.ListIndex)
    End If
    CalcTotal
End Sub

Private Sub cmbService_Click()
    If Len(txtZip.Text) >= 3 Then
        Zone = FindZone(txtZip.Text, cmbService.ListIndex)
    End If
    
    If cmbService.ListIndex > 1 Then
        chkLetter.Enabled = True
    Else
        chkLetter.Enabled = False
        chkLetter.Value = 0
        If txtWeight.Text = "Letter" Then txtWeight.Text = ""
    End If
    
    CalcTotal
End Sub

Private Sub CalcTotal()
Dim SubTotal As Single, InsureValue As Integer
    
    If Len(txtZip.Text) >= 3 Then
        Zone = FindZone(txtZip.Text, cmbService.ListIndex)
    End If
    
    If Zone <> "-" And Left(Zone, 1) <> "[" Then
        If Len(txtWeight.Text) > 0 Then
            If Val(txtWeight.Text) < 150 Or txtWeight.Text = "Letter" Then
                SubTotal = GetBaseCost(Zone, txtWeight.Text, cmbService.ListIndex)
                If chkCOD.Value Then SubTotal = SubTotal + 7
                
                InsureValue = Val(txtValue.Text)
                InsureValue = (InsureValue \ 100)
                If Right(txtValue.Text, 2) = "00" Then InsureValue = InsureValue - 1
                SubTotal = SubTotal + (InsureValue * 0.35)
                
                If cmbService.ListIndex = 6 Then SubTotal = SubTotal + EarlyCharge
                
                If cmbService.ListIndex = 7 Then SubTotal = SubTotal + SaturdayCharge
                
                txtCost.Text = Format(SubTotal, "$0.00")
                txtCharge.Text = Format(SubTotal + (SubTotal * (Handling / 100)), "$0.00")
            Else
                MsgBox "Weight can't exceed 150 pounds."
            End If
        End If
    Else
        If Zone = "-" Then
            txtCharge.Text = "Not Avaliable."
        Else
            txtCharge.Text = "See Zones Chart."
        End If
    End If
End Sub
