VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmLogin 
   BackColor       =   &H8000000E&
   Caption         =   "Login"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   5790
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc adoPengguna 
      Height          =   495
      Left            =   3240
      Top             =   2760
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Penjualan Tunai\penjualan_tunai.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Penjualan Tunai\penjualan_tunai.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from pengguna"
      Caption         =   "Ado Pengguna"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton btnLogin 
      Caption         =   "Login"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtKataKunci 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtHakAkses 
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtUserID 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Image imgLogin 
      Height          =   1455
      Left            =   3360
      Picture         =   "FrmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5520
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kata Kunci"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hak Akses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub bersih()
    For Each X In Me
        If TypeOf X Is TextBox Then
            X.Text = ""
        End If
    Next
End Sub

Sub txtHidup()
    For Each X In Me
        If TypeOf X Is TextBox Then
            X.Enabled = True
        End If
    Next
End Sub

Sub txtMati()
    For Each X In Me
        If TypeOf X Is TextBox Then
            X.Enabled = False
        End If
    Next
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnLogin_Click()
    koneksi
    qry = "select * from pengguna where Kata_kunci='" & txtKataKunci.Text & "' and User_id = '" & txtUserID.Text & "'"
    Set rsPengguna = kon.Execute(qry, , adCmdText)
    With rsPengguna
        If !Hak_akses = "User" Then
            FrmMenuUtama.stMenuUtama.Panels(1) = !User_id
            FrmMenuUtama.stMenuUtama.Panels(2) = !Nama_pengguna
            FrmMenuUtama.stMenuUtama.Panels(3) = !Hak_akses
            FrmMenuUtama.mnMaster.Enabled = True
            FrmMenuUtama.mnJurnal.Enabled = True
            FrmMenuUtama.mnPenjualan.Enabled = True
            FrmMenuUtama.smnPengguna.Enabled = False
            FrmMenuUtama.smnLogin.Enabled = False
            FrmMenuUtama.smnLogout.Enabled = True
            Unload Me
        Else
            FrmMenuUtama.stMenuUtama.Panels(1) = !User_id
            FrmMenuUtama.stMenuUtama.Panels(2) = !Nama_pengguna
            FrmMenuUtama.stMenuUtama.Panels(3) = !Hak_akses
            FrmMenuUtama.mnMaster.Enabled = True
            FrmMenuUtama.mnJurnal.Enabled = True
            FrmMenuUtama.mnPenjualan.Enabled = True
            FrmMenuUtama.smnLogin.Enabled = False
            FrmMenuUtama.smnLogout.Enabled = True
            Unload Me
        End If
        kon.Close
    End With
End Sub

Private Sub Form_Load()
    btnLogin.Enabled = False
    txtHakAkses.Enabled = False
    txtKataKunci.Enabled = False
End Sub

Private Sub txtKataKunci_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If KeyAscii = 13 Then
            koneksi
            qry = "select * from pengguna where Kata_kunci='" & txtKataKunci.Text & "' and User_id='" & txtUserID.Text & "'"
            Set rsPengguna = kon.Execute(qry, , adCmdText)
            With rsPengguna
                If .BOF And .EOF Then
                    MsgBox "Kata Kunci salah!", , "Peringatan"
                    txtKataKunci.Text = ""
                Else
                    txtKataKunci.Enabled = False
                    txtHakAkses.Text = rsPengguna!Hak_akses
                    btnLogin.Enabled = True
                    btnLogin.SetFocus
                End If
            End With
            kon.Close
        End If
    End If
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If KeyAscii = 13 Then
            koneksi
            qry = "select * from pengguna where User_id='" & txtUserID.Text & "'"
            Set rsPengguna = kon.Execute(qry, , adCmdText)
            With rsPengguna
                If .BOF And .EOF Then
                    MsgBox "User ID " + txtUserID.Text + " tidak ada", , "Peringatan"
                    txtUserID.Text = ""
                Else
                    txtUserID.Enabled = False
                    txtKataKunci.Enabled = True
                    txtKataKunci.SetFocus
                End If
            End With
            kon.Close
        End If
    End If
End Sub
