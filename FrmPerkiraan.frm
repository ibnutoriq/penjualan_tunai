VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmPerkiraan 
   BackColor       =   &H00FFFF00&
   Caption         =   "Perkiraan"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   11235
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Pencarian"
      Height          =   855
      Left            =   6840
      TabIndex        =   13
      Top             =   3240
      Width           =   4095
      Begin VB.CommandButton btnCari 
         Caption         =   "&Cari"
         Height          =   495
         Left            =   3000
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtCari 
         Height          =   495
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   1300
      End
      Begin VB.Label Label5 
         Caption         =   "Kode Akun"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton btnBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   2640
      TabIndex        =   12
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton btnSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton btnHapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   3840
      TabIndex        =   11
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton btnKeluar 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton btnTambah 
      Caption         =   "&Tambah"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid dtgPerkiraan 
      Bindings        =   "FrmPerkiraan.frx":0000
      Height          =   1935
      Left            =   4800
      TabIndex        =   8
      Top             =   1080
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3413
      _Version        =   393216
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Kode_akun"
         Caption         =   "Kode_akun"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Nama_akun"
         Caption         =   "Nama_akun"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Tipe_akun"
         Caption         =   "Tipe_akun"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2340,284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2550,047
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtNamaAkun 
      Height          =   495
      Left            =   1680
      MaxLength       =   25
      TabIndex        =   1
      Top             =   1800
      Width           =   2800
   End
   Begin VB.TextBox txtKodeAkun 
      Height          =   495
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtTipeAkun 
      Height          =   495
      Left            =   1680
      MaxLength       =   25
      TabIndex        =   3
      Top             =   2520
      Width           =   2800
   End
   Begin MSAdodcLib.Adodc adoPerkiraan 
      Height          =   495
      Left            =   8400
      Top             =   2520
      Width           =   2415
      _ExtentX        =   4260
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
      RecordSource    =   "select * from perkiraan"
      Caption         =   "Ado Perkiraan"
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
   Begin VB.Label Label4 
      Caption         =   "Tipe Akun"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Kode Akun"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Akun"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Perkiraan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   10695
   End
End
Attribute VB_Name = "FrmPerkiraan"
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
            txtCari.Enabled = True
        End If
    Next
End Sub

Sub simpanPerkiraan()
    ' Membuat validasi pada field Kode_akun dan jika false maka simpan isi TextBox ke table perkiraan
    On Error Resume Next
    With adoPerkiraan.Recordset
        adoPerkiraan.Recordset.Find "Kode_akun='" & txtKodeAkun.Text & "'", , adSearchForward
        If Not adoPerkiraan.Recordset.EOF Then
            MsgBox "Kode Akun Sudah Terdaftar", vbOKOnly, "Informasi"
            Exit Sub
        Else
            .AddNew
                !Kode_akun = txtKodeAkun.Text
                !Nama_akun = txtNamaAkun.Text
                !Tipe_akun = txtTipeAkun.Text
            .Update
        End If
    End With
    On Error GoTo 0
End Sub

Private Sub btnBatal_Click()
    bersih
    txtMati
    btnBatal.Enabled = False
    btnTambah.Enabled = True
    btnSimpan.Enabled = False
    btnHapus.Enabled = True
End Sub

Private Sub btnCari_Click()
    On Error Resume Next
    adoPerkiraan.Recordset.MoveFirst
    adoPerkiraan.Recordset.Find "Kode_akun='" & txtCari.Text & "'", , adSearchForward
    If adoPerkiraan.Recordset.EOF Then
        MsgBox "Kode Akun Tidak Ada!", vbOKOnly, "Informasi"
    End If
    On Error GoTo 0
End Sub

Private Sub btnHapus_Click()
    On Error Resume Next
    a = MsgBox("Yakin Mau Dihapus  ???", vbYesNo + vbInformation, "Konfirmasi")
    If a = vbYes Then
        adoPerkiraan.Recordset.Delete
        txtMati
        bersih
    End If
    On Error GoTo 0
End Sub

Private Sub btnKeluar_Click()
    Unload Me
End Sub

Private Sub btnSimpan_Click()
    simpanPerkiraan
    bersih
    txtMati
    btnSimpan.Enabled = False
    btnTambah.Enabled = True
    btnHapus.Enabled = True
    btnBatal.Enabled = False
    btnTambah.SetFocus
End Sub

Private Sub btnTambah_Click()
    txtHidup
    bersih
    btnTambah.Enabled = True
    btnBatal.Enabled = True
    btnSimpan.Enabled = True
    txtKodeAkun.SetFocus
End Sub

Private Sub dtgPerkiraan_DblClick()
    On Error Resume Next
    adoPerkiraan.Recordset.Delete
    On Error GoTo 0
End Sub

Private Sub dtgPerkiraan_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    On Error GoTo 0
End Sub

Private Sub Form_Activate()
    txtMati
    btnTambah.SetFocus
    btnSimpan.Enabled = False
    btnBatal.Enabled = False
End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        btnCari.SetFocus
    End If
End Sub

Private Sub txtKodeAkun_KeyPress(KeyAscii As Integer)
    ' Membuat validasi pada field Kode_akun dan jika true maka bersih
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        adoPerkiraan.Recordset.Find "Kode_akun='" & txtKodeAkun.Text & "'", , adSearchForward
        If Not adoPerkiraan.Recordset.EOF Then
            MsgBox "Kode Akun Sudah Terdaftar", vbOKOnly, "Informasi"
            txtKodeAkun.Text = ""
            txtKodeAkun.SetFocus
            Exit Sub
            btnTambah.Enabled = True
            btnSimpan.Enabled = True
            btnBatal.Enabled = False
        End If
        txtNamaAkun.SetFocus
        btnSimpan.Enabled = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtNamaAkun_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTipeAkun.SetFocus
    End If
End Sub

Private Sub txtTipeAkun_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        btnSimpan.SetFocus
    End If
End Sub
