VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmSuratJalan 
   BackColor       =   &H00FFFF00&
   Caption         =   "Surat Jalan"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   13980
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbKodeBarang 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin VB.ComboBox cmbKodePelanggan 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pencarian"
      Height          =   855
      Left            =   9120
      TabIndex        =   10
      Top             =   4920
      Width           =   4575
      Begin VB.CommandButton btnCari 
         Caption         =   "&Cari"
         Height          =   495
         Left            =   3480
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtCari 
         Height          =   495
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   1300
      End
      Begin VB.Label Label5 
         Caption         =   "Nomor Surat Jalan"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton btnBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton btnSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton btnHapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton btnKeluar 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   5040
      TabIndex        =   7
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton btnTambah 
      Caption         =   "&Tambah"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtNomorSuratJalan 
      Height          =   495
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   0
      Top             =   1200
      Width           =   1800
   End
   Begin VB.TextBox txtKuantitas 
      Height          =   495
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   3
      Top             =   3360
      Width           =   1500
   End
   Begin VB.TextBox txtKeterangan 
      Height          =   495
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   5
      Top             =   4080
      Width           =   2800
   End
   Begin MSDataGridLib.DataGrid dtgSuratJalan 
      Bindings        =   "FrmSuratJalan.frx":0000
      Height          =   3375
      Left            =   4920
      TabIndex        =   14
      Top             =   1200
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5953
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "No_sj"
         Caption         =   "No_sj"
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
         DataField       =   "Kode_pelanggan"
         Caption         =   "Kode_pelanggan"
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
         DataField       =   "Kode_barang"
         Caption         =   "Kode_barang"
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
      BeginProperty Column03 
         DataField       =   "Kuantitas"
         Caption         =   "Kuantitas"
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
      BeginProperty Column04 
         DataField       =   "Keterangan_sj"
         Caption         =   "Keterangan_sj"
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
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1395,213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   4080,189
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoSuratJalan 
      Height          =   495
      Left            =   11160
      Top             =   4080
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      RecordSource    =   "select * from surat_jalan"
      Caption         =   "Ado Surat Jalan"
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
      Caption         =   "Kode Barang"
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Nomor Surat Jalan"
      Height          =   495
      Left            =   240
      TabIndex        =   19
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Kode Pelanggan"
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Surat Jalan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   13455
   End
   Begin VB.Label Label6 
      Caption         =   "Kuantitas"
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Keterangan"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   4080
      Width           =   1335
   End
End
Attribute VB_Name = "FrmSuratJalan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub bersih()
    ' Untuk membersihkan TextBox dan ComboBox
    For Each X In Me
        If TypeOf X Is TextBox Then
            X.Text = ""
        End If
    Next
    cmbKodePelanggan.Clear
    cmbKodeBarang.Clear
End Sub

Sub txtHidup()
    ' Untuk mengaktifkan TextBox dan ComboBox
    For Each X In Me
        If TypeOf X Is TextBox Then
            X.Enabled = True
        End If
    Next
    txtNomorSuratJalan.Enabled = False
    cmbKodePelanggan.Enabled = True
    cmbKodeBarang.Enabled = True
End Sub

Sub txtMati()
    ' Untuk menonaktifkan TextBox dan ComboBox
    For Each X In Me
        If TypeOf X Is TextBox Then
            X.Enabled = False
        End If
    Next
    txtCari.Enabled = True
    cmbKodePelanggan.Enabled = False
    cmbKodeBarang.Enabled = False
End Sub

Sub noSuratJalanOtomatis()
    On Error Resume Next
    Dim nomor As String
    ' Mengurutkan field No_sj pada tabel surat_jalan
    adoSuratJalan.Recordset.Sort = "No_sj"
    ' Menampilkan semua record pada tabel surat_jalan
    adoSuratJalan.RecordSource = "select * from surat_jalan"
    With adoSuratJalan.Recordset
        If .RecordCount = 0 Then
            nomor = "SJ" + Format(Date, "YYMMDD") + "001"
        Else
            .MoveLast
            If Mid(!No_sj, 3, 6) <> Format(Date, "YYMMDD") Then
                nomor = "SJ" + Format(Date, "YYMMDD") + "001"
            Else
                hitung = Right(!No_sj, 3) + 1
                nomor = "SJ" + Format(Date, "YYMMDD") + Right("000" & hitung, 3)
            End If
        End If
    End With
    txtNomorSuratJalan.Text = nomor
    On Error GoTo 0
End Sub

Sub simpanSuratJalan()
    On Error Resume Next
    With adoSuratJalan.Recordset
        .AddNew
            !No_sj = txtNomorSuratJalan.Text
            !Kode_pelanggan = cmbKodePelanggan.Text
            !Kode_barang = cmbKodeBarang.Text
            !Kuantitas = txtKuantitas.Text
            !Keterangan_sj = txtKeterangan.Text
        .Update
    End With
    On Error GoTo 0
End Sub

Sub ambilDataKodePelanggan()
    On Error Resume Next
    koneksi
    Set rsPelanggan = New ADODB.Recordset
    rsPelanggan.Open "select Kode_pelanggan from pelanggan group by Kode_pelanggan having count(*) >= 1", kon
    Do While Not rsPelanggan.EOF
        cmbKodePelanggan.AddItem rsPelanggan!Kode_pelanggan
        rsPelanggan.MoveNext
    Loop
    kon.Close
    On Error GoTo 0
End Sub

Sub ambilDataKodeBarang()
    On Error Resume Next
    koneksi
    Set rsBarang = New ADODB.Recordset
    rsBarang.Open "select Kode_barang from barang group by Kode_barang having count(*) >= 1", kon
    Do While Not rsBarang.EOF
        cmbKodeBarang.AddItem rsBarang!Kode_barang
        rsBarang.MoveNext
    Loop
    kon.Close
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
    adoSuratJalan.Recordset.MoveFirst
    adoSuratJalan.Recordset.Find "No_sj='" & txtCari.Text & "'", , adSearchForward
    If adoSuratJalan.Recordset.EOF Then
        MsgBox "Nomor Surat Jalan Tidak Ada!", vbOKOnly, "Informasi"
    End If
    On Error GoTo 0
End Sub

Private Sub btnHapus_Click()
    On Error Resume Next
    a = MsgBox("Yakin Mau Dihapus  ???", vbYesNo + vbInformation, "Konfirmasi")
    If a = vbYes Then
        adoSuratJalan.Recordset.Delete
        txtMati
        bersih
    End If
    On Error GoTo 0
End Sub

Private Sub btnKeluar_Click()
    Unload Me
End Sub

Private Sub btnSimpan_Click()
    simpanSuratJalan
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
    noSuratJalanOtomatis
    ambilDataKodePelanggan
    ambilDataKodeBarang
End Sub

Private Sub dtgSuratJalan_DblClick()
    On Error Resume Next
        adoSuratJalan.Recordset.Delete
    On Error GoTo 0
End Sub

Private Sub dtgSuratJalan_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    On Error GoTo 0
End Sub

Private Sub Form_Activate()
    txtMati
    txtNomorSuratJalan.Enabled = False
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

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        btnSimpan.SetFocus
    End If
End Sub

Private Sub txtKuantitas_Change()
    If Len(txtKuantitas.Text) > 0 Then
        If Not IsNumeric(Right(txtKuantitas.Text, 1)) Then
            txtKuantitas.Text = ""
            txtKuantitas.SetFocus
        End If
    End If
End Sub

Private Sub txtKuantitas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtKeterangan.SetFocus
    End If
End Sub
