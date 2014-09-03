VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmTransaksi 
   BackColor       =   &H00FFFF00&
   Caption         =   "Perkiraan"
   ClientHeight    =   8775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport crTransaksi 
      Left            =   7080
      Top             =   8040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   7320
      TabIndex        =   18
      Top             =   3360
      Width           =   4455
      Begin VB.TextBox txtSpesifikasi 
         Height          =   495
         Left            =   1560
         TabIndex        =   26
         Top             =   960
         Width           =   2800
      End
      Begin VB.TextBox txtKeterangan 
         Height          =   495
         Left            =   1560
         TabIndex        =   25
         Top             =   240
         Width           =   2800
      End
      Begin VB.Label Label3 
         Caption         =   "Spesifikasi"
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Keterangan"
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   5280
      TabIndex        =   17
      Top             =   960
      Width           =   3735
      Begin VB.TextBox txtHargaSatuan 
         Height          =   495
         Left            =   1680
         TabIndex        =   21
         Top             =   240
         Width           =   1800
      End
      Begin VB.TextBox txtTotal 
         Height          =   495
         Left            =   1680
         TabIndex        =   20
         Top             =   1680
         Width           =   1800
      End
      Begin VB.TextBox txtKuantitas 
         Height          =   495
         Left            =   1680
         TabIndex        =   19
         Top             =   960
         Width           =   1800
      End
      Begin VB.Label Label2 
         Caption         =   "Kuantitas"
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Total"
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Harga Satuan"
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton btnCetak 
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   3840
      TabIndex        =   16
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton btnTambah 
      Caption         =   "&Tambah"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton btnKeluar 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   5640
      TabIndex        =   14
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton btnSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   1440
      TabIndex        =   13
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton btnBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   2640
      TabIndex        =   12
      Top             =   8040
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid dtgTmpTransaksi 
      Bindings        =   "FrmTransaksi.frx":0000
      Height          =   2415
      Left            =   240
      TabIndex        =   11
      Top             =   5280
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   4260
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "No_po"
         Caption         =   "No_po"
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
         DataField       =   "Tanggal"
         Caption         =   "Tanggal"
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
      BeginProperty Column03 
         DataField       =   "Nama_pelanggan"
         Caption         =   "Nama_pelanggan"
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
      BeginProperty Column05 
         DataField       =   "Harga_satuan"
         Caption         =   "Harga_satuan"
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
      BeginProperty Column06 
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
      BeginProperty Column07 
         DataField       =   "Total"
         Caption         =   "Total"
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
      BeginProperty Column08 
         DataField       =   "Keterangan"
         Caption         =   "Keterangan"
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
      BeginProperty Column09 
         DataField       =   "Spesifikasi"
         Caption         =   "Spesifikasi"
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
            ColumnWidth     =   1470,047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbKodePelanggan 
      Height          =   315
      Left            =   1800
      TabIndex        =   10
      Top             =   2520
      Width           =   1455
   End
   Begin VB.ComboBox cmbKodeBarang 
      Height          =   315
      Left            =   1800
      TabIndex        =   9
      Top             =   3960
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtpTanggal 
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      Format          =   179372033
      CurrentDate     =   41622
   End
   Begin VB.TextBox txtNamaPelanggan 
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   3240
      Width           =   2800
   End
   Begin VB.TextBox txtNoPO 
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   1080
      Width           =   1800
   End
   Begin MSAdodcLib.Adodc adoTransaksi 
      Height          =   495
      Left            =   9000
      Top             =   7200
      Width           =   2775
      _ExtentX        =   4895
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
      RecordSource    =   "select * from transaksi"
      Caption         =   "Ado Transaksi"
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
   Begin MSAdodcLib.Adodc adoTmpTransaksi 
      Height          =   495
      Left            =   6120
      Top             =   7200
      Width           =   2775
      _ExtentX        =   4895
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
      RecordSource    =   "select * from tmp_transaksi"
      Caption         =   "Ado Tmp Transaksi"
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
   Begin VB.Label Label11 
      Caption         =   "No PO"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Tanggal"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Kode Pelanggan"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Nama Pelanggan"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Kode Barang"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Transaksi"
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
      TabIndex        =   0
      Top             =   240
      Width           =   11535
   End
End
Attribute VB_Name = "FrmTransaksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub bersih()
    ' Untuk membersihkan TextBox
    For Each X In Me
        If TypeOf X Is TextBox Then
            X.Text = ""
        End If
    Next
End Sub

Sub txtHidup()
    ' Untuk mengaktifkan item
    For Each X In Me
        If TypeOf X Is TextBox Then
            X.Enabled = True
        End If
    Next
    txtNoPO.Enabled = False
    dtpTanggal.Enabled = True
    cmbKodePelanggan.Enabled = True
    cmbKodeBarang.Enabled = True
End Sub

Sub txtMati()
    'Untuk menonaktifkan item
    For Each X In Me
        If TypeOf X Is TextBox Then
            X.Enabled = False
        End If
    Next
    dtpTanggal.Enabled = False
    cmbKodePelanggan.Enabled = False
    cmbKodeBarang.Enabled = False
End Sub

Sub hapusTmpTransaksi()
    On Error Resume Next
    koneksi
    qry = "delete from tmp_transaksi"
    Set rsTmpTransaksi = kon.Execute(qry, , adCmdText)
    kon.Close
    adoTmpTransaksi.Refresh
    On Error GoTo 0
End Sub

Sub noPOOtomatis()
    On Error Resume Next
    Dim nomor As String
    koneksi
    qry = "select No_po from transaksi order by No_po DESC"
    Set rsTransaksi = kon.Execute(qry, , adCmdText)
    With rsTransaksi
        If .RecordCount = 0 Then
            nomor = "TR" + Format(Date, "YYMMDD") + "001"
        Else
            If Mid(!No_po, 3, 6) <> Format(Date, "YYMMDD") Then
                nomor = "TR" + Format(Date, "YYMMDD") + "001"
            Else
                hitung = Right(!No_po, 3) + 1
                nomor = "TR" + Format(Date, "YYMMDD") + Right("000" & hitung, 3)
            End If
        End If
    End With
    txtNoPO.Text = nomor
    kon.Close
    On Error GoTo 0
End Sub

Sub simpanTmpTransaksi()
    On Error Resume Next
    With adoTmpTransaksi.Recordset
        .AddNew
            !No_po = txtNoPO.Text
            !Tanggal = dtpTanggal.Value
            !Kode_pelanggan = cmbKodePelanggan.Text
            !Nama_pelanggan = txtNamaPelanggan.Text
            !Kode_barang = cmbKodeBarang.Text
            !Harga_satuan = txtHargaSatuan.Text
            !Kuantitas = txtKuantitas.Text
            !Total = txtTotal.Text
            !Keterangan = txtKeterangan.Text
            !Spesifikasi = txtSpesifikasi.Text
        .Update
    End With
    On Error GoTo 0
End Sub

Sub salinTmpTransaksiKeTransaksi()
    On Error Resume Next
    koneksi
    qry = "insert into transaksi select * from tmp_transaksi"
    Set rsTransaksi = kon.Execute(qry, , adCmdText)
    kon.Close
    On Error GoTo 0
End Sub

Sub updateMinBarang()
    On Error Resume Next
    koneksi
    qry = "update barang set Stok=Stok-'" & txtKuantitas.Text & "' where Kode_barang='" & cmbKodeBarang.Text & "'"
    Set rsBarang = kon.Execute(qry, , adCmdText)
    conn.Close
    On Error GoTo 0
End Sub

Sub updatePlusBarang()
    On Error Resume Next
    koneksi
    qry = "update barang set Stok = Stok + '" & dtgTmpTransaksi.Columns(6) & "' where Kode_barang = '" & dtgTmpTransaksi.Columns(4) & "'"
    Set rsBarang = kon.Execute(qry, , adCmdText)
    kon.Close
    adoTmpTransaksi.Recordset.Delete
    On Error GoTo 0
End Sub

Sub ambilDataKodePelanggan()
    On Error Resume Next
    cmbKodePelanggan.Clear
    koneksi
    rsPelanggan.Open "select Kode_pelanggan from pelanggan order by Kode_pelanggan ASC", kon
    Do While Not rsPelanggan.EOF
        cmbKodePelanggan.AddItem rsPelanggan!Kode_pelanggan
        rsPelanggan.MoveNext
    Loop
    kon.Close
    On Error GoTo 0
End Sub

Sub ambilDataKodeBarang()
    On Error Resume Next
    cmbKodeBarang.Clear
    koneksi
    rsBarang.Open "select Kode_barang from barang order by Kode_barang ASC", kon
    Do While Not rsBarang.EOF
        cmbKodeBarang.AddItem rsBarang!Kode_barang
        rsBarang.MoveNext
    Loop
    kon.Close
    On Error GoTo 0
End Sub

Sub tampilDataKodePelanggan()
    On Error Resume Next
    koneksi
    rsPelanggan.Open "select * from pelanggan where Kode_pelanggan = '" + cmbKodePelanggan.Text + "'", kon
    With rsPelanggan
        txtNamaPelanggan.Text = rsPelanggan!Nama_pelanggan
    End With
    kon.Close
    On Error GoTo 0
End Sub

Sub tampilDataKodeBarang()
    On Error Resume Next
    koneksi
    rsBarang.Open "select * from barang where Kode_barang = '" + cmbKodeBarang.Text + "'", kon
    With rsBarang
        txtHargaSatuan.Text = rsBarang!Harga_satuan
    End With
    kon.Close
    On Error GoTo 0
End Sub

Sub cetakTransaksi()
    On Error Resume Next
    With crTransaksi
        .ReportFileName = App.Path & "\Laporan\lapTransaksi.rpt"
        .SelectionFormula = ""
        .ParameterFields(0) = "formulaNoPO;" & txtNoPO.Text & ";True"
        .ParameterFields(1) = "formulaKodePelanggan;" & cmbKodePelanggan.Text & ";True"
        .ParameterFields(2) = "formulaNamaPelanggan;" & txtNamaPelanggan.Text & ";True"
        .RetrieveDataFiles
        .WindowState = crptMaximized
        .Action = 1
    End With
    On Error GoTo 0
End Sub

Private Sub btnBatal_Click()
    bersih
    txtMati
    btnBatal.Enabled = False
    btnTambah.Enabled = True
    btnSimpan.Enabled = False
End Sub

Private Sub btnCetak_Click()
    salinTmpTransaksiKeTransaksi
    cetakTransaksi
    hapusTmpTransaksi
    bersih
    txtMati
    btnSimpan.Enabled = False
    btnTambah.Enabled = True
    btnBatal.Enabled = False
End Sub

Private Sub btnKeluar_Click()
    Unload Me
End Sub

Private Sub btnSimpan_Click()
    simpanTmpTransaksi
    updateMinBarang
    cmbKodeBarang.Clear
    ambilDataKodeBarang
    txtHargaSatuan.Text = ""
    txtKuantitas.Text = ""
    txtTotal.Text = ""
End Sub

Private Sub btnTambah_Click()
    txtHidup
    bersih
    btnTambah.Enabled = True
    btnBatal.Enabled = True
    btnSimpan.Enabled = True
    noPOOtomatis
    ambilDataKodePelanggan
    ambilDataKodeBarang
End Sub

Private Sub cmbKodeBarang_Click()
    tampilDataKodeBarang
    txtKuantitas.SetFocus
End Sub

Private Sub cmbKodePelanggan_Click()
    tampilDataKodePelanggan
End Sub

Private Sub dtgTmpTransaksi_DblClick()
    updatePlusBarang
End Sub

Private Sub Form_Activate()
    txtMati
    txtNoPO.Enabled = False
    btnTambah.SetFocus
    btnSimpan.Enabled = False
    btnBatal.Enabled = False
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSpesifikasi.SetFocus
    End If
End Sub

Private Sub txtKuantitas_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        txtTotal.Text = Val(txtHargaSatuan.Text * txtKuantitas.Text)
        txtKeterangan.SetFocus
    End If
    On Error GoTo 0
End Sub

Private Sub txtSpesifikasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        btnSimpan.SetFocus
    End If
End Sub
