VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmBarang 
   BackColor       =   &H00FFFF00&
   Caption         =   "Barang"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14010
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   14010
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtStok 
      Height          =   495
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   20
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pencarian"
      Height          =   855
      Left            =   9600
      TabIndex        =   9
      Top             =   4920
      Width           =   4095
      Begin VB.CommandButton btnCari 
         Caption         =   "&Cari"
         Height          =   495
         Left            =   3000
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtCari 
         Height          =   495
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   1300
      End
      Begin VB.Label Label5 
         Caption         =   "Kode Barang"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton btnBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton btnSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton btnHapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton btnKeluar 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   5040
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton btnTambah 
      Caption         =   "&Tambah"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtNamaBarang 
      Height          =   495
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   1
      Top             =   1920
      Width           =   2800
   End
   Begin VB.TextBox txtKodeBarang 
      Height          =   495
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtHargaSatuan 
      Height          =   495
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   2
      Top             =   3360
      Width           =   2100
   End
   Begin VB.TextBox txtSatuan 
      Height          =   495
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid dtgBarang 
      Bindings        =   "FrmBarang.frx":0000
      Height          =   3375
      Left            =   4920
      TabIndex        =   13
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
      BeginProperty Column01 
         DataField       =   "Nama_barang"
         Caption         =   "Nama_barang"
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
         DataField       =   "Stok"
         Caption         =   "Stok"
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
      BeginProperty Column04 
         DataField       =   "Satuan"
         Caption         =   "Satuan"
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
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3000,189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1620,284
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoBarang 
      Height          =   495
      Left            =   11280
      Top             =   4080
      Visible         =   0   'False
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
      RecordSource    =   "select * from barang"
      Caption         =   "Ado Barang"
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
      Caption         =   "Stok"
      Height          =   495
      Left            =   240
      TabIndex        =   19
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Kode Barang"
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Barang"
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Barang"
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
      TabIndex        =   16
      Top             =   240
      Width           =   13455
   End
   Begin VB.Label Label6 
      Caption         =   "Harga Satuan"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Satuan"
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   4080
      Width           =   1335
   End
End
Attribute VB_Name = "FrmBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub bersih()
    ' Membersihkan TextBox pada Form
    For Each X In Me
        If TypeOf X Is TextBox Then
            X.Text = ""
        End If
    Next
End Sub

Sub txtHidup()
    ' Menonaktifkan TextBox pada Form
    For Each X In Me
        If TypeOf X Is TextBox Then
            X.Enabled = True
        End If
    Next
End Sub

Sub txtMati()
    ' Mengaktifkan TextBox pada Form
    For Each X In Me
        If TypeOf X Is TextBox Then
            X.Enabled = False
            txtCari.Enabled = True
        End If
    Next
End Sub

Sub simpanBarang()
    ' Membuat validasi pada field Kode_barang dan jika false maka simpan isi TextBox ke table barang
    On Error Resume Next
    With adoBarang.Recordset
        adoBarang.Recordset.Find "Kode_barang='" & txtKodeBarang.Text & "'", , adSearchForward
        If Not adoBarang.Recordset.EOF Then
            MsgBox "Kode Barang Sudah Terdaftar", vbOKOnly, "Informasi"
            Exit Sub
        Else
            .AddNew
                !Kode_barang = txtKodeBarang.Text
                !nama_barang = txtNamaBarang.Text
                !stok = txtStok.Text
                !Harga_satuan = txtHargaSatuan.Text
                !satuan = txtSatuan.Text
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
    ' Mencari Record pada table barang berdasarkan field Kode_barang
    On Error Resume Next
    adoBarang.Recordset.MoveFirst
    adoBarang.Recordset.Find "Kode_barang='" & txtCari.Text & "'", , adSearchForward
    If adoBarang.Recordset.EOF Then
        MsgBox "Kode Barang Tidak Ada!", vbOKOnly, "Informasi"
    End If
    On Error GoTo 0
End Sub

Private Sub btnHapus_Click()
    ' Menghapus Record dengan validasi
    On Error Resume Next
    a = MsgBox("Yakin Mau Dihapus  ???", vbYesNo + vbInformation, "Konfirmasi")
    If a = vbYes Then
        adoBarang.Recordset.Delete
        txtMati
        bersih
    End If
    On Error GoTo 0
End Sub

Private Sub btnKeluar_Click()
    Unload Me
End Sub

Private Sub btnSimpan_Click()
    simpanBarang
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
    txtKodeBarang.SetFocus
End Sub

Private Sub dtgBarang_DblClick()
    On Error Resume Next
        adoBarang.Recordset.Delete
    On Error GoTo 0
End Sub

Private Sub dtgBarang_KeyPress(KeyAscii As Integer)
    ' Mengaktifkan tombol Enter untuk navigasi
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

Private Sub txtHargaSatuan_Change()
    ' Hanya boleh input numeric
    If Len(txtHargaSatuan.Text) > 0 Then
        If Not IsNumeric(Right(txtHargaSatuan.Text, 1)) Then
            txtHargaSatuan.Text = ""
            txtHargaSatuan.SetFocus
        End If
    End If
End Sub

Private Sub txtHargaSatuan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSatuan.SetFocus
    End If
End Sub

Private Sub txtKodeBarang_KeyPress(KeyAscii As Integer)
    ' Membuat validasi pada field Kode_barang dan jika true maka bersih
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        adoBarang.Recordset.Find "Kode_barang='" & txtKodeBarang.Text & "'", , adSearchForward
        If Not adoBarang.Recordset.EOF Then
            MsgBox "Kode Barang Sudah Terdaftar", vbOKOnly, "Informasi"
            txtKodeBarang.Text = ""
            txtKodeBarang.SetFocus
            Exit Sub
            btnTambah.Enabled = True
            btnSimpan.Enabled = True
            btnBatal.Enabled = False
        End If
        txtNamaBarang.SetFocus
        btnSimpan.Enabled = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtNamaBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtStok.SetFocus
    End If
End Sub

Private Sub txtSatuan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        btnSimpan.SetFocus
    End If
End Sub

Private Sub txtStok_Change()
    ' Hanya boleh input numeric
    If Len(txtStok.Text) > 0 Then
        If Not IsNumeric(Right(txtStok.Text, 1)) Then
            txtStok.Text = ""
            txtStok.SetFocus
        End If
    End If
End Sub

Private Sub txtStok_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtHargaSatuan.SetFocus
    End If
End Sub
