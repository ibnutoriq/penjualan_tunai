VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmPelanggan 
   BackColor       =   &H00FFFF00&
   Caption         =   "Pelanggan"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14010
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   14010
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFax 
      Height          =   495
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   5
      Top             =   4320
      Width           =   2100
   End
   Begin VB.TextBox txtNoTelpon 
      Height          =   495
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   4
      Top             =   3600
      Width           =   2100
   End
   Begin VB.TextBox txtAlamat 
      Height          =   855
      Left            =   1800
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2625
      Width           =   2300
   End
   Begin VB.TextBox txtKodePelanggan 
      Height          =   495
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtNamaPelanggan 
      Height          =   495
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   2
      Top             =   1920
      Width           =   2800
   End
   Begin VB.CommandButton btnTambah 
      Caption         =   "&Tambah"
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton btnKeluar 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   5040
      TabIndex        =   12
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton btnHapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   3840
      TabIndex        =   11
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton btnSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton btnBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   2640
      TabIndex        =   10
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pencarian"
      Height          =   855
      Left            =   9600
      TabIndex        =   0
      Top             =   5160
      Width           =   4095
      Begin VB.TextBox txtCari 
         Height          =   495
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   1300
      End
      Begin VB.CommandButton btnCari 
         Caption         =   "&Cari"
         Height          =   495
         Left            =   3000
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Kode Pelanggan"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid dtgPelanggan 
      Bindings        =   "FrmPelanggan.frx":0000
      Height          =   3615
      Left            =   4920
      TabIndex        =   14
      Top             =   1200
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6376
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
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "Alamat"
         Caption         =   "Alamat"
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
         DataField       =   "No_telpon"
         Caption         =   "No_telpon"
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
         DataField       =   "Fax"
         Caption         =   "Fax"
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
            ColumnWidth     =   1379,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1890,142
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1574,929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1844,787
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoPelanggan 
      Height          =   495
      Left            =   11280
      Top             =   4320
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
      RecordSource    =   "select * from pelanggan"
      Caption         =   "Ado Pelanggan"
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
   Begin VB.Label Label7 
      Caption         =   "Fax"
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "No Telpon"
      Height          =   495
      Left            =   240
      TabIndex        =   19
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Pelanggan"
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
      TabIndex        =   18
      Top             =   240
      Width           =   13455
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Pelanggan"
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Kode Pelanggan"
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Alamat"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   2640
      Width           =   1335
   End
End
Attribute VB_Name = "FrmPelanggan"
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
            txtKodePelanggan.Enabled = False
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

Sub tambahPelanggan()
    ' Menambahkan Record pada field Kode_pelanggan secara otomatis dan setiap record nilainya bertambah +1
    On Error Resume Next
    adoPelanggan.Recordset.Sort = "Kode_pelanggan"
    adoPelanggan.RecordSource = "select * from pelanggan"
    Dim Urutan As String * 6 'lebar data 11 karakter
    Dim hitung As Long
    With adoPelanggan.Recordset
        If .RecordCount = 0 Then
            Urutan = "PLG" + "001"
            txtuser = Urutan
        Else
            .MoveLast
            If Mid(!Kode_pelanggan, 4, 3) = "" Then
                Urutan = "PLG" + "001"
            Else
                hitung = Right(!Kode_pelanggan, 3) + 1
                Urutan = "PLG" + Right("00" & hitung, 3)
            End If
        End If
        'menampilkan penomoran pada textbox
        txtKodePelanggan.Text = Urutan
    End With
    On Error GoTo 0
End Sub

Sub simpanPelanggan()
    On Error Resume Next
    With adoPelanggan.Recordset
        .AddNew
            !Kode_pelanggan = txtKodePelanggan.Text
            !Nama_pelanggan = txtNamaPelanggan.Text
            !Alamat = txtAlamat.Text
            !No_telpon = txtNoTelpon.Text
            !Fax = txtFax.Text
        .Update
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
    adoPelanggan.Recordset.MoveFirst
    adoPelanggan.Recordset.Find "Kode_pelanggan='" & txtCari.Text & "'", , adSearchForward
    If adoPelanggan.Recordset.EOF Then
        MsgBox "Kode Pelanggan Tidak Ada!", vbOKOnly, "Informasi"
    End If
    On Error GoTo 0
End Sub

Private Sub btnHapus_Click()
    On Error Resume Next
    a = MsgBox("Yakin Mau Dihapus  ???", vbYesNo + vbInformation, "Konfirmasi")
    If a = vbYes Then
        adoPelanggan.Recordset.Delete
        txtMati
        bersih
    End If
    On Error GoTo 0
End Sub

Private Sub btnKeluar_Click()
    Unload Me
End Sub

Private Sub btnSimpan_Click()
    simpanPelanggan
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
    tambahPelanggan
    txtNamaPelanggan.SetFocus
End Sub

Private Sub dtgPelanggan_DblClick()
    On Error Resume Next
        adoPelanggan.Recordset.Delete
    On Error GoTo 0
End Sub

Private Sub dtgPelanggan_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    On Error GoTo 0
End Sub

Private Sub Form_Activate()
    txtMati
    txtKodePelanggan.Enabled = False
    btnTambah.SetFocus
    btnSimpan.Enabled = False
    btnBatal.Enabled = False
End Sub

Private Sub txtAlamat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNoTelpon.SetFocus
    End If
End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        btnCari.SetFocus
    End If
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        btnSimpan.SetFocus
    End If
End Sub

Private Sub txtNamaPelanggan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAlamat.SetFocus
    End If
End Sub

Private Sub txtNoTelpon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFax.SetFocus
    End If
End Sub
