VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmPengguna 
   BackColor       =   &H00FFFF00&
   Caption         =   "Pengguna"
   ClientHeight    =   5115
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   12345
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbHakAkses 
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pencarian"
      Height          =   855
      Left            =   8040
      TabIndex        =   14
      Top             =   3960
      Width           =   3975
      Begin VB.CommandButton btnCari 
         Caption         =   "Cari"
         Height          =   495
         Left            =   2880
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtCari 
         Height          =   495
         Left            =   1560
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "User ID"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid dtgPengguna 
      Bindings        =   "FrmPengguna.frx":0000
      Height          =   2415
      Left            =   4800
      TabIndex        =   13
      Top             =   1200
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4260
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "User_id"
         Caption         =   "User_id"
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
         DataField       =   "Kata_kunci"
         Caption         =   "Kata_kunci"
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
         DataField       =   "Nama_pengguna"
         Caption         =   "Nama_pengguna"
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
         DataField       =   "Hak_akses"
         Caption         =   "Hak_akses"
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
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2684,977
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2745,071
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoPengguna 
      Height          =   495
      Left            =   7920
      Top             =   3120
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
   Begin VB.TextBox txtKataKunci 
      Height          =   495
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton btnHapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   2520
      TabIndex        =   11
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton btnTambah 
      Caption         =   "&Tambah"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton btnBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   3720
      TabIndex        =   9
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton btnKeluar 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   4920
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton btnSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtNamaPengguna 
      Height          =   495
      Left            =   1680
      MaxLength       =   25
      TabIndex        =   3
      Top             =   2640
      Width           =   2800
   End
   Begin VB.TextBox txtUserID 
      Height          =   495
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Pengguna"
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
      Width           =   11775
   End
   Begin VB.Label Label5 
      Caption         =   "Kata Kunci"
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Nama Pengguna"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Hak Akses"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "User ID"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "FrmPengguna"
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
    ' Mengaktifkan TextBox pada Form
    For Each X In Me
        If TypeOf X Is TextBox Then
            X.Enabled = True
            txtUserID.Enabled = False
        End If
    Next
End Sub

Sub txtMati()
    ' Menonaktifkan TextBox pada Form
    For Each X In Me
        If TypeOf X Is TextBox Then
            X.Enabled = False
            txtCari.Enabled = True
        End If
    Next
End Sub

Sub itemHakAkses()
    ' Mengisi item pada cmbHakAkses
    cmbHakAkses.AddItem "User"
    cmbHakAkses.AddItem "Admin"
End Sub

Sub tambahPengguna()
    ' Menambahkan Record pada field Kode_pengguna secara otomatis dan setiap record nilainya bertambah +1
    On Error Resume Next
    adoPengguna.Recordset.Sort = "User_id"
    adoPengguna.RecordSource = "select * from pengguna"
    Dim Urutan As String * 6 'lebar data 11 karakter
    Dim hitung As Long
    With adoPengguna.Recordset
        If .RecordCount = 0 Then
            Urutan = "PGN" + "001"
            txtuser = Urutan
        Else
            .MoveLast
            If Mid(!User_id, 4, 3) = "" Then
                Urutan = "PGN" + "001"
            Else
                hitung = Right(!User_id, 3) + 1
                Urutan = "PGN" + Right("00" & hitung, 3)
            End If
        End If
        'menampilkan penomoran pada textbox
        txtUserID.Text = Urutan
    End With
    On Error GoTo 0
End Sub

Sub simpanPengguna()
    On Error Resume Next
    With adoPengguna.Recordset
        .AddNew
            !User_id = txtUserID.Text
            !Kata_kunci = txtKataKunci.Text
            !Nama_pengguna = txtNamaPengguna.Text
            !Hak_akses = cmbHakAkses.Text
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
    adoPengguna.Recordset.MoveFirst
    adoPengguna.Recordset.Find "User_id='" & txtCari.Text & "'", , adSearchForward
    If adoPengguna.Recordset.EOF Then
        MsgBox "User ID Tidak Ada!", vbOKOnly, "Informasi"
    End If
    On Error GoTo 0
End Sub

Private Sub btnHapus_Click()
    On Error Resume Next
    a = MsgBox("Yakin Mau Dihapus  ???", vbYesNo + vbInformation, "Konfirmasi")
    If a = vbYes Then
        adoPengguna.Recordset.Delete
        txtMati
        bersih
    End If
    On Error GoTo 0
End Sub

Private Sub btnKeluar_Click()
    Unload Me
End Sub

Private Sub btnSimpan_Click()
    simpanPengguna
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
    txtKataKunci.SetFocus
    tambahPengguna
    cmbHakAkses.Clear
    itemHakAkses
End Sub

Private Sub dtgPengguna_DblClick()
    On Error Resume Next
    adoPengguna.Recordset.Delete
    On Error GoTo 0
End Sub

Private Sub dtgPengguna_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    On Error GoTo 0
End Sub

Private Sub Form_Activate()
    txtMati
    txtUserID.Enabled = False
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

Private Sub txtKataKunci_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNamaPengguna.SetFocus
    End If
End Sub

Private Sub txtNamaPengguna_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbHakAkses.SetFocus
    End If
End Sub
