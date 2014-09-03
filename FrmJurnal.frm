VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmJurnal 
   BackColor       =   &H00FFFF00&
   Caption         =   "Jurnal"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   15075
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker dtpTanggal 
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      Format          =   178061313
      CurrentDate     =   41622
   End
   Begin VB.ComboBox cmbKodeAkun 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtKredit 
      Height          =   495
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   6
      Top             =   4800
      Width           =   1800
   End
   Begin VB.TextBox txtKeterangan 
      Height          =   495
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   7
      Top             =   5520
      Width           =   2800
   End
   Begin VB.TextBox txtDebit 
      Height          =   495
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   5
      Top             =   4080
      Width           =   1800
   End
   Begin VB.TextBox txtNamaAkun 
      Height          =   495
      Left            =   1800
      MaxLength       =   25
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2640
      Width           =   2800
   End
   Begin VB.TextBox txtNoJurnal 
      Height          =   495
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton btnTambah 
      Caption         =   "&Tambah"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton btnKeluar 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   5040
      TabIndex        =   14
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton btnHapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   3840
      TabIndex        =   13
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton btnSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton btnBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   2640
      TabIndex        =   12
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pencarian"
      Height          =   855
      Left            =   10680
      TabIndex        =   0
      Top             =   6480
      Width           =   4095
      Begin VB.TextBox txtCari 
         Height          =   495
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   1300
      End
      Begin VB.CommandButton btnCari 
         Caption         =   "&Cari"
         Height          =   495
         Left            =   3000
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "No Jurnal"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid dtgJurnal 
      Bindings        =   "FrmJurnal.frx":0000
      Height          =   4815
      Left            =   4920
      TabIndex        =   16
      Top             =   1200
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8493
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "No_jurnal"
         Caption         =   "No_jurnal"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "Debit"
         Caption         =   "Debit"
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
         DataField       =   "Kredit"
         Caption         =   "Kredit"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   884,976
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoJurnal 
      Height          =   495
      Left            =   11280
      Top             =   5520
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
      RecordSource    =   "select * from jurnal"
      Caption         =   "Ado Jurnal"
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
   Begin VB.Label Label9 
      Caption         =   "Kredit"
      Height          =   495
      Left            =   240
      TabIndex        =   24
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Keterangan"
      Height          =   495
      Left            =   240
      TabIndex        =   23
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Debit"
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Tanggal"
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Jurnal"
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
      TabIndex        =   20
      Top             =   240
      Width           =   14535
   End
   Begin VB.Label Label2 
      Caption         =   "Kode Akun"
      Height          =   495
      Left            =   240
      TabIndex        =   19
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "No Jurnal"
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Nama Akun"
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   2640
      Width           =   1335
   End
End
Attribute VB_Name = "FrmJurnal"
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
    cmbKodeAkun.Clear
    dtpTanggal.Value = Now
End Sub

Sub txtHidup()
    For Each X In Me
        If TypeOf X Is TextBox Then
            X.Enabled = True
        End If
    Next
    txtNoJurnal.Enabled = False
    dtpTanggal.Enabled = True
    cmbKodeAkun.Enabled = True
End Sub

Sub txtMati()
    For Each X In Me
        If TypeOf X Is TextBox Then
            X.Enabled = False
        End If
    Next
    dtpTanggal.Enabled = False
    cmbKodeAkun.Enabled = False
End Sub

Sub noAutomatis()
    On Error Resume Next
    Dim nmr As String
    adoJurnal.Recordset.Sort = "No_jurnal"
    adoJurnal.RecordSource = "select * from jurnal"
    With adoJurnal.Recordset
        If .RecordCount = 0 Then
            nmr = "1"
        Else
            .MoveLast
            htg = !No_jurnal + 1
            nmr = htg
        End If
    End With
    txtNoJurnal.Text = nmr
    On Error GoTo 0
End Sub

Sub simpanJurnal()
    On Error Resume Next
    With adoJurnal.Recordset
        .AddNew
            !No_jurnal = txtNoJurnal.Text
            !Kode_akun = cmbKodeAkun.Text
            !Nama_akun = txtNamaAkun.Text
            !Tanggal = dtpTanggal.Value
            !Debit = txtDebit.Text
            !Kredit = txtKredit.Text
            !Keterangan = txtKeterangan.Text
        .Update
    End With
    On Error GoTo 0
End Sub


Sub ambilDataKodeAkun()
    On Error Resume Next
    koneksi
    Set rsPerkiraan = New ADODB.Recordset
    rsPerkiraan.Open "select Kode_akun from perkiraan group by Kode_akun having count(*) >= 1", kon
    Do While Not rsPerkiraan.EOF
        cmbKodeAkun.AddItem rsPerkiraan!Kode_akun
        rsPerkiraan.MoveNext
    Loop
    kon.Close
    On Error GoTo 0
End Sub

Sub tampilNamaAkun()
    On Error Resume Next
    koneksi
    Set rsPerkiraan = New ADODB.Recordset
    rsPerkiraan.Open "select Nama_akun from perkiraan where Kode_akun = '" & cmbKodeAkun.Text & "'", kon
    With rsPerkiraan
        txtNamaAkun.Text = rsPerkiraan!Nama_akun
    End With
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
    adoJurnal.Recordset.MoveFirst
    adoJurnal.Recordset.Find "No_jurnal='" & txtCari.Text & "'", , adSearchForward
    If adoJurnal.Recordset.EOF Then
        MsgBox "No Jurnal Tidak Ada!", vbOKOnly, "Informasi"
    End If
    On Error GoTo 0
End Sub

Private Sub btnHapus_Click()
    On Error Resume Next
    a = MsgBox("Yakin Mau Dihapus  ???", vbYesNo + vbInformation, "Konfirmasi")
    If a = vbYes Then
        adoJurnal.Recordset.Delete
        txtMati
        bersih
    End If
    On Error GoTo 0
End Sub

Private Sub btnKeluar_Click()
    Unload Me
End Sub

Private Sub btnSimpan_Click()
    simpanJurnal
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
    noAutomatis
    ambilDataKodeAkun
End Sub

Private Sub cmbKodeAkun_Click()
    tampilNamaAkun
    txtDebit.SetFocus
End Sub

Private Sub dtgJurnal_DblClick()
    On Error Resume Next
        adoJurnal.Recordset.Delete
    On Error GoTo 0
End Sub

Private Sub dtgJurnal_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    On Error GoTo 0
End Sub

Private Sub Form_Activate()
    txtMati
    txtNoJurnal.Enabled = False
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

Private Sub txtDebit_Change()
    If Len(txtDebit.Text) > 0 Then
        If Not IsNumeric(Right(txtDebit.Text, 1)) Then
            txtDebit.Text = ""
            txtDebit.SetFocus
        End If
    End If
End Sub

Private Sub txtDebit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtKredit.SetFocus
    End If
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        btnSimpan.SetFocus
    End If
End Sub

Private Sub txtKredit_Change()
    If Len(txtKredit.Text) > 0 Then
        If Not IsNumeric(Right(txtKredit.Text, 1)) Then
            txtKredit.Text = ""
            txtKredit.SetFocus
        End If
    End If
End Sub

Private Sub txtKredit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtKeterangan.SetFocus
    End If
End Sub
