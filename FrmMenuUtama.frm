VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FrmMenuUtama 
   Caption         =   "Menu Utama"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stMenuUtama 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2655
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Image imgMenutUtama 
      Height          =   2655
      Left            =   0
      Picture         =   "FrmMenuUtama.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
   Begin VB.Menu mnHakAkses 
      Caption         =   "Hak Akses"
      Begin VB.Menu smnLogin 
         Caption         =   "Login"
         Shortcut        =   {F1}
      End
      Begin VB.Menu smnLogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu smnKeluar 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu mnMaster 
      Caption         =   "Master"
      Begin VB.Menu smnPengguna 
         Caption         =   "Pengguna"
         Shortcut        =   {F2}
      End
      Begin VB.Menu smnBarang 
         Caption         =   "Barang"
         Shortcut        =   {F3}
      End
      Begin VB.Menu smnPelanggan 
         Caption         =   "Pelanggan"
         Shortcut        =   {F4}
      End
      Begin VB.Menu smnPerkiraan 
         Caption         =   "Perkiraan"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnPenjualan 
      Caption         =   "Penjualan"
      Begin VB.Menu smnTransaksi 
         Caption         =   "Transaksi"
         Shortcut        =   {F6}
      End
      Begin VB.Menu smnInvoice 
         Caption         =   "Invoice"
         Shortcut        =   {F7}
      End
      Begin VB.Menu smnSuratJalan 
         Caption         =   "Surat Jalan"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnJurnal 
      Caption         =   "Jurnal"
   End
End
Attribute VB_Name = "FrmMenuUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    mnMaster.Enabled = False
    mnPenjualan.Enabled = False
    mnJurnal.Enabled = False
    smnLogout.Enabled = False
End Sub

Private Sub Form_Resize()
    imgMenutUtama.Width = Width
    imgMenutUtama.Height = Height
End Sub

Private Sub mnJurnal_Click()
    FrmJurnal.Show
End Sub

Private Sub smnBarang_Click()
    FrmBarang.Show
End Sub

Private Sub smnInvoice_Click()
    FrmInvoice.Show
End Sub

Private Sub smnKeluar_Click()
    End
End Sub

Private Sub smnLogin_Click()
    FrmLogin.Show
End Sub

Private Sub smnLogout_Click()
    mnMaster.Enabled = False
    mnPenjualan.Enabled = False
    mnJurnal.Enabled = False
    smnLogout.Enabled = False
    smnLogin.Enabled = True
    smnPengguna.Enabled = True
    stMenuUtama.Panels(1) = ""
    stMenuUtama.Panels(2) = ""
    stMenuUtama.Panels(3) = ""
End Sub

Private Sub smnPelanggan_Click()
    FrmPelanggan.Show
End Sub

Private Sub smnPengguna_Click()
    FrmPengguna.Show
End Sub

Private Sub smnPerkiraan_Click()
    FrmPerkiraan.Show
End Sub

Private Sub smnSuratJalan_Click()
    FrmSuratJalan.Show
End Sub

Private Sub smnTransaksi_Click()
    FrmTransaksi.Show
End Sub
