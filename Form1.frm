VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Example C_Crypt"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8115
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCryptfile 
      Caption         =   "Encrypt / Decrypt file"
      Height          =   1335
      Left            =   120
      TabIndex        =   21
      Top             =   5040
      Width           =   7815
      Begin VB.CommandButton cmdFileDecrypt 
         Caption         =   "Decrypt"
         Height          =   288
         Left            =   6120
         TabIndex        =   44
         Top             =   960
         Width           =   1572
      End
      Begin VB.CommandButton cmdFileEncrypt 
         Caption         =   "Encrypt"
         Height          =   288
         Left            =   6120
         TabIndex        =   43
         Top             =   600
         Width           =   1572
      End
      Begin VB.TextBox txtCryptKey 
         Height          =   285
         Left            =   6120
         TabIndex        =   41
         Top             =   240
         Width           =   1572
      End
      Begin MSComctlLib.ProgressBar prbCrypt 
         Height          =   132
         Left            =   960
         TabIndex        =   40
         Top             =   1020
         Width           =   4452
         _ExtentX        =   7858
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.TextBox txtTargetFile 
         Height          =   285
         Left            =   960
         TabIndex        =   37
         Top             =   600
         Width           =   4455
      End
      Begin VB.CommandButton cmdEncryptFile 
         Height          =   288
         Left            =   5040
         TabIndex        =   36
         Top             =   250
         Width           =   372
      End
      Begin VB.TextBox txtSourceFile 
         Height          =   288
         Left            =   960
         TabIndex        =   35
         Top             =   240
         Width           =   4092
      End
      Begin VB.Label lblCryptKey 
         Alignment       =   1  'Right Justify
         Caption         =   "Key"
         Height          =   255
         Left            =   5520
         TabIndex        =   42
         Top             =   260
         Width           =   495
      End
      Begin VB.Label lblDest 
         Alignment       =   1  'Right Justify
         Caption         =   "Dest."
         Height          =   255
         Left            =   180
         TabIndex        =   39
         Top             =   640
         Width           =   615
      End
      Begin VB.Label lblSource 
         Alignment       =   1  'Right Justify
         Caption         =   "Source"
         Height          =   255
         Left            =   180
         TabIndex        =   38
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.Frame fraCryptstr 
      Caption         =   "Encr;ypt / Decrypt string"
      Height          =   1092
      Left            =   120
      TabIndex        =   20
      Top             =   3840
      Width           =   7812
      Begin VB.CommandButton cmdDecryptstr 
         Caption         =   "Decr;ypt"
         Height          =   288
         Left            =   4680
         TabIndex        =   33
         Top             =   620
         Width           =   1332
      End
      Begin VB.TextBox txtCrypt2 
         Height          =   288
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   4452
      End
      Begin VB.CommandButton cmdEncryptstr 
         Caption         =   "Encrypt"
         Height          =   288
         Left            =   4680
         TabIndex        =   31
         Top             =   260
         Width           =   1332
      End
      Begin VB.TextBox txtKey1 
         Height          =   288
         Left            =   6120
         TabIndex        =   30
         Top             =   600
         Width           =   1572
      End
      Begin VB.TextBox txtCrypt1 
         Height          =   288
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   4452
      End
      Begin VB.Label lblKeystr 
         Alignment       =   2  'Center
         Caption         =   "Key"
         Height          =   252
         Left            =   6240
         TabIndex        =   34
         Top             =   360
         Width           =   1332
      End
   End
   Begin VB.Frame fraWipe 
      Caption         =   "Secure filewipe"
      Height          =   852
      Left            =   120
      TabIndex        =   19
      Top             =   2880
      Width           =   7812
      Begin MSComctlLib.ProgressBar prbWipe1 
         Height          =   132
         Left            =   5160
         TabIndex        =   28
         Top             =   600
         Width           =   2532
         _ExtentX        =   4471
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Min             =   1e-4
         Scrolling       =   1
      End
      Begin VB.OptionButton optWipe2 
         Caption         =   "Governmentwipe"
         Height          =   192
         Left            =   1440
         TabIndex        =   27
         Top             =   600
         Width           =   1932
      End
      Begin VB.OptionButton optWipe1 
         Caption         =   "Normal"
         Height          =   192
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Width           =   1092
      End
      Begin VB.CommandButton cmdWipe 
         Caption         =   "Wipe file"
         Height          =   288
         Left            =   5160
         TabIndex        =   25
         Top             =   240
         Width           =   2532
      End
      Begin VB.CommandButton cmdWipeSelect 
         Height          =   288
         Left            =   4200
         TabIndex        =   24
         Top             =   250
         Width           =   372
      End
      Begin VB.TextBox txtWipe 
         Height          =   288
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   4092
      End
   End
   Begin VB.Frame fraPW 
      Caption         =   "Generate password"
      Height          =   972
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   7812
      Begin VB.ComboBox cboPW 
         Height          =   360
         ItemData        =   "Form1.frx":030A
         Left            =   3675
         List            =   "Form1.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   560
         Width           =   732
      End
      Begin VB.CommandButton cmdPW 
         Caption         =   "Generate password"
         Height          =   288
         Left            =   5160
         TabIndex        =   16
         Top             =   600
         Width           =   2532
      End
      Begin VB.CheckBox chkPW4 
         Caption         =   "signs"
         Height          =   252
         Left            =   2040
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox chkPW3 
         Caption         =   "numbers"
         Height          =   252
         Left            =   2040
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkPW2 
         Caption         =   "uppercase"
         Height          =   252
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chkPW1 
         Caption         =   "lower case"
         Height          =   252
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblChars 
         Alignment       =   2  'Center
         Caption         =   "Nr. of chars"
         Height          =   255
         Left            =   3360
         TabIndex        =   17
         Top             =   260
         Width           =   1335
      End
      Begin VB.Label lblPW 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   275
         Left            =   5160
         TabIndex        =   11
         Top             =   240
         Width           =   2532
      End
   End
   Begin VB.Frame fraCRCfile 
      Caption         =   "Calculate CRC32 on a file"
      Height          =   732
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   7812
      Begin VB.CommandButton cmdCRCFile 
         Caption         =   "CRC32 ==>"
         Height          =   288
         Left            =   4680
         TabIndex        =   9
         Top             =   240
         Width           =   1212
      End
      Begin VB.TextBox txtCRCfile 
         Height          =   288
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4092
      End
      Begin VB.CommandButton cmdSelectFile 
         Height          =   288
         Left            =   4200
         TabIndex        =   6
         Top             =   250
         Width           =   372
      End
      Begin VB.Label lblCRCfile 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   275
         Left            =   6120
         TabIndex        =   7
         Top             =   240
         Width           =   1572
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   6360
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      DialogTitle     =   "Select a file"
   End
   Begin VB.Frame fraCRC_string 
      Caption         =   "Calculate CRC32 on a string"
      Height          =   732
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7812
      Begin VB.TextBox txtCRCstr 
         Height          =   288
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4452
      End
      Begin VB.CommandButton cmdCRCstr 
         Caption         =   "CRC32 ==>"
         Height          =   288
         Left            =   4680
         TabIndex        =   2
         Top             =   210
         Width           =   1212
      End
      Begin VB.Label lblCRCstr 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   275
         Left            =   6120
         TabIndex        =   3
         Top             =   240
         Width           =   1572
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   288
      Left            =   6240
      TabIndex        =   0
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label lblDemo 
      Caption         =   "Demo program using C_Crypt.DLL "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   6480
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' |----------------------------------------------------------------------------------------------------
' Project                : C_Crypt  , demo-project
' Form-name             :  Form1
' Company                : CONFITE
' Programmer             : Joost Rongen
' Date last access       : 10-nov-2000
' Functional discription : CRC32 / Encryption / Decryption / Secure wiping
' -----------------------------------------------------------------------------------------------------

' TO MAKE IT WORK:

' First register C_Crypt.dll to your computer, it you don't have the dll yet, you can download
' it from :      ftp://ftp.confite.nl/pub/

' DOS-prompt>  regsvr32 C_Crypt.dll

' Next goto References in menu Project and set a reference to this ActiveX-component.
' (see reference.jpg)

' -----------------------------------------------------------------------------------------------------

' the program starts here

Option Explicit

' We need to instance the some classes manually, allthought they have the
' globalmultiuse property set, for getting it's events raised to this
' form. New instance is done in Form_load.

Private WithEvents WIPE As C_Crypt.C_WIPE
Attribute WIPE.VB_VarHelpID = -1
Private WithEvents CEF As C_Crypt.C_CEF
Attribute CEF.VB_VarHelpID = -1

Private Sub cboPW_Click()
  Me.lblPW.Caption = ""
End Sub

Private Sub CEF_ProgbarDecryptFile(intProgress As Integer)
  ' update progressbar, event raised by the CEF object
  Me.prbCrypt.Value = intProgress
End Sub

Private Sub CEF_ProgbarEncryptFile(intProgress As Integer)
  ' update progressbar, event raised by the CEF object
  Me.prbCrypt.Value = intProgress
End Sub

Private Sub chkPW1_Click()
  Me.lblPW.Caption = ""
End Sub

Private Sub chkPW2_Click()
  Me.lblPW.Caption = ""
End Sub

Private Sub chkPW3_Click()
  Me.lblPW.Caption = ""
End Sub

Private Sub chkPW4_Click()
   Me.lblPW.Caption = ""
End Sub

Private Sub cmdClose_Click()
  End
End Sub

Private Sub cmdCRCFile_Click()
  ' calculate CRC32 of selected file
  Me.lblCRCfile.Caption = CRC32file(Me.txtCRCfile.Text)
End Sub

Private Sub cmdCRCstr_Click()
  ' calculate CRC32 of input in textbox
  Me.lblCRCstr.Caption = CRC32str(Me.txtCRCstr.Text)
End Sub

Private Sub cmdDecryptstr_Click()
  If Len(Trim(Me.txtCrypt2.Text)) > 0 Then
    If Len(Trim(Me.txtKey1)) > 0 Then
       '  decrypt using the provided key
       Me.txtCrypt1.Text = DecryptString(Me.txtCrypt2.Text, Me.txtKey1.Text)
     Else
     '  no key provided, decryption with built-in default key
       Me.txtCrypt1.Text = DecryptString(Me.txtCrypt2.Text)
    End If
    If Len(Trim(Me.txtCrypt1)) = 0 Then _
       MsgBox "Invalid key", vbOKOnly + vbCritical, "Decryption"
  End If
End Sub

Private Sub cmdEncryptFile_Click()
   ' select file thrue windows dialog
   Me.CommonDialog1.ShowOpen
   Me.txtSourceFile.Text = Me.CommonDialog1.FileName
   Me.prbCrypt.Value = 0.1
End Sub

Private Sub cmdEncryptstr_Click()
  If Len(Trim(Me.txtCrypt1.Text)) > 0 Then
    If Len(Trim(Me.txtKey1)) > 0 Then
        ' encrypt with provided key
        Me.txtCrypt2.Text = EncryptString(Me.txtCrypt1.Text, Me.txtKey1.Text)
      Else
        ' no key provided, encryption with the built-in default key
        Me.txtCrypt2.Text = EncryptString(Me.txtCrypt1.Text)
    End If
  End If
End Sub

Private Sub cmdFileDecrypt_Click()
  If Len(Trim(Me.txtSourceFile.Text)) > 0 Then
    ' if no key is given by the user, the DLL's default one is to be used
     Me.txtTargetFile = _
       CEF.DecryptFile(Me.txtSourceFile.Text, _
                       Me.txtTargetFile.Text, _
                       Me.txtCryptKey.Text)
     Me.prbCrypt.Value = 0.1
  End If
End Sub

Private Sub cmdFileEncrypt_Click()
  If Len(Trim(Me.txtSourceFile.Text)) > 0 Then
    ' if no key is given by the user, the DLL's default one is to be used
    Me.txtTargetFile.Text = _
       CEF.EncryptFile(Me.txtSourceFile.Text, _
                          Me.txtTargetFile.Text, _
                          Me.txtCryptKey.Text, _
                          "This is what it looks like")
    Me.prbCrypt.Value = 0.1
  End If
End Sub

Private Sub cmdPW_Click()
Dim intPWSetting As Integer
' take password-properties from checkboxes
If Me.chkPW1.Value = 1 Then intPWSetting = intPWSetting + 1
If Me.chkPW2.Value = 1 Then intPWSetting = intPWSetting + 2
If Me.chkPW3.Value = 1 Then intPWSetting = intPWSetting + 4
If Me.chkPW4.Value = 1 Then intPWSetting = intPWSetting + 8
' generate password
Me.lblPW = GeneratePassword(Me.cboPW.ListIndex + 1, intPWSetting)
End Sub

Private Sub cmdSelectFile_Click()
   ' select file thrue windows dialog
   Me.CommonDialog1.ShowOpen
   Me.txtCRCfile.Text = Me.CommonDialog1.FileName
   Me.lblCRCfile.Caption = ""
End Sub

Private Sub cmdWipe_Click()
Dim intWipeType As Integer
    ' warn the user about what he's doing
    If MsgBox("WARNING !!!" & vbCrLf & _
        "Keep in mind that it will NEVER be possible to restore a wiped file !" & _
         vbCrLf & "Continue ?", vbYesNo + vbCritical + vbDefaultButton2, "WIPE .. ??") = vbYes Then
       ' used clicked OK to wipe the file definately
       If Me.optWipe1 Then
          If WIPE.FileWipe(Me.txtWipe.Text, Normal) Then
            MsgBox "File has been removed from your disk using the NORMAL wipe methode", vbInformation + vbOKOnly, "Done . ."
          End If
        Else
          If WIPE.FileWipe(Me.txtWipe.Text, Governmentwipe) Then
            MsgBox "File has been removed from your disk using the GOVERNMENT wipe methode", vbInformation + vbOKOnly, "Done . ."
          End If
       End If
       Me.txtWipe.Text = ""
       Me.prbWipe1.Value = 0.1
    End If
End Sub

Private Sub cmdWipeSelect_Click()
   ' select file thrue windows dialog
   Me.CommonDialog1.ShowOpen
   Me.txtWipe.Text = Me.CommonDialog1.FileName
   Me.prbWipe1.Value = 0.1
End Sub

Private Sub Form_Activate()
  Me.cboPW.ListIndex = 7   ' default 8 characters PW
  Me.chkPW1.Value = 1      ' default PW containing lowercase
  Me.chkPW2.Value = 1      '   and uppercase chars
  Me.optWipe1.Value = True  ' nomal wipe
End Sub

Private Sub Form_Load()
  ' create instance of C_WIPE and C_CEF
  Set WIPE = New C_Crypt.C_WIPE
  Set CEF = New C_Crypt.C_CEF
End Sub

Private Sub Form_Terminate()
  Set WIPE = Nothing
  Set CEF = Nothing
End Sub

Private Sub txtCRCfile_KeyPress(KeyAscii As Integer)
  Me.lblCRCfile.Caption = ""    ' empty lable
End Sub

Private Sub txtCRCstr_KeyPress(KeyAscii As Integer)
 Me.lblCRCstr.Caption = ""    ' empty lable
End Sub

Private Sub txtSourceFile_KeyPress(KeyAscii As Integer)
  Me.prbCrypt.Value = 0.1
End Sub

Private Sub WIPE_ProgbarFileWipe(intProgress As Integer)
   ' event raised by the WIPE object to update the progressbar
   Me.prbWipe1.Value = intProgress
End Sub
