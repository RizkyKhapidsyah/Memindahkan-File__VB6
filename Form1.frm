VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memindahkan File"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Pindahkan Ke Folder1"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pindahkan Ke Folder2"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Contoh ini memindahkan file 'c:\MyFile.Zip' ke 'direktori 'c:\MyDir'.
 Private Sub Command1_Click()
  A = MoveFile(App.Path + "\Folder1\Pindahkan-File-Ini.txt", App.Path + "\Folder2\Pindahkan-File-Ini.txt")
  If A Then
     MsgBox "File berhasil dipindahkan!", _
            vbInformation, "Sukses Pindah File"
  Else
     MsgBox "Error. File belum dipindahkan!" & Chr(13) & "Kemungkinan file asal tidak ada" & _
            Chr(13) & "atau file sudah ada di dalam " & _
            Chr(13) & _
            "direktori tujuan!", vbCritical, "Gagal Pindah File"
  End If
End Sub


Private Sub Command2_Click()
  A = MoveFile(App.Path + "\Folder2\Pindahkan-File-Ini.txt", App.Path + "\Folder1\Pindahkan-File-Ini.txt")
  If A Then
     MsgBox "File berhasil dipindahkan!", _
            vbInformation, "Sukses Pindah File"
  Else
     MsgBox "Error. File belum dipindahkan!" & Chr(13) & "Kemungkinan file asal tidak ada" & _
            Chr(13) & "atau file sudah ada di dalam " & _
            Chr(13) & _
            "direktori tujuan!", vbCritical, "Gagal Pindah File"
  End If
End Sub
