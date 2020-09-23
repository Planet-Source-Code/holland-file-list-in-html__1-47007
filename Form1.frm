VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File list in HTML"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Down"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4800
      Width           =   4455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options:"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   4695
      Begin VB.CheckBox Check1 
         Caption         =   "Make the every item in the list have a dot before it"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   4455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Make every item a hyperlink to the file"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Value           =   1  'Checked
         Width           =   4455
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rem."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Drag && Drop files from your file manager here:"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.CheckBox Check3 
         Caption         =   "On top"
         Height          =   495
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1680
         Value           =   1  'Checked
         Width           =   495
      End
      Begin VB.ListBox List1 
         Height          =   3540
         IntegralHeight  =   0   'False
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check3_Click()
If Check3.Value = 1 Then
Call AlwaysOnTop("Enabled", Form1)
Else
Call AlwaysOnTop("Disabled", Form1)
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next 'This basicly prevents errors from showing up
List1.RemoveItem List1.ListIndex 'This will remove the item from the list
End Sub

Private Sub Command2_Click()
If MsgBox("Are you sure?", vbYesNo + vbQuestion, "Clear") = vbYes Then 'Will ask if the user is sure
List1.Clear 'This will clear the list
End If
End Sub

Private Sub Command3_Click()
If List1.ListCount = 0 Then Exit Sub
List1.Visible = False
Open App.Path & "\Filelist.html" For Output As #1 'Create the Filelsit.html file and open it for writing in it
Print #1, "<! You can enter some stuff here>"
For i = 0 To List1.ListCount - 1
List1.ListIndex = i 'Selects a item in the list
If Check2.Value = 1 Then 'If the second checkbox is checked then...
'This is all HTML stuff
ddd = """"
ggg = "<a href=" & ddd & List1.Text & ddd & ">" & List1.Text & "</a><BR>"
Else
ggg = List1.Text & "<BR>"
End If
If Check1.Value = 1 Then ggg = "<li>" & ggg & "</li>"
'Here the HTML stuff ends
Print #1, ggg 'Print ggg in the file
DoEvents
Next i
Print #1, "<! You can enter some stuff here>"
Close #1 'Close the file
MsgBox "Done!", , "Save"
List1.ListIndex = 0
List1.Visible = True
End Sub

Private Sub Command4_Click()
Call MoveUp(List1)
End Sub

Private Sub Command5_Click()
Call MoveDown(List1)
End Sub

Private Sub Form_Load()
Call AlwaysOnTop("Enabled", Form1)
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
List1.Visible = False 'Makes the list invisible, this will increase speed when dropping a big ammount of files
For i = 1 To Data.Files.Count 'This thing loops until it had all the files that are being dropped
ggg = GetFileName(Data.Files.Item(i)) 'This gets only the filename of thie file, this requires the function in the module
If Len(ggg) < 5 Then 'Length of the filename should be at least 5 chars.
Else
If Mid$(ggg, Len(ggg) - 3, 1) = "." Then 'Checks if this actually has an extension
List1.AddItem ggg 'This adds 'ggg' to the list
End If
End If
DoEvents 'This is only needed when many files are being dropped, this code gives your CPU a little break every time
Next i 'Here it restarts again
List1.Visible = True 'Makes it visible again
End Sub
