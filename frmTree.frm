VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTree 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TreeView for Real VB Beginners"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgTree 
      Left            =   750
      Top             =   5805
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTree.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTree.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTree.frx":0654
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTree.frx":0766
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTree.frx":0BB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7800
      TabIndex        =   5
      Top             =   6000
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      Height          =   5565
      Left            =   4125
      TabIndex        =   4
      Top             =   345
      Width           =   4800
      Begin VB.TextBox txtUnitPrice 
         Height          =   315
         Left            =   1615
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1890
         Width           =   1500
      End
      Begin VB.TextBox txtUM 
         Height          =   315
         Left            =   1615
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1485
         Width           =   1500
      End
      Begin VB.TextBox txtItemDesc 
         Height          =   315
         Left            =   1615
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1080
         Width           =   3000
      End
      Begin VB.TextBox txtType 
         Height          =   315
         Left            =   1615
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   690
         Width           =   3000
      End
      Begin VB.TextBox txtItemNum 
         Height          =   315
         Left            =   1615
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label Label16 
         Caption         =   $"frmTree.frx":100A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   345
         TabIndex        =   30
         Top             =   4695
         Width           =   4275
      End
      Begin VB.Label Label15 
         Caption         =   "* Both treeviews share the same imagelist (named imgTree in this project)."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   330
         TabIndex        =   29
         Top             =   4185
         Width           =   4035
      End
      Begin VB.Label Label14 
         Caption         =   "* Sub DisplayItem() is called when a child node for each treeview is clicked or the result listbox in Quick Search is clicked"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   360
         TabIndex        =   28
         Top             =   3480
         Width           =   4035
      End
      Begin VB.Label Label13 
         Caption         =   $"frmTree.frx":10D4
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   390
         TabIndex        =   27
         Top             =   2790
         Width           =   4035
      End
      Begin VB.Label Label12 
         Caption         =   "Few notes about the program  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   195
         TabIndex        =   26
         Top             =   2460
         Width           =   2430
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         BorderWidth     =   2
         X1              =   0
         X2              =   4755
         Y1              =   2370
         Y2              =   2370
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Unit Price :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   555
         TabIndex        =   15
         Top             =   1860
         Width           =   810
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Unit of Measure :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   105
         TabIndex        =   13
         Top             =   1455
         Width           =   1275
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Item Description :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   105
         TabIndex        =   11
         Top             =   1050
         Width           =   1275
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Type : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   915
         TabIndex        =   9
         Top             =   690
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Item Number :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   330
         TabIndex        =   7
         Top             =   300
         Width           =   1050
      End
   End
   Begin TabDlg.SSTab sstabTreeView 
      Height          =   5490
      Left            =   90
      TabIndex        =   0
      Top             =   435
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   9684
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "By Letter"
      TabPicture(0)   =   "frmTree.frx":116B
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tvwLetters"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "By Type"
      TabPicture(1)   =   "frmTree.frx":1187
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tvwType"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "QSearch"
      TabPicture(2)   =   "frmTree.frx":11A3
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstResult"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "chkAllType"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cboType"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdGo"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtKey"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lblResult"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label9"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label8"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      Begin VB.ListBox lstResult 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3570
         Left            =   -74940
         TabIndex        =   23
         Top             =   1875
         Width           =   3855
      End
      Begin VB.CheckBox chkAllType 
         Alignment       =   1  'Right Justify
         Caption         =   "&All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71670
         TabIndex        =   22
         Top             =   1275
         Value           =   1  'Checked
         Width           =   495
      End
      Begin VB.ComboBox cboType 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74850
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1245
         Width           =   3030
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "&Go"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -71580
         TabIndex        =   19
         Top             =   645
         Width           =   420
      End
      Begin VB.TextBox txtKey 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74865
         TabIndex        =   18
         Top             =   675
         Width           =   3015
      End
      Begin MSComctlLib.TreeView tvwType 
         Height          =   5130
         Left            =   -75000
         TabIndex        =   6
         Top             =   300
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   9049
         _Version        =   393217
         Style           =   7
         ImageList       =   "imgTree"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwLetters 
         Height          =   5145
         Left            =   0
         TabIndex        =   1
         Top             =   300
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   9075
         _Version        =   393217
         Style           =   7
         HotTracking     =   -1  'True
         ImageList       =   "imgTree"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblResult 
         Caption         =   "Result : "
         Height          =   240
         Left            =   -74850
         TabIndex        =   25
         Top             =   1680
         Width           =   3660
      End
      Begin VB.Label Label9 
         Caption         =   "Select Type :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74850
         TabIndex        =   20
         Top             =   1035
         Width           =   1005
      End
      Begin VB.Label Label8 
         Caption         =   "Quick Search Key :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74865
         TabIndex        =   17
         Top             =   465
         Width           =   1635
      End
   End
   Begin VB.Label Label10 
      Caption         =   "Hit <ESC> key to end program without prompt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   165
      TabIndex        =   24
      Top             =   6195
      Width           =   4020
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4125
      TabIndex        =   3
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TreeView"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   105
      TabIndex        =   2
      Top             =   120
      Width           =   3990
   End
End
Attribute VB_Name = "frmTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Please read the Readme.txt
' I should have written all of it here, but I dont want the code to  appear too "text-y", so
' I made a text file where I can explain some of the brief introductions and code starters.
'
' by: delTex 2002
'
' Start of Code here:
' -------------------
'
Option Explicit
Private WithEvents conTree As ADODB.Connection      ' for our database connection
Attribute conTree.VB_VarHelpID = -1
Private WithEvents rsTreeItems As ADODB.Recordset   ' for tblItem recordset
Attribute rsTreeItems.VB_VarHelpID = -1
Private WithEvents rsType As ADODB.Recordset        ' for tblType recordset
Attribute rsType.VB_VarHelpID = -1
Public strSQL As String                             ' SQL statement holder
Public blnAllType As Boolean                        ' indicator variable that tells whether to display all types (true) or specific type (false)
Public strNodeLabel As String                       ' Node description, to be used
                                                    ' to restore the original label
                                                    ' after the node has been edited (you'll find that out later,
                                                    ' or when you run the program and trick the treeviews).
                                                    ' It is because Treeviews, by default, have editable labels
                                                    ' If you dont want to use APIs to make treeview uneditable,
                                                    ' You are free to use this technique

Private Sub chkAllType_Click()
    If cboType.Enabled = False Then
        cboType.Enabled = True
        cboType.SetFocus
        blnAllType = False
    Else
        cboType.Enabled = False
        cmdGo.SetFocus
        blnAllType = True
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGo_Click()
    Call GoSearch
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then       ' User presses ESC key to exit program without prompt
                                ' Make sure you set the form's KEYPREVIEW property to TRUE
                                ' to capture KeyPress event
                                '
                                ' to learn more about using ESC to Exit application,
                                ' please visit my tutorial regarding this:
                                '    http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=35559&lngWId=1
        End
    End If
End Sub

Private Sub Form_Load()
    Call connectDB              ' Connect to database
    Call initializeLetterTree   ' initialize the BY LETTER treeview
    Call initializeTypeTree     ' initialize the BY TYPE TreeView
    blnAllType = True
    
    'Center the form
    Top = (Screen.Height - ScaleHeight) / 2
    Left = (Screen.Width - ScaleWidth) / 2
    
    MsgBox "Thanks for downloading this tutorial." & vbCrLf & "Please read the Readme.txt", vbInformation, "Read"
End Sub

Private Sub initializeLetterTree()
    Dim strLetters As String, strKey As String * 1
    Dim nodLetter As Node
    Dim i As Integer
    
    ' This technique is from SAMS 21 Days book
    tvwLetters.LineStyle = tvwRootLines     ' Simply change the style of line of the treeview
    tvwLetters.Nodes.Clear                  ' always clear
    
    strLetters = "ABCDEFGHIJKLMNOPQRSTUVXYZ"    ' initialize key for the letter tree
    
    For i = 1 To 26
        strKey = Mid(strLetters, i, 1)       ' chop a letter
        Set nodLetter = tvwLetters.Nodes.Add(, , strKey, strKey, 1) ' Notice that we have the same key with the text
                                                                    ' because strKey is unique and at the same time,
                                                                    ' we use it as a description to group the table by letters
        Set rsTreeItems = New ADODB.Recordset
        
        ' Get all item(s) that start with current letter
        strSQL = "Select * from tblItems where itemdesc like '" & strKey & "%' order by itemdesc"
        rsTreeItems.Open strSQL, conTree, adOpenStatic, adLockOptimistic
        
        If rsTreeItems.RecordCount > 0 Then ' If we have item(s) that start with current letter, insert it as a child to that letter
            While Not rsTreeItems.EOF
                Set nodLetter = tvwLetters.Nodes.Add(strKey, tvwChild, strKey & Str(rsTreeItems.Fields("itemno")), rsTreeItems.Fields("itemdesc"), 3)
                rsTreeItems.MoveNext
            Wend
        End If
    Next i
    Set rsTreeItems = Nothing   ' Release rsTreeItems from memory
End Sub

Private Sub connectDB()     ' Self-explanatory. This simply connects to database
                            ' But you must remember that will only connect ONCE!
                            
    Dim strPath As String, strCon As String, strDBase As String
    strDBase = "mdbTree.mdb"
    strPath = App.Path & "\"
    strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & strPath & strDBase
    
    Set conTree = New ADODB.Connection
    conTree.Open strCon
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Are you sure you want to exit?", vbYesNo + vbCritical + vbDefaultButton2, Me.Caption) = vbYes Then
        Set rsTreeItems = Nothing
        Set rsType = Nothing
        End
    Else
        Cancel = 1
    End If
End Sub

Private Sub lstResult_Click()
    Call DisplayItem(lstResult.Text)
End Sub

Private Sub tvwLetters_AfterLabelEdit(Cancel As Integer, NewString As String)
    NewString = strNodeLabel        ' bring back to its orignal label
End Sub

Private Sub tvwLetters_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Image = 1      ' When a letter-node is closed, display the CLOSED FOLDER image
                        ' Refer to table in Readme.txt file
End Sub

Private Sub tvwLetters_Expand(ByVal Node As MSComctlLib.Node)
    Node.Image = 2      ' When the letter-node is expanded, display the OPEN FOLDER image
                        ' Refer to table in Readme.txt file
End Sub

Private Sub initializeTypeTree()
    ' This is much harder to explain. Just trust your instinct here
    ' ... anyway, let's continue ...
    
    ' declare variables
    Dim nodType As Node
    Dim intTypeKey As Integer
    Dim rsTypeChild As ADODB.Recordset
    Dim myTypeKey As String
    
    intTypeKey = 1      ' we need this, since we need a unique node key.
                        ' this variale will increment by one.
                        ' Just read on to understand more.
    
    strSQL = "Select * from tblType order by typedesc"
        
    Set rsType = New ADODB.Recordset
    rsType.Open strSQL, conTree, adOpenStatic, adLockOptimistic
    
    tvwType.LineStyle = tvwRootLines
    tvwType.Nodes.Clear
    
    ' if you have a better approach on this please feel free to contact me
    
    If rsType.RecordCount > 0 Then      ' check if rsType is not empty
        While Not rsType.EOF            ' go around each record and add it to the treeview
            myTypeKey = Str(intTypeKey) & Str(rsType.Fields("typeno"))  ' create a unique node key
                                                                        ' notice that we converted the intTypeKey to String
                                                                        ' since node key doesnt accept integer, AND I DUNNO WHY.
            Set nodType = tvwType.Nodes.Add(, , myTypeKey, rsType.Fields("typedesc"), 4)
                        
            ' Now let's check from tblItems if we have record(s) under the current type
            strSQL = "Select * from tblITems where typeno = " & rsType.Fields("typeno") & " order by itemdesc"
                
            Set rsTypeChild = New ADODB.Recordset
            rsTypeChild.Open strSQL, conTree, adOpenStatic, adLockOptimistic
            
            If rsTypeChild.RecordCount > 0 Then     ' if it has record(s) under the current type, insert it as a child node to the current node.
                                                    ' please notice how the changes it syntax appear.
                While Not rsTypeChild.EOF
                    Set nodType = tvwType.Nodes.Add(myTypeKey, tvwChild, myTypeKey & Str(rsTypeChild.Fields("itemno")), rsTypeChild.Fields("itemdesc"), 3)
                    rsTypeChild.MoveNext
                Wend
            End If
            
            cboType.AddItem rsType.Fields("typedesc")
            rsType.MoveNext
            intTypeKey = intTypeKey + 1             ' increment this variable to make it unique
        Wend
    End If
    Set rsType = Nothing        ' Release recordset from memory
End Sub

Private Sub tvwLetters_NodeClick(ByVal Node As MSComctlLib.Node)
    strNodeLabel = Node
    If Node.Image = 3 Then  ' If node is not a parent noe, as represented by
                            ' TEXT image (image no. 3 of imgTree )
        Call DisplayItem(Node)
    End If
End Sub

Private Sub tvwType_AfterLabelEdit(Cancel As Integer, NewString As String)
    NewString = strNodeLabel    ' Bring back to its original label
End Sub

Private Sub tvwType_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Image = 4      ' Just like the other treeview, when this type-node is collapsed,
                        ' display the CLOSED CATALOG image. Refer to table in Readme.txt file
End Sub

Private Sub tvwType_Expand(ByVal Node As MSComctlLib.Node)
    Node.Image = 5      ' Just like the other treeview, when this type-node is collapsed,
                        ' display the OPEN CATALOG image. Refer to table in Readme.txt file
End Sub

Private Sub DisplayItem(ByVal strClickedItem As String)
    ' this sub will be called when a child node is clicked
    
    Dim rsDisplay As ADODB.Recordset
    Set rsDisplay = New ADODB.Recordset
    strSQL = "Select * from tblItems where trim(itemdesc) = '" & Trim(strClickedItem) & "' order by itemdesc"
    rsDisplay.Open strSQL, conTree, adOpenStatic, adLockOptimistic
    
    ' Use Immediate If Function (IIF), to avoid "INVALID USE OF NULL" error
    txtItemNum.Text = IIf(IsNull(rsDisplay.Fields("itemno")) = True, "", rsDisplay.Fields("itemno"))
    txtItemDesc.Text = IIf(IsNull(rsDisplay.Fields("itemdesc")) = True, "", rsDisplay.Fields("itemdesc"))
    txtUM.Text = IIf(IsNull(rsDisplay.Fields("um")) = True, "None entered", rsDisplay.Fields("um"))
    txtUnitPrice.Text = IIf(IsNull(rsDisplay.Fields("unitprice")) = True, 0#, rsDisplay.Fields("unitprice"))
    txtType.Text = IIf(IsNull(getTypeDesc(rsDisplay.Fields("typeno"), conTree)) = True, "None entered", getTypeDesc(rsDisplay.Fields("typeno"), conTree))
    ' Release rsDisplay from memory
    Set rsDisplay = Nothing
End Sub

Private Sub tvwType_NodeClick(ByVal Node As MSComctlLib.Node)
    strNodeLabel = Node
    If Node.Image = 3 Then  ' If node is not a parent node, as represented by
                            ' TEXT image (image no. 3 of imgTree )
        Call DisplayItem(Node)
    End If
End Sub

Private Sub GoSearch()
    Dim rsSearchItem As ADODB.Recordset
    
    lstResult.Clear     ' to clear the previous result
      
    If blnAllType = True Then   ' if all Type checkbox is checked
        strSQL = "Select * from tblItems where ucase(trim(itemdesc)) like '" & Trim(UCase(txtKey.Text)) & "%' order by itemdesc"
    Else                        ' otherwise
        strSQL = "Select * from tblItems where ucase(trim(itemdesc)) like '" & Trim(UCase(txtKey.Text)) & "%' and typeno = " & getTypeNo(UCase(Trim(cboType.Text)), conTree) & " order by itemdesc"
    End If
    
    'explanation above also applies here.
    Set rsSearchItem = New ADODB.Recordset
    rsSearchItem.Open strSQL, conTree, adOpenStatic, adLockOptimistic
    If rsSearchItem.RecordCount > 0 Then
        While Not rsSearchItem.EOF
            lstResult.AddItem rsSearchItem.Fields("itemdesc")
            rsSearchItem.MoveNext
        Wend
    Else
        MsgBox "No records found for the search key you specified.", vbInformation, "No find"
    End If
    
    lblResult.Caption = "Result : " & "(" & rsSearchItem.RecordCount & " items)"
    Set rsSearchItem = Nothing
    
End Sub

Private Sub txtKey_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdGo_Click
    End If
End Sub
