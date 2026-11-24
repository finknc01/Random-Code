VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditServer
   Caption         =   "Edit Server | RackScan"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   OleObjectBlob   =   "frmEditServer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEditServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Paste into a UserForm named frmEditServer. Control names are in EIGHT_BUTTON_SETUP.md.
Private mTarget As ListRow
Private mRackCapacity As Long
Private WithEvents cboRack As MSForms.ComboBox
Private WithEvents txtUPosition As MSForms.TextBox
Private WithEvents txtHeightU As MSForms.TextBox
Private txtHostname As MSForms.TextBox
Private txtSerial As MSForms.TextBox
Private txtAssetTag As MSForms.TextBox
Private WithEvents cmdSave As MSForms.CommandButton
Private WithEvents cmdCancel As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Dim rackIds As Collection, item As Variant
    BuildControls
    Me.Caption = "Edit Server | RackScan"
    Me.cboRack.Style = fmStyleDropDownList
    Me.cmdSave.Default = True
    Me.cmdCancel.Cancel = True
    Me.Controls("lblMessage").Caption = ""
    Set rackIds = modServerEntry.RegisteredRackIds()
    For Each item In rackIds
        Me.cboRack.AddItem CStr(item)
    Next item
End Sub

Public Sub LoadServer(ByVal target As ListRow)
    Dim inv As ListObject
    Dim rawHeight As String
    Dim oldRack As String, canonical As String, issue As String
    Set mTarget = target
    Set inv = modRackData.GetTable("Server Inventory", "tblInventory")
    oldRack = modRackData.RowText(target, modRackData.ColumnIndex(inv, "Rack ID"))
    If modServerEntry.RackCapacity(oldRack, canonical, issue) > 0 Then
        Me.cboRack.Value = canonical
    Else
        Me.Controls("lblMessage").Caption = "Original rack is invalid: " & issue & _
                                ". Choose a registered rack before saving."
    End If
    Me.txtUPosition.Value = modRackData.RowText(target, modRackData.ColumnIndex(inv, "U Position"))
    rawHeight = modRackData.RowText(target, modRackData.ColumnIndex(inv, "Height (U)"))
    If Len(rawHeight) = 0 Then rawHeight = "1"
    Me.txtHeightU.Value = rawHeight
    Me.txtHostname.Value = modRackData.RowText(target, modRackData.ColumnIndex(inv, "Hostname"))
    Me.txtSerial.Value = modRackData.RowText(target, modRackData.ColumnIndex(inv, "Serial Number"))
    Me.txtAssetTag.Value = modRackData.RowText(target, modRackData.ColumnIndex(inv, "Asset Tag"))
    Me.Controls("lblTarget").Caption = "Editing inventory row " & target.Range.Row
    UpdatePlacement
End Sub

Private Sub cboRack_Change()
    UpdatePlacement
End Sub

Private Sub txtUPosition_Change()
    UpdatePlacement
End Sub

Private Sub txtHeightU_Change()
    UpdatePlacement
End Sub

Private Sub UpdatePlacement()
    Dim canonical As String, issue As String
    Dim bottomU As Long, heightU As Long
    On Error GoTo HandleError
    mRackCapacity = modServerEntry.RackCapacity( _
        modRackData.TextValue(Me.cboRack.Value), canonical, issue)
    If mRackCapacity = 0 Then
        Me.Controls("lblPlacement").Caption = issue
        Exit Sub
    End If
    bottomU = modRackData.ParseU(modRackData.TextValue(Me.txtUPosition.Value))
    heightU = modRackData.ParseU(modRackData.TextValue(Me.txtHeightU.Value), False)
    If bottomU = 0 Or heightU = 0 Then
        Me.Controls("lblPlacement").Caption = "Enter bottom U and server height (1-42)."
    ElseIf bottomU + heightU - 1 > mRackCapacity Then
        Me.Controls("lblPlacement").Caption = "Out of range: top U exceeds U" & mRackCapacity
    Else
        Me.Controls("lblPlacement").Caption = "Occupies U" & bottomU & "-U" & (bottomU + heightU - 1)
    End If
    Exit Sub
HandleError:
    Me.Controls("lblPlacement").Caption = Err.Description
End Sub

Private Sub cmdSave_Click()
    Dim message As String, rackId As String, bottomU As Long
    On Error GoTo HandleError
    Me.cmdSave.Enabled = False
    If Not modServerActions.TryEditServer( _
        mTarget, modRackData.TextValue(Me.cboRack.Value), _
        modRackData.TextValue(Me.txtUPosition.Value), _
        modRackData.TextValue(Me.txtHeightU.Value), _
        modRackData.TextValue(Me.txtHostname.Value), _
        modRackData.TextValue(Me.txtSerial.Value), _
        modRackData.TextValue(Me.txtAssetTag.Value), message) Then
        Me.Controls("lblMessage").Caption = message
        Me.cmdSave.Enabled = True
        Exit Sub
    End If
    rackId = modRackData.TextValue(Me.cboRack.Value)
    bottomU = modRackData.ParseU(modRackData.TextValue(Me.txtUPosition.Value))
    Unload Me
    On Error GoTo RefreshError
    modServerEntry.ShowAddedServer rackId, bottomU
    MsgBox message, vbInformation, "Server Updated"
    Exit Sub
RefreshError:
    MsgBox message & vbCrLf & "Lookup refresh failed: " & Err.Description, vbExclamation
    Exit Sub
HandleError:
    Me.Controls("lblMessage").Caption = Err.Description
    Me.cmdSave.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


' This imported form uses an empty Microsoft Forms designer and creates its
' controls at initialization. The companion .frx is required for import.
Private Sub BuildControls()
    Me.Width = 430
    Me.Height = 425
    Me.BackColor = RGB(247, 249, 252)

    AddLabel "lblTitle", "EDIT SERVER", 18, 12, 380, 23, True
    AddLabel "lblTarget", "Select a server in Rack Lookup", 18, 41, 380, 22, False

    AddLabel "lblRackLabel", "Rack ID", 18, 75, 126, 18, False
    AddLabel "lblULabel", "Bottom U Position", 18, 106, 126, 18, False
    AddLabel "lblHeightLabel", "Height (U)", 18, 137, 126, 18, False
    AddLabel "lblHostnameLabel", "Hostname", 18, 168, 126, 18, False
    AddLabel "lblSerialLabel", "Serial Number", 18, 199, 126, 18, False
    AddLabel "lblAssetLabel", "Asset Tag", 18, 230, 126, 18, False

    Set cboRack = Me.Controls.Add("Forms.ComboBox.1", "cboRack", True)
    With cboRack
        .Left = 152: .Top = 72: .Width = 246: .Height = 22
        .BackColor = RGB(255, 250, 231)
        .TabIndex = 0
    End With
    Set txtUPosition = AddTextField("txtUPosition", 103, 1)
    Set txtHeightU = AddTextField("txtHeightU", 134, 2)
    Set txtHostname = AddTextField("txtHostname", 165, 3)
    Set txtSerial = AddTextField("txtSerial", 196, 4)
    Set txtAssetTag = AddTextField("txtAssetTag", 227, 5)

    AddLabel "lblPlacement", "Occupied U preview", 18, 267, 380, 28, False
    With Me.Controls("lblPlacement")
        .BackStyle = fmBackStyleOpaque
        .BackColor = RGB(225, 236, 246)
    End With
    AddLabel "lblMessage", "", 18, 302, 380, 42, False
    With Me.Controls("lblMessage")
        .ForeColor = RGB(153, 49, 59)
        .WordWrap = True
    End With

    Set cmdSave = Me.Controls.Add("Forms.CommandButton.1", "cmdSave", True)
    With cmdSave
        .Caption = "Save Changes"
        .Left = 171: .Top = 359: .Width = 132: .Height = 29
        .BackColor = RGB(245, 224, 159)
        .ForeColor = RGB(105, 77, 18)
        .Font.Bold = True
        .TabIndex = 6
    End With
    Set cmdCancel = Me.Controls.Add("Forms.CommandButton.1", "cmdCancel", True)
    With cmdCancel
        .Caption = "Cancel"
        .Left = 312: .Top = 359: .Width = 86: .Height = 29
        .TabIndex = 7
    End With
End Sub

Private Sub AddLabel(ByVal controlName As String, _
                     ByVal captionText As String, _
                     ByVal x As Single, ByVal y As Single, _
                     ByVal w As Single, ByVal h As Single, _
                     ByVal emphatic As Boolean)
    Dim labelControl As MSForms.Label
    Set labelControl = Me.Controls.Add("Forms.Label.1", controlName, True)
    With labelControl
        .Caption = captionText
        .Left = x: .Top = y: .Width = w: .Height = h
        .BackStyle = fmBackStyleTransparent
        .ForeColor = RGB(42, 60, 80)
        .Font.Name = "Segoe UI"
        .Font.Size = IIf(emphatic, 13, 10)
        .Font.Bold = emphatic
    End With
End Sub

Private Function AddTextField(ByVal controlName As String, _
                              ByVal y As Single, _
                              ByVal index As Integer) As MSForms.TextBox
    Dim field As MSForms.TextBox
    Set field = Me.Controls.Add("Forms.TextBox.1", controlName, True)
    With field
        .Left = 152: .Top = y: .Width = 246: .Height = 22
        .BackColor = RGB(255, 250, 231)
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .TabIndex = index
    End With
    Set AddTextField = field
End Function

