Option Explicit

'============================================================
' 共有メールボックスの1フォルダを監視するクラス
'============================================================

Private WithEvents mItems As Outlook.Items
Private WithEvents mFolders As Outlook.Folders

Private mTargetFolder As Outlook.Folder

Private mMode As Long

Private mOwner As Object


'============================================================
' 初期化
'============================================================

Public Sub Initialize( _
    ByVal targetFolder As Outlook.Folder, _
    ByVal mode As Long, _
    ByVal owner As Object _
)

    Set mTargetFolder = targetFolder

    Set mItems = _
        targetFolder.Items

    Set mFolders = _
        targetFolder.Folders

    mMode = mode

    Set mOwner = owner

End Sub


'============================================================
' このフォルダにアイテムが追加された
'============================================================

Private Sub mItems_ItemAdd( _
    ByVal Item As Object _
)

    On Error GoTo ErrHandler

    If mOwner Is Nothing Then
        Exit Sub
    End If

    mOwner.HandleSharedFolderItem _
        Item, _
        mMode

    Exit Sub


ErrHandler:

    Debug.Print _
        "CSharedFolderWatcher ItemAdd Error " & _
        Err.Number & " : " & _
        Err.Description

End Sub


'============================================================
' このフォルダの下に新しいサブフォルダが追加された
'============================================================

Private Sub mFolders_FolderAdd( _
    ByVal Folder As Outlook.Folder _
)

    On Error GoTo ErrHandler

    If mOwner Is Nothing Then
        Exit Sub
    End If

    ' 新しいフォルダ自身＋その下位フォルダを監視
    ' すでにメールが入っていれば同時に保存
    mOwner.RegisterSharedFolderTree _
        Folder, _
        mMode, _
        True

    Exit Sub


ErrHandler:

    Debug.Print _
        "CSharedFolderWatcher FolderAdd Error " & _
        Err.Number & " : " & _
        Err.Description

End Sub
