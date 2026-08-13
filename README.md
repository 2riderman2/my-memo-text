Option Explicit

'============================================================
' 設定
'============================================================

' 受信トレイ直下の一時保存フォルダ名
Private Const TEMP_FOLDER_NAME As String = "一時保存"

' ドキュメントフォルダ内の保存先フォルダ名
Private Const SAVE_FOLDER_NAME As String = "受信メール"

' 保存後に付ける分類項目
Private Const CATEGORY_NAME As String = "保存済み"

' ファイル名の最大長
Private Const MAX_FILE_NAME_LENGTH As Long = 180


'============================================================
' イベント監視用
'============================================================

Private WithEvents TempItems As Outlook.Items


'============================================================
' Outlook起動時
'============================================================

Private Sub Application_Startup()

    InitializeMailSaver

End Sub


'============================================================
' 初期化
'============================================================

Private Sub InitializeMailSaver()

    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim tempFolder As Outlook.Folder

    On Error GoTo ErrHandler

    Set ns = Application.Session
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    '----------------------------------------
    ' 一時保存フォルダを取得
    '----------------------------------------
    Set tempFolder = GetSubFolder(inbox, TEMP_FOLDER_NAME)

    If tempFolder Is Nothing Then

        MsgBox _
            "受信トレイの直下に「" & TEMP_FOLDER_NAME & "」フォルダが見つかりません。" & vbCrLf & _
            "フォルダ名を確認してください。", _
            vbExclamation, _
            "メール自動保存"

        Exit Sub

    End If

    '----------------------------------------
    ' 一時保存フォルダを監視
    '----------------------------------------
    Set TempItems = tempFolder.Items

    '----------------------------------------
    ' 分類項目「保存済み」を準備
    '----------------------------------------
    EnsureCategoryExists CATEGORY_NAME

    '----------------------------------------
    ' Outlook停止中に入ったメールを処理
    '----------------------------------------
    ProcessExistingItems tempFolder

    Exit Sub


ErrHandler:

    MsgBox _
        "メール自動保存の初期化中にエラーが発生しました。" & vbCrLf & vbCrLf & _
        "エラー番号：" & Err.Number & vbCrLf & _
        "内容：" & Err.Description, _
        vbExclamation, _
        "メール自動保存"

End Sub


'============================================================
' 一時保存フォルダにメールが追加されたとき
'============================================================

Private Sub TempItems_ItemAdd(ByVal Item As Object)

    On Error GoTo ErrHandler

    ' メール以外は処理しない
    If Not TypeOf Item Is Outlook.MailItem Then
        Exit Sub
    End If

    ProcessMail Item

    Exit Sub


ErrHandler:

    ' エラー時はメールを一時保存フォルダに残す
    Debug.Print _
        "TempItems_ItemAdd Error " & _
        Err.Number & " : " & _
        Err.Description

End Sub


'============================================================
' Outlook起動時に一時保存フォルダ内を処理
'============================================================

Private Sub ProcessExistingItems(ByVal tempFolder As Outlook.Folder)

    Dim i As Long
    Dim obj As Object
    Dim mail As Outlook.MailItem

    On Error GoTo ErrHandler

    ' MoveするとItemsの数が変わるため後ろから処理
    For i = tempFolder.Items.Count To 1 Step -1

        Set obj = tempFolder.Items.Item(i)

        If TypeOf obj Is Outlook.MailItem Then

            Set mail = obj

            ProcessMail mail

        End If

        Set mail = Nothing
        Set obj = Nothing

    Next i

    Exit Sub


ErrHandler:

    Debug.Print _
        "ProcessExistingItems Error " & _
        Err.Number & " : " & _
        Err.Description

End Sub


'============================================================
' メール1件を保存するメイン処理
'============================================================

Private Sub ProcessMail(ByVal mail As Outlook.MailItem)

    Dim saveFolder As String
    Dim fileName As String
    Dim fullPath As String

    Dim inbox As Outlook.Folder

    On Error GoTo ErrHandler

    '----------------------------------------
    ' 保存先フォルダ取得
    '----------------------------------------
    saveFolder = GetSaveFolderPath()

    ' 保存先フォルダが無ければ作成
    EnsureFolderExists saveFolder

    '----------------------------------------
    ' 保存ファイル名作成
    '----------------------------------------
    fileName = BuildFileName(mail)

    fullPath = saveFolder & "\" & fileName

    '----------------------------------------
    ' 同名ファイルが存在するか確認
    '----------------------------------------
    If Dir(fullPath) = "" Then

        ' Outlookメッセージ形式で保存
        mail.SaveAs fullPath, olMSG

    Else

        ' 同名がある場合は保存しない
        Debug.Print "重複のため保存をスキップ：" & fullPath

    End If

    '----------------------------------------
    ' 分類項目「保存済み」を付与
    '----------------------------------------
    AddCategory mail, CATEGORY_NAME

    ' 変更内容保存
    mail.Save

    '----------------------------------------
    ' 受信トレイへ戻す
    '----------------------------------------
    Set inbox = Application.Session.GetDefaultFolder(olFolderInbox)

    mail.Move inbox

    Exit Sub


ErrHandler:

    ' 保存に失敗した場合はMoveしないので
    ' 一時保存フォルダにメールが残る
    Debug.Print _
        "ProcessMail Error " & _
        Err.Number & " : " & _
        Err.Description

End Sub


'============================================================
' ファイル名作成
'============================================================

Private Function BuildFileName(ByVal mail As Outlook.MailItem) As String

    Dim dateText As String
    Dim senderText As String
    Dim subjectText As String
    Dim result As String

    On Error GoTo ErrHandler

    '----------------------------------------
    ' 受信日時
    '----------------------------------------
    dateText = Format(mail.ReceivedTime, "yyyymmdd_HHmmss")

    '----------------------------------------
    ' 差出人
    '----------------------------------------
    senderText = mail.SenderName

    If Trim(senderText) = "" Then
        senderText = "差出人不明"
    End If

    senderText = CleanFileName(senderText)

    '----------------------------------------
    ' 件名
    '----------------------------------------
    subjectText = mail.Subject

    If Trim(subjectText) = "" Then
        subjectText = "件名なし"
    End If

    subjectText = CleanFileName(subjectText)

    '----------------------------------------
    ' ファイル名
    '----------------------------------------
    result = _
        dateText & "_" & _
        senderText & "_" & _
        subjectText

    ' 長すぎる場合は短くする
    If Len(result) > MAX_FILE_NAME_LENGTH Then

        result = Left(result, MAX_FILE_NAME_LENGTH)

    End If

    BuildFileName = result & ".msg"

    Exit Function


ErrHandler:

    BuildFileName = _
        Format(Now, "yyyymmdd_HHmmss") & _
        "_メール.msg"

End Function


'============================================================
' ファイル名に使用できない文字を除去
'============================================================

Private Function CleanFileName(ByVal text As String) As String

    Dim invalidChars As Variant
    Dim ch As Variant

    invalidChars = Array( _
        "\", _
        "/", _
        ":", _
        "*", _
        "?", _
        """", _
        "<", _
        ">", _
        "|" _
    )

    For Each ch In invalidChars

        text = Replace(text, CStr(ch), "_")

    Next ch

    ' 改行・タブ除去
    text = Replace(text, vbCr, " ")
    text = Replace(text, vbLf, " ")
    text = Replace(text, vbTab, " ")

    ' 前後の空白除去
    text = Trim(text)

    ' Windowsでは末尾のピリオドや空白が問題になるため除去
    Do While Len(text) > 0

        If Right(text, 1) = "." Or _
           Right(text, 1) = " " Then

            text = Left(text, Len(text) - 1)

        Else

            Exit Do

        End If

    Loop

    CleanFileName = text

End Function


'============================================================
' Windowsの「ドキュメント」フォルダを取得
'============================================================

Private Function GetSaveFolderPath() As String

    Dim shell As Object
    Dim documentsPath As String

    Set shell = CreateObject("WScript.Shell")

    documentsPath = _
        shell.SpecialFolders("MyDocuments")

    GetSaveFolderPath = _
        documentsPath & "\" & SAVE_FOLDER_NAME

    Set shell = Nothing

End Function


'============================================================
' 保存先フォルダを作成
'============================================================

Private Sub EnsureFolderExists(ByVal folderPath As String)

    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then

        fso.CreateFolder folderPath

    End If

    Set fso = Nothing

End Sub


'============================================================
' 指定したサブフォルダを取得
'============================================================

Private Function GetSubFolder( _
    ByVal parentFolder As Outlook.Folder, _
    ByVal folderName As String _
) As Outlook.Folder

    Dim folder As Outlook.Folder

    On Error Resume Next

    Set folder = parentFolder.Folders(folderName)

    On Error GoTo 0

    Set GetSubFolder = folder

End Function


'============================================================
' 分類項目が無ければ作成
'============================================================

Private Sub EnsureCategoryExists(ByVal categoryName As String)

    Dim category As Outlook.Category

    On Error Resume Next

    Set category = _
        Application.Session.Categories.Item(categoryName)

    On Error GoTo 0

    If category Is Nothing Then

        On Error Resume Next

        Application.Session.Categories.Add _
            categoryName, _
            olCategoryColorNone

        On Error GoTo 0

    End If

    Set category = Nothing

End Sub


'============================================================
' メールに分類項目を追加
'============================================================

Private Sub AddCategory( _
    ByVal mail As Outlook.MailItem, _
    ByVal categoryName As String _
)

    Dim existingCategories As String

    On Error GoTo ErrHandler

    existingCategories = Trim(mail.Categories)

    '----------------------------------------
    ' すでに分類されていれば何もしない
    '----------------------------------------
    If HasCategory(existingCategories, categoryName) Then
        Exit Sub
    End If

    '----------------------------------------
    ' 既存分類がない場合
    '----------------------------------------
    If existingCategories = "" Then

        mail.Categories = categoryName

    Else

        ' 既存分類を残したまま追加
        mail.Categories = _
            existingCategories & "," & categoryName

    End If

    Exit Sub


ErrHandler:

    Debug.Print _
        "AddCategory Error " & _
        Err.Number & " : " & _
        Err.Description

End Sub


'============================================================
' 指定した分類項目が付いているか確認
'============================================================

Private Function HasCategory( _
    ByVal categoriesText As String, _
    ByVal categoryName As String _
) As Boolean

    Dim categories As Variant
    Dim category As Variant

    If Trim(categoriesText) = "" Then

        HasCategory = False
        Exit Function

    End If

    categories = Split(categoriesText, ",")

    For Each category In categories

        If StrComp( _
            Trim(CStr(category)), _
            categoryName, _
            vbTextCompare _
        ) = 0 Then

            HasCategory = True
            Exit Function

        End If

    Next category

    HasCategory = False

End Function
