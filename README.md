Option Explicit

'============================================================
' 設定
'============================================================

' 受信トレイ直下の一時保存フォルダ名
Private Const TEMP_FOLDER_NAME As String = "一時保存"

' ドキュメントフォルダ内の受信メール保存先
Private Const RECEIVE_SAVE_FOLDER_NAME As String = "受信メール"

' ドキュメントフォルダ内の送信メール保存先
Private Const SENT_SAVE_FOLDER_NAME As String = "送信メール"

' 受信メール保存後に付ける分類項目
Private Const CATEGORY_NAME As String = "保存済み"

' ファイル名本体の最大文字数
Private Const MAX_FILE_NAME_LENGTH As Long = 180


'============================================================
' 受信メール監視用
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
    Set tempFolder = GetSubFolder( _
        inbox, _
        TEMP_FOLDER_NAME _
    )

    If tempFolder Is Nothing Then

        MsgBox _
            "受信トレイの直下に「" & _
            TEMP_FOLDER_NAME & _
            "」フォルダが見つかりません。" & vbCrLf & _
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
    ' Outlook停止中に一時保存へ入った
    ' メールを処理
    '----------------------------------------
    ProcessExistingItems tempFolder

    Exit Sub


ErrHandler:

    MsgBox _
        "メール自動保存の初期化中にエラーが発生しました。" & _
        vbCrLf & vbCrLf & _
        "エラー番号：" & Err.Number & vbCrLf & _
        "内容：" & Err.Description, _
        vbExclamation, _
        "メール自動保存"

End Sub


'============================================================
' 受信：
' 一時保存フォルダにメールが追加されたとき
'============================================================

Private Sub TempItems_ItemAdd(ByVal Item As Object)

    On Error GoTo ErrHandler

    ' メール以外は処理しない
    If Not TypeOf Item Is Outlook.MailItem Then
        Exit Sub
    End If

    ProcessReceivedMail Item

    Exit Sub


ErrHandler:

    ' エラー時はメールを一時保存フォルダに残す
    Debug.Print _
        "TempItems_ItemAdd Error " & _
        Err.Number & " : " & _
        Err.Description

End Sub


'============================================================
' 受信：
' Outlook起動時に一時保存フォルダ内を処理
'============================================================

Private Sub ProcessExistingItems( _
    ByVal tempFolder As Outlook.Folder _
)

    Dim i As Long
    Dim obj As Object
    Dim mail As Outlook.MailItem

    On Error GoTo ErrHandler

    ' MoveするとItemsの件数が変わるため
    ' 後ろから処理する
    For i = tempFolder.Items.Count To 1 Step -1

        Set obj = tempFolder.Items.Item(i)

        If TypeOf obj Is Outlook.MailItem Then

            Set mail = obj

            ProcessReceivedMail mail

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
' 受信：
' メール1件を保存する処理
'============================================================

Private Sub ProcessReceivedMail( _
    ByVal mail As Outlook.MailItem _
)

    Dim saveFolder As String
    Dim fileName As String
    Dim fullPath As String

    Dim inbox As Outlook.Folder

    On Error GoTo ErrHandler

    '----------------------------------------
    ' 保存先
    '----------------------------------------
    saveFolder = GetReceiveSaveFolderPath()

    EnsureFolderExists saveFolder

    '----------------------------------------
    ' ファイル名
    '----------------------------------------
    fileName = BuildReceivedFileName(mail)

    fullPath = _
        saveFolder & "\" & fileName

    '----------------------------------------
    ' 同名ファイルが無ければ保存
    '----------------------------------------
    If Dir(fullPath) = "" Then

        mail.SaveAs _
            fullPath, _
            olMSG

    Else

        Debug.Print _
            "受信メール：重複のため保存をスキップ：" & _
            fullPath

    End If

    '----------------------------------------
    ' 分類項目「保存済み」を付与
    '----------------------------------------
    AddCategory _
        mail, _
        CATEGORY_NAME

    mail.Save

    '----------------------------------------
    ' 受信トレイへ戻す
    '----------------------------------------
    Set inbox = _
        Application.Session.GetDefaultFolder( _
            olFolderInbox _
        )

    mail.Move inbox

    Exit Sub


ErrHandler:

    ' 保存失敗時は受信トレイへ戻さない
    Debug.Print _
        "ProcessReceivedMail Error " & _
        Err.Number & " : " & _
        Err.Description

End Sub


'============================================================
' 受信：
' ファイル名作成
'
' yyyyMMdd_HHmmss_差出人_件名.msg
'============================================================

Private Function BuildReceivedFileName( _
    ByVal mail As Outlook.MailItem _
) As String

    Dim dateText As String
    Dim senderText As String
    Dim subjectText As String
    Dim result As String

    On Error GoTo ErrHandler

    '----------------------------------------
    ' 受信日時
    '----------------------------------------
    dateText = _
        Format( _
            mail.ReceivedTime, _
            "yyyymmdd_HHmmss" _
        )

    '----------------------------------------
    ' 差出人
    '----------------------------------------
    senderText = mail.SenderName

    If Trim(senderText) = "" Then
        senderText = "差出人不明"
    End If

    ' ( または （ 以降を削除
    senderText = _
        RemoveParenthesisPart(senderText)

    If Trim(senderText) = "" Then
        senderText = "差出人不明"
    End If

    senderText = _
        CleanFileName(senderText)

    '----------------------------------------
    ' 件名
    '----------------------------------------
    subjectText = mail.Subject

    If Trim(subjectText) = "" Then
        subjectText = "件名なし"
    End If

    subjectText = _
        CleanFileName(subjectText)

    '----------------------------------------
    ' ファイル名
    '----------------------------------------
    result = _
        dateText & "_" & _
        senderText & "_" & _
        subjectText

    '----------------------------------------
    ' 長すぎる場合は短縮
    '----------------------------------------
    If Len(result) > MAX_FILE_NAME_LENGTH Then

        result = _
            Left( _
                result, _
                MAX_FILE_NAME_LENGTH _
            )

    End If

    BuildReceivedFileName = _
        result & ".msg"

    Exit Function


ErrHandler:

    BuildReceivedFileName = _
        Format( _
            Now, _
            "yyyymmdd_HHmmss" _
        ) & _
        "_受信メール.msg"

End Function


'============================================================
' 送信：
' メールを送信したとき
'============================================================

Private Sub Application_ItemSend( _
    ByVal Item As Object, _
    Cancel As Boolean _
)

    On Error GoTo ErrHandler

    ' メール以外は処理しない
    If Not TypeOf Item Is Outlook.MailItem Then
        Exit Sub
    End If

    SaveSentMail Item

    Exit Sub


ErrHandler:

    ' 保存に失敗しても送信自体は止めない
    Debug.Print _
        "Application_ItemSend Error " & _
        Err.Number & " : " & _
        Err.Description

End Sub


'============================================================
' 送信：
' 送信メール保存処理
'============================================================

Private Sub SaveSentMail( _
    ByVal mail As Outlook.MailItem _
)

    Dim saveFolder As String
    Dim fileName As String
    Dim fullPath As String

    On Error GoTo ErrHandler

    '----------------------------------------
    ' 保存先
    '----------------------------------------
    saveFolder = _
        GetSentSaveFolderPath()

    EnsureFolderExists saveFolder

    '----------------------------------------
    ' ファイル名
    '----------------------------------------
    fileName = _
        BuildSentFileName(mail)

    fullPath = _
        saveFolder & "\" & fileName

    '----------------------------------------
    ' 同名ファイルが無ければ保存
    '----------------------------------------
    If Dir(fullPath) = "" Then

        mail.SaveAs _
            fullPath, _
            olMSG

    Else

        Debug.Print _
            "送信メール：重複のため保存をスキップ：" & _
            fullPath

    End If

    Exit Sub


ErrHandler:

    Debug.Print _
        "SaveSentMail Error " & _
        Err.Number & " : " & _
        Err.Description

End Sub


'============================================================
' 送信：
' ファイル名作成
'
' yyyyMMdd_HHmmss_宛先_件名.msg
'============================================================

Private Function BuildSentFileName( _
    ByVal mail As Outlook.MailItem _
) As String

    Dim dateText As String
    Dim recipientText As String
    Dim subjectText As String
    Dim result As String

    On Error GoTo ErrHandler

    '----------------------------------------
    ' 送信日時
    '
    ' ItemSendのタイミングなので
    ' 現在時刻を使用
    '----------------------------------------
    dateText = _
        Format( _
            Now, _
            "yyyymmdd_HHmmss" _
        )

    '----------------------------------------
    ' 宛先
    '----------------------------------------
    recipientText = _
        GetRecipientFileName(mail)

    '----------------------------------------
    ' 件名
    '----------------------------------------
    subjectText = mail.Subject

    If Trim(subjectText) = "" Then
        subjectText = "件名なし"
    End If

    subjectText = _
        CleanFileName(subjectText)

    '----------------------------------------
    ' ファイル名
    '----------------------------------------
    result = _
        dateText & "_" & _
        recipientText & "_" & _
        subjectText

    '----------------------------------------
    ' 長すぎる場合は短縮
    '----------------------------------------
    If Len(result) > MAX_FILE_NAME_LENGTH Then

        result = _
            Left( _
                result, _
                MAX_FILE_NAME_LENGTH _
            )

    End If

    BuildSentFileName = _
        result & ".msg"

    Exit Function


ErrHandler:

    BuildSentFileName = _
        Format( _
            Now, _
            "yyyymmdd_HHmmss" _
        ) & _
        "_送信メール.msg"

End Function


'============================================================
' 送信：
' ファイル名に使用する宛先名を取得
'
' 1名
'   山田太郎
'
' 複数名
'   山田太郎他
'
' ※Toのみを対象
' ※CC/BCCは人数に含めない
' ※「(」「（」以降は削除
'============================================================

Private Function GetRecipientFileName( _
    ByVal mail As Outlook.MailItem _
) As String

    Dim recipient As Outlook.Recipient

    Dim firstName As String
    Dim toCount As Long

    On Error GoTo ErrHandler

    firstName = ""
    toCount = 0

    '----------------------------------------
    ' 宛先を順番に確認
    '----------------------------------------
    For Each recipient In mail.Recipients

        ' Toのみ
        If recipient.Type = olTo Then

            toCount = _
                toCount + 1

            ' 最初の1名だけ取得
            If toCount = 1 Then

                firstName = _
                    recipient.Name

            End If

        End If

    Next recipient

    '----------------------------------------
    ' 宛先なし
    '----------------------------------------
    If Trim(firstName) = "" Then

        GetRecipientFileName = _
            "宛先なし"

        Exit Function

    End If

    '----------------------------------------
    ' ( または （ 以降を削除
    '----------------------------------------
    firstName = _
        RemoveParenthesisPart(firstName)

    If Trim(firstName) = "" Then

        firstName = _
            "宛先不明"

    End If

    '----------------------------------------
    ' ファイル名禁止文字除去
    '----------------------------------------
    firstName = _
        CleanFileName(firstName)

    '----------------------------------------
    ' 複数名なら「他」を追加
    '----------------------------------------
    If toCount >= 2 Then

        GetRecipientFileName = _
            firstName & "他"

    Else

        GetRecipientFileName = _
            firstName

    End If

    Exit Function


ErrHandler:

    GetRecipientFileName = _
        "宛先不明"

End Function


'============================================================
' 「(」または「（」以降を削除
'
' 例：
' 山田太郎(営業部)
' → 山田太郎
'
' 山田太郎（営業部）
' → 山田太郎
'============================================================

Private Function RemoveParenthesisPart( _
    ByVal text As String _
) As String

    Dim posHalf As Long
    Dim posFull As Long
    Dim cutPos As Long

    posHalf = _
        InStr( _
            1, _
            text, _
            "(", _
            vbBinaryCompare _
        )

    posFull = _
        InStr( _
            1, _
            text, _
            "（", _
            vbBinaryCompare _
        )

    cutPos = 0

    '----------------------------------------
    ' 半角と全角の両方がある
    '----------------------------------------
    If posHalf > 0 And posFull > 0 Then

        If posHalf < posFull Then

            cutPos = posHalf

        Else

            cutPos = posFull

        End If

    '----------------------------------------
    ' 半角だけ
    '----------------------------------------
    ElseIf posHalf > 0 Then

        cutPos = posHalf

    '----------------------------------------
    ' 全角だけ
    '----------------------------------------
    ElseIf posFull > 0 Then

        cutPos = posFull

    End If

    '----------------------------------------
    ' 括弧以降を削除
    '----------------------------------------
    If cutPos > 0 Then

        text = _
            Left( _
                text, _
                cutPos - 1 _
            )

    End If

    RemoveParenthesisPart = _
        Trim(text)

End Function


'============================================================
' ファイル名に使用できない文字を置換
'============================================================

Private Function CleanFileName( _
    ByVal text As String _
) As String

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

    '----------------------------------------
    ' 禁止文字を _ に置換
    '----------------------------------------
    For Each ch In invalidChars

        text = _
            Replace( _
                text, _
                CStr(ch), _
                "_" _
            )

    Next ch

    '----------------------------------------
    ' 改行・タブ
    '----------------------------------------
    text = _
        Replace( _
            text, _
            vbCr, _
            " " _
        )

    text = _
        Replace( _
            text, _
            vbLf, _
            " " _
        )

    text = _
        Replace( _
            text, _
            vbTab, _
            " " _
        )

    '----------------------------------------
    ' 前後の空白
    '----------------------------------------
    text = Trim(text)

    '----------------------------------------
    ' Windowsでは末尾の
    ' ピリオドや空白が問題になるため削除
    '----------------------------------------
    Do While Len(text) > 0

        If Right(text, 1) = "." Or _
           Right(text, 1) = " " Then

            text = _
                Left( _
                    text, _
                    Len(text) - 1 _
                )

        Else

            Exit Do

        End If

    Loop

    CleanFileName = text

End Function


'============================================================
' 受信メール保存先
'
' Windowsが認識しているドキュメント
'     ↓
' 受信メール
'============================================================

Private Function GetReceiveSaveFolderPath() As String

    Dim shell As Object
    Dim documentsPath As String

    Set shell = _
        CreateObject( _
            "WScript.Shell" _
        )

    documentsPath = _
        shell.SpecialFolders( _
            "MyDocuments" _
        )

    GetReceiveSaveFolderPath = _
        documentsPath & "\" & _
        RECEIVE_SAVE_FOLDER_NAME

    Set shell = Nothing

End Function


'============================================================
' 送信メール保存先
'
' Windowsが認識しているドキュメント
'     ↓
' 送信メール
'============================================================

Private Function GetSentSaveFolderPath() As String

    Dim shell As Object
    Dim documentsPath As String

    Set shell = _
        CreateObject( _
            "WScript.Shell" _
        )

    documentsPath = _
        shell.SpecialFolders( _
            "MyDocuments" _
        )

    GetSentSaveFolderPath = _
        documentsPath & "\" & _
        SENT_SAVE_FOLDER_NAME

    Set shell = Nothing

End Function


'============================================================
' Windowsフォルダが無ければ作成
'============================================================

Private Sub EnsureFolderExists( _
    ByVal folderPath As String _
)

    Dim fso As Object

    Set fso = _
        CreateObject( _
            "Scripting.FileSystemObject" _
        )

    If Not fso.FolderExists(folderPath) Then

        fso.CreateFolder folderPath

    End If

    Set fso = Nothing

End Sub


'============================================================
' Outlookのサブフォルダを取得
'============================================================

Private Function GetSubFolder( _
    ByVal parentFolder As Outlook.Folder, _
    ByVal folderName As String _
) As Outlook.Folder

    Dim folder As Outlook.Folder

    On Error Resume Next

    Set folder = _
        parentFolder.Folders(folderName)

    On Error GoTo 0

    Set GetSubFolder = folder

End Function


'============================================================
' 分類項目が存在しなければ作成
'============================================================

Private Sub EnsureCategoryExists( _
    ByVal categoryName As String _
)

    Dim category As Outlook.Category

    On Error Resume Next

    Set category = _
        Application.Session.Categories.Item( _
            categoryName _
        )

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
'
' 既存の分類項目は消さない
'============================================================

Private Sub AddCategory( _
    ByVal mail As Outlook.MailItem, _
    ByVal categoryName As String _
)

    Dim existingCategories As String

    On Error GoTo ErrHandler

    existingCategories = _
        Trim(mail.Categories)

    '----------------------------------------
    ' すでに付いていれば何もしない
    '----------------------------------------
    If HasCategory( _
        existingCategories, _
        categoryName _
    ) Then

        Exit Sub

    End If

    '----------------------------------------
    ' 分類なし
    '----------------------------------------
    If existingCategories = "" Then

        mail.Categories = _
            categoryName

    Else

        '----------------------------------------
        ' 既存分類を残して追加
        '----------------------------------------
        mail.Categories = _
            existingCategories & _
            "," & _
            categoryName

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

    categories = _
        Split( _
            categoriesText, _
            "," _
        )

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
