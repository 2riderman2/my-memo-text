# my-memo-text

Option Explicit

'============================================================
' 設定
'============================================================

' 受信トレイ直下に作成するフォルダ名
Private Const TEMP_FOLDER_NAME As String = "一時保存"

' ドキュメントフォルダ内の保存先フォルダ名
Private Const SAVE_FOLDER_NAME As String = "受信メール"

' 正常保存時または重複スキップ時の分類項目
Private Const SAVED_CATEGORY As String = "保存済み"

' エラー発生時の分類項目
Private Const ERROR_CATEGORY As String = "保存処理エラー"

' 保存先パスをメール内に記録するユーザー定義プロパティ
Private Const SAVE_PATH_PROPERTY As String = "VBA_MSG_SAVE_PATH"

' ファイル名の日付・時刻部分
'
' VBAでは、
' 月 = mm
' 分 = nn
'
' 作成例：
' 20260804_083105_差出人_件名.msg
Private Const DATE_FORMAT As String = "yyyymmdd_hhnnss"

' Windowsのパス長制限を考慮した安全側の上限
Private Const MAX_PATH_LENGTH As Long = 240

'============================================================
' イベント監視用変数
'============================================================

Private WithEvents mTemporaryItems As Outlook.Items

' 多重実行防止
Private mIsProcessing As Boolean

'============================================================
' Outlook起動時の処理
'============================================================

Private Sub Application_Startup()

    InitializeMailSaver

End Sub

'============================================================
' 初期化処理
'
' Outlook起動時に自動実行される。
' Alt + F8から手動実行することも可能。
'============================================================

Public Sub InitializeMailSaver()

    Dim tempFolder As Outlook.Folder

    On Error GoTo ErrorHandler

    ' 一時保存フォルダを取得
    Set tempFolder = GetTemporaryFolder()

    ' 一時保存フォルダのアイテム追加イベントを監視
    Set mTemporaryItems = tempFolder.Items

    ' 保存先フォルダを作成
    EnsureFolderExists GetSaveFolderPath()

    ' 分類項目を作成
    EnsureMasterCategory SAVED_CATEGORY
    EnsureMasterCategory ERROR_CATEGORY

    ' Outlook停止中などに残っていたメールも処理
    ProcessTemporaryFolderNow

    Exit Sub

ErrorHandler:

    MsgBox _
        "メール保存処理を初期化できませんでした。" & vbCrLf & vbCrLf & _
        "エラー番号：" & CStr(Err.Number) & vbCrLf & _
        "内容：" & Err.Description & vbCrLf & vbCrLf & _
        "受信トレイ直下に「" & TEMP_FOLDER_NAME & _
        "」フォルダが存在するか確認してください。", _
        vbExclamation, _
        "メール保存処理"

End Sub

'============================================================
' 一時保存フォルダにアイテムが追加されたとき
'============================================================

Private Sub mTemporaryItems_ItemAdd(ByVal Item As Object)

    ProcessTemporaryFolderNow

End Sub

'============================================================
' 一時保存フォルダ内のメールを一括処理
'
' Alt + F8から手動実行することも可能。
'============================================================

Public Sub ProcessTemporaryFolderNow()

    Dim tempFolder As Outlook.Folder
    Dim currentItem As Object
    Dim itemIndex As Long

    Dim errorNumber As Long
    Dim errorDescription As String

    If mIsProcessing Then Exit Sub

    mIsProcessing = True

    On Error GoTo ErrorHandler

    Set tempFolder = GetTemporaryFolder()

    EnsureFolderExists GetSaveFolderPath()

    ' 処理中にメールを移動するため末尾から処理する
    For itemIndex = tempFolder.Items.Count To 1 Step -1

        Set currentItem = tempFolder.Items.Item(itemIndex)

        If TypeOf currentItem Is Outlook.MailItem Then
            ProcessOneMail currentItem
        End If

        Set currentItem = Nothing

    Next itemIndex

CleanExit:

    mIsProcessing = False
    Exit Sub

ErrorHandler:

    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next

    WriteErrorLog _
        Nothing, _
        "一時保存フォルダ一括処理", _
        errorNumber, _
        errorDescription, _
        ""

    On Error GoTo 0

    GoTo CleanExit

End Sub

'============================================================
' メール1件の保存処理
'============================================================

Private Sub ProcessOneMail(ByVal mail As Outlook.MailItem)

    Dim saveFolderPath As String
    Dim savePath As String
    Dim registeredPath As String

    Dim senderText As String
    Dim subjectText As String
    Dim fileStem As String

    Dim inboxFolder As Outlook.Folder

    Dim errorNumber As Long
    Dim errorDescription As String
    Dim fileExistsAfterError As Boolean

    On Error GoTo ErrorHandler

    saveFolderPath = GetSaveFolderPath()

    EnsureFolderExists saveFolderPath

    Set inboxFolder = GetInboxFolder()

    '--------------------------------------------------------
    ' 以前の処理で保存先パスがメール内に記録されている場合
    '
    ' MSG保存後、受信トレイへの移動だけ失敗した場合などは、
    ' MSGを再保存せず、受信トレイへの移動だけを行う。
    '--------------------------------------------------------

    registeredPath = GetRegisteredSavePath(mail)

    If Len(registeredPath) > 0 Then

        savePath = registeredPath

        If FileExists(registeredPath) Then

            RemoveCategoryIfPresent mail, ERROR_CATEGORY
            AddCategoryIfMissing mail, SAVED_CATEGORY

            mail.Save
            mail.Move inboxFolder

            Exit Sub

        Else

            ' 記録されたファイルが存在しない場合は記録を削除
            ClearRegisteredSavePath mail

        End If

    End If

    '--------------------------------------------------------
    ' ファイル名の作成
    '--------------------------------------------------------

    senderText = CleanFileComponent( _
        mail.SenderName, _
        "差出人不明")

    subjectText = CleanFileComponent( _
        mail.Subject, _
        "件名なし")

    ' 差出人名が長すぎる場合に短縮
    senderText = Left$(senderText, 60)

    fileStem = _
        Format$(mail.ReceivedTime, DATE_FORMAT) & "_" & _
        senderText & "_" & _
        subjectText

    ' パス全体が長くなりすぎないように調整
    fileStem = FitFileStemToPath( _
        saveFolderPath, _
        fileStem, _
        ".msg")

    savePath = _
        saveFolderPath & "\" & fileStem & ".msg"

    '--------------------------------------------------------
    ' 同名ファイルが存在する場合
    '
    ' MSG保存は行わず、既存ファイルを保存済みとみなす。
    '--------------------------------------------------------

    If FileExists(savePath) Then

        RemoveCategoryIfPresent mail, ERROR_CATEGORY
        AddCategoryIfMissing mail, SAVED_CATEGORY

        ' 既存ファイルのパスをメール内に記録
        SetRegisteredSavePath mail, savePath

        mail.Save

        ' 重複スキップログを出力
        WriteDuplicateLog mail, savePath

        ' 受信トレイへ戻す
        mail.Move inboxFolder

        Exit Sub

    End If

    '--------------------------------------------------------
    ' 通常の保存処理
    '--------------------------------------------------------

    RemoveCategoryIfPresent mail, ERROR_CATEGORY
    AddCategoryIfMissing mail, SAVED_CATEGORY

    ' 保存予定パスをメール内に記録
    SetRegisteredSavePath mail, savePath

    ' 分類項目とユーザー定義プロパティを保存
    mail.Save

    ' Unicode対応のMSG形式で保存
    mail.SaveAs savePath, olMSGUnicode

    ' 保存成功後に受信トレイへ移動
    mail.Move inboxFolder

    Exit Sub

ErrorHandler:

    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next

    ' savePathが未設定ならメール内の記録を確認
    If Len(savePath) = 0 Then
        savePath = GetRegisteredSavePath(mail)
    End If

    fileExistsAfterError = False

    If Len(savePath) > 0 Then
        fileExistsAfterError = FileExists(savePath)
    End If

    ' MSGファイルが作成されていない場合は保存済み扱いを取り消す
    If Not fileExistsAfterError Then

        ClearRegisteredSavePath mail
        RemoveCategoryIfPresent mail, SAVED_CATEGORY

    End If

    ' エラーカテゴリを設定
    AddCategoryIfMissing mail, ERROR_CATEGORY

    mail.Save

    ' エラーログを出力
    WriteErrorLog _
        mail, _
        "メール保存処理", _
        errorNumber, _
        errorDescription, _
        savePath

    On Error GoTo 0

End Sub

'============================================================
' 一時保存フォルダを取得
'
' 想定するフォルダ構成：
'
' 受信トレイ
' └ 一時保存
'============================================================

Private Function GetTemporaryFolder() As Outlook.Folder

    Dim inboxFolder As Outlook.Folder
    Dim tempFolder As Outlook.Folder

    Set inboxFolder = _
        Application.Session.GetDefaultFolder(olFolderInbox)

    On Error Resume Next

    Set tempFolder = _
        inboxFolder.Folders(TEMP_FOLDER_NAME)

    On Error GoTo 0

    If tempFolder Is Nothing Then

        Err.Raise _
            vbObjectError + 1001, _
            "GetTemporaryFolder", _
            "受信トレイ直下に「" & TEMP_FOLDER_NAME & _
            "」フォルダが見つかりません。"

    End If

    Set GetTemporaryFolder = tempFolder

End Function

'============================================================
' 受信トレイを取得
'============================================================

Private Function GetInboxFolder() As Outlook.Folder

    Set GetInboxFolder = _
        Application.Session.GetDefaultFolder(olFolderInbox)

End Function

'============================================================
' 保存先フォルダを取得
'
' Windowsが認識している実際のドキュメントフォルダを取得する。
'
' 例：
' C:\Users\ユーザー名\OneDrive - 会社名\Documents\受信メール
'============================================================

Private Function GetSaveFolderPath() As String

    Dim shellObject As Object
    Dim documentsPath As String

    Set shellObject = CreateObject("WScript.Shell")

    documentsPath = _
        CStr(shellObject.SpecialFolders("MyDocuments"))

    If Len(Trim$(documentsPath)) = 0 Then

        Err.Raise _
            vbObjectError + 1002, _
            "GetSaveFolderPath", _
            "Windowsのドキュメントフォルダを取得できませんでした。"

    End If

    GetSaveFolderPath = _
        documentsPath & "\" & SAVE_FOLDER_NAME

End Function

'============================================================
' 保存先を確認するためのマクロ
'
' Alt + F8からShowSaveFolderPathを実行すると、
' 実際の保存先が表示される。
'============================================================

Public Sub ShowSaveFolderPath()

    MsgBox _
        "メールの保存先は次のフォルダです。" & vbCrLf & vbCrLf & _
        GetSaveFolderPath(), _
        vbInformation, _
        "MSG保存先"

End Sub

'============================================================
' フォルダが存在しない場合は作成
'============================================================

Private Sub EnsureFolderExists(ByVal folderPath As String)

    Dim fileSystem As Object
    Dim parentPath As String

    Set fileSystem = _
        CreateObject("Scripting.FileSystemObject")

    If fileSystem.FolderExists(folderPath) Then Exit Sub

    parentPath = _
        fileSystem.GetParentFolderName(folderPath)

    If Len(parentPath) > 0 Then

        If Not fileSystem.FolderExists(parentPath) Then
            EnsureFolderExists parentPath
        End If

    End If

    fileSystem.CreateFolder folderPath

End Sub

'============================================================
' ファイルの存在確認
'============================================================

Private Function FileExists(ByVal filePath As String) As Boolean

    Dim fileSystem As Object

    If Len(filePath) = 0 Then Exit Function

    Set fileSystem = _
        CreateObject("Scripting.FileSystemObject")

    FileExists = _
        fileSystem.FileExists(filePath)

End Function

'============================================================
' ファイル名として使用できない文字を除去・置換
'============================================================

Private Function CleanFileComponent( _
    ByVal sourceText As String, _
    ByVal fallbackText As String) As String

    Dim invalidCharacters As Variant
    Dim invalidCharacter As Variant

    Dim characterCode As Long
    Dim cleanedText As String

    cleanedText = sourceText

    ' 制御文字を空白に置換
    For characterCode = 0 To 31

        cleanedText = _
            Replace( _
                cleanedText, _
                Chr$(characterCode), _
                " ")

    Next characterCode

    ' Windowsファイル名の禁則文字
    invalidCharacters = _
        Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For Each invalidCharacter In invalidCharacters

        cleanedText = _
            Replace( _
                cleanedText, _
                CStr(invalidCharacter), _
                "_")

    Next invalidCharacter

    cleanedText = Trim$(cleanedText)

    ' 連続する空白を1つにまとめる
    Do While InStr(cleanedText, "  ") > 0

        cleanedText = _
            Replace(cleanedText, "  ", " ")

    Loop

    ' ファイル名末尾のピリオドと空白を削除
    Do While Len(cleanedText) > 0

        If Right$(cleanedText, 1) = "." _
            Or Right$(cleanedText, 1) = " " Then

            cleanedText = _
                Left$(cleanedText, Len(cleanedText) - 1)

        Else

            Exit Do

        End If

    Loop

    If Len(cleanedText) = 0 Then
        cleanedText = fallbackText
    End If

    CleanFileComponent = cleanedText

End Function

'============================================================
' パス全体が長くなりすぎないようにファイル名を短縮
'============================================================

Private Function FitFileStemToPath( _
    ByVal folderPath As String, _
    ByVal fileStem As String, _
    ByVal extensionText As String) As String

    Dim maximumStemLength As Long
    Dim adjustedStem As String

    maximumStemLength = _
        MAX_PATH_LENGTH _
        - Len(folderPath) _
        - 1 _
        - Len(extensionText)

    If maximumStemLength < 30 Then

        Err.Raise _
            vbObjectError + 1003, _
            "FitFileStemToPath", _
            "保存先フォルダのパスが長すぎます。"

    End If

    adjustedStem = fileStem

    If Len(adjustedStem) > maximumStemLength Then

        adjustedStem = _
            Left$(adjustedStem, maximumStemLength)

    End If

    ' 短縮後の末尾のピリオドと空白を削除
    Do While Len(adjustedStem) > 0

        If Right$(adjustedStem, 1) = "." _
            Or Right$(adjustedStem, 1) = " " Then

            adjustedStem = _
                Left$(adjustedStem, Len(adjustedStem) - 1)

        Else

            Exit Do

        End If

    Loop

    FitFileStemToPath = adjustedStem

End Function

'============================================================
' 保存先パスをメール内に記録
'============================================================

Private Sub SetRegisteredSavePath( _
    ByVal mail As Outlook.MailItem, _
    ByVal savePath As String)

    Dim userProperty As Outlook.UserProperty

    Set userProperty = _
        mail.UserProperties.Find( _
            SAVE_PATH_PROPERTY, _
            True)

    If userProperty Is Nothing Then

        Set userProperty = _
            mail.UserProperties.Add( _
                SAVE_PATH_PROPERTY, _
                olText, _
                False)

    End If

    userProperty.Value = savePath

End Sub

'============================================================
' メール内に記録された保存先パスを取得
'============================================================

Private Function GetRegisteredSavePath( _
    ByVal mail As Outlook.MailItem) As String

    Dim userProperty As Outlook.UserProperty

    Set userProperty = _
        mail.UserProperties.Find( _
            SAVE_PATH_PROPERTY, _
            True)

    If Not userProperty Is Nothing Then

        GetRegisteredSavePath = _
            CStr(userProperty.Value)

    End If

End Function

'============================================================
' メール内に記録された保存先パスを削除
'============================================================

Private Sub ClearRegisteredSavePath( _
    ByVal mail As Outlook.MailItem)

    Dim userProperty As Outlook.UserProperty

    Set userProperty = _
        mail.UserProperties.Find( _
            SAVE_PATH_PROPERTY, _
            True)

    If Not userProperty Is Nothing Then
        userProperty.Delete
    End If

End Sub

'============================================================
' Outlookの分類項目一覧に分類項目を作成
'============================================================

Private Sub EnsureMasterCategory(ByVal categoryName As String)

    Dim outlookCategory As Outlook.Category

    On Error Resume Next

    Set outlookCategory = _
        Application.Session.Categories.Item(categoryName)

    If outlookCategory Is Nothing Then

        Set outlookCategory = _
            Application.Session.Categories.Add(categoryName)

    End If

    On Error GoTo 0

End Sub

'============================================================
' 指定した分類項目が設定されているか確認
'============================================================

Private Function HasCategory( _
    ByVal categoryText As String, _
    ByVal categoryName As String) As Boolean

    Dim separatorText As String
    Dim categoryArray As Variant
    Dim categoryValue As Variant

    If Len(Trim$(categoryText)) = 0 Then Exit Function

    separatorText = GetCategorySeparator()
    categoryArray = Split(categoryText, separatorText)

    For Each categoryValue In categoryArray

        If StrComp( _
            Trim$(CStr(categoryValue)), _
            categoryName, _
            vbTextCompare) = 0 Then

            HasCategory = True
            Exit Function

        End If

    Next categoryValue

End Function

'============================================================
' 分類項目を追加
'
' 既存の分類項目は残す。
'============================================================

Private Function AddCategoryIfMissing( _
    ByVal mail As Outlook.MailItem, _
    ByVal categoryName As String) As Boolean

    Dim separatorText As String

    If HasCategory(mail.Categories, categoryName) Then
        Exit Function
    End If

    separatorText = GetCategorySeparator()

    If Len(Trim$(mail.Categories)) = 0 Then

        mail.Categories = categoryName

    Else

        mail.Categories = _
            mail.Categories & _
            separatorText & _
            categoryName

    End If

    AddCategoryIfMissing = True

End Function

'============================================================
' 指定した分類項目を削除
'
' その他の分類項目は残す。
'============================================================

Private Sub RemoveCategoryIfPresent( _
    ByVal mail As Outlook.MailItem, _
    ByVal categoryName As String)

    Dim separatorText As String
    Dim categoryArray As Variant
    Dim categoryValue As Variant

    Dim rebuiltCategories As String
    Dim currentCategory As String

    If Len(Trim$(mail.Categories)) = 0 Then Exit Sub

    separatorText = GetCategorySeparator()
    categoryArray = Split(mail.Categories, separatorText)

    For Each categoryValue In categoryArray

        currentCategory = Trim$(CStr(categoryValue))

        If Len(currentCategory) > 0 Then

            If StrComp( _
                currentCategory, _
                categoryName, _
                vbTextCompare) <> 0 Then

                If Len(rebuiltCategories) > 0 Then

                    rebuiltCategories = _
                        rebuiltCategories & separatorText

                End If

                rebuiltCategories = _
                    rebuiltCategories & currentCategory

            End If

        End If

    Next categoryValue

    mail.Categories = rebuiltCategories

End Sub

'============================================================
' Windowsのリスト区切り文字を取得
'============================================================

Private Function GetCategorySeparator() As String

    Dim registryShell As Object
    Dim separatorText As String

    On Error Resume Next

    Set registryShell = CreateObject("WScript.Shell")

    separatorText = _
        CStr(registryShell.RegRead( _
            "HKCU\Control Panel\International\sList"))

    On Error GoTo 0

    If Len(separatorText) = 0 Then
        separatorText = ","
    End If

    GetCategorySeparator = separatorText

End Function

'============================================================
' エラーログを出力
'
' 保存先：
' Documents\受信メール\保存エラー.log
'============================================================

Private Sub WriteErrorLog( _
    ByVal mail As Outlook.MailItem, _
    ByVal processingStage As String, _
    ByVal errorNumber As Long, _
    ByVal errorDescription As String, _
    ByVal savePath As String)

    Dim logFolderPath As String
    Dim logFilePath As String
    Dim fileNumber As Integer

    Dim mailSubject As String
    Dim mailEntryID As String
    Dim senderName As String

    On Error Resume Next

    logFolderPath = GetSaveFolderPath()

    EnsureFolderExists logFolderPath

    logFilePath = _
        logFolderPath & "\保存エラー.log"

    If Not mail Is Nothing Then

        mailSubject = CleanLogText(mail.Subject)
        mailEntryID = CleanLogText(mail.EntryID)
        senderName = CleanLogText(mail.SenderName)

    Else

        mailSubject = ""
        mailEntryID = ""
        senderName = ""

    End If

    fileNumber = FreeFile

    Open logFilePath For Append As #fileNumber

    Print #fileNumber, _
        Format$(Now, "yyyy/mm/dd hh:nn:ss") & vbTab & _
        processingStage & vbTab & _
        "エラー番号=" & CStr(errorNumber) & vbTab & _
        "内容=" & CleanLogText(errorDescription) & vbTab & _
        "差出人=" & senderName & vbTab & _
        "件名=" & mailSubject & vbTab & _
        "保存先=" & CleanLogText(savePath) & vbTab & _
        "EntryID=" & mailEntryID

    Close #fileNumber

    On Error GoTo 0

End Sub

'============================================================
' 同名ファイルのため保存をスキップした履歴を出力
'
' 保存先：
' Documents\受信メール\重複スキップ.log
'============================================================

Private Sub WriteDuplicateLog( _
    ByVal mail As Outlook.MailItem, _
    ByVal existingFilePath As String)

    Dim logFolderPath As String
    Dim logFilePath As String
    Dim fileNumber As Integer

    On Error Resume Next

    logFolderPath = GetSaveFolderPath()

    EnsureFolderExists logFolderPath

    logFilePath = _
        logFolderPath & "\重複スキップ.log"

    fileNumber = FreeFile

    Open logFilePath For Append As #fileNumber

    Print #fileNumber, _
        Format$(Now, "yyyy/mm/dd hh:nn:ss") & vbTab & _
        "同名ファイルのためMSG保存をスキップ" & vbTab & _
        "受信日時=" & _
            Format$(mail.ReceivedTime, "yyyy/mm/dd hh:nn:ss") & vbTab & _
        "差出人=" & CleanLogText(mail.SenderName) & vbTab & _
        "件名=" & CleanLogText(mail.Subject) & vbTab & _
        "既存ファイル=" & CleanLogText(existingFilePath) & vbTab & _
        "EntryID=" & CleanLogText(mail.EntryID)

    Close #fileNumber

    On Error GoTo 0

End Sub

'============================================================
' ログ内に改行やタブが入らないようにする
'============================================================

Private Function CleanLogText( _
    ByVal sourceText As String) As String

    Dim cleanedText As String

    cleanedText = sourceText

    cleanedText = Replace(cleanedText, vbCr, " ")
    cleanedText = Replace(cleanedText, vbLf, " ")
    cleanedText = Replace(cleanedText, vbTab, " ")

    CleanLogText = cleanedText

End Function































    ' 同名ファイルがすでに存在する場合
    '
    ' 新しいMSGファイルは保存しない。
    ' 既存ファイルを保存済みファイルとみなして、
    ' 保存済みカテゴリを設定し、受信トレイへ戻す。
    '--------------------------------------------------------

    If FileExists(savePath) Then

        RemoveCategoryIfPresent mail, ERROR_CATEGORY
        AddCategoryIfMissing mail, SAVED_CATEGORY

        ' 既存ファイルのパスをメール内部に記録する
        SetRegisteredSavePath mail, savePath

        mail.Save

        ' 重複スキップの履歴をログに記録する
        WriteDuplicateLog mail, savePath

        mail.Move inboxFolder

        Exit Sub

    End If

    '--------------------------------------------------------
    ' 通常のMSG保存処理
    '--------------------------------------------------------

    RemoveCategoryIfPresent mail, ERROR_CATEGORY
    AddCategoryIfMissing mail, SAVED_CATEGORY

    ' 保存予定のファイルパスをメール内部に記録する
    SetRegisteredSavePath mail, savePath

    ' 分類項目とユーザー定義プロパティをメールに反映する
    mail.Save

    ' Unicode対応のMSG形式で保存する
    mail.SaveAs savePath, olMSGUnicode

    ' 保存成功後、受信トレイへ移動する
    mail.Move inboxFolder

    Exit Sub

ErrorHandler:

    Dim errorNumber As Long
    Dim errorDescription As String
    Dim fileExistsAfterError As Boolean

    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next

    ' savePathがまだ設定されていない場合は、
    ' メール内部に保存されたパスを取得する
    If Len(savePath) = 0 Then
        savePath = GetRegisteredSavePath(mail)
    End If

    fileExistsAfterError = False

    If Len(savePath) > 0 Then
        fileExistsAfterError = FileExists(savePath)
    End If

    If Not fileExistsAfterError Then

        ' MSGファイルが存在しない場合は保存済み扱いを取り消す
        ClearRegisteredSavePath mail
        RemoveCategoryIfPresent mail, SAVED_CATEGORY

    End If

    ' 一時保存フォルダに残したまま、
    ' エラーが発生したことを分類項目で示す
    AddCategoryIfMissing mail, ERROR_CATEGORY

    mail.Save

    WriteErrorLog _
        mail, _
        "メール保存処理", _
        errorNumber, _
        errorDescription, _
        savePath

    On Error GoTo 0

End Sub

'============================================================
' 一時保存フォルダを取得
'
' 想定する構成：
'
' 受信トレイ
' └ 一時保存
'============================================================

Private Function GetTemporaryFolder() As Outlook.Folder

    Dim inboxFolder As Outlook.Folder

    Set inboxFolder = _
        Application.Session.GetDefaultFolder(olFolderInbox)

    On Error Resume Next

    Set GetTemporaryFolder = _
        inboxFolder.Folders(TEMP_FOLDER_NAME)

    On Error GoTo 0

    If GetTemporaryFolder Is Nothing Then

        Err.Raise _
            vbObjectError + 1001, _
            "GetTemporaryFolder", _
            "受信トレイ直下に「" & TEMP_FOLDER_NAME & _
            "」フォルダが見つかりません。"

    End If

End Function

'============================================================
' 受信トレイを取得
'============================================================

Private Function GetInboxFolder() As Outlook.Folder

    Set GetInboxFolder = _
        Application.Session.GetDefaultFolder(olFolderInbox)

End Function

'============================================================
' 保存先フォルダのパスを取得
'
' C:\Users\ユーザー名\Documents\受信メール
'============================================================

Private Function GetSaveFolderPath() As String

    GetSaveFolderPath = _
        Environ$("USERPROFILE") & _
        "\Documents\" & _
        SAVE_FOLDER_NAME

End Function

'============================================================
' フォルダが存在しなければ作成
'============================================================

Private Sub EnsureFolderExists(ByVal folderPath As String)

    Dim fileSystem As Object
    Dim parentPath As String

    Set fileSystem = _
        CreateObject("Scripting.FileSystemObject")

    If fileSystem.FolderExists(folderPath) Then Exit Sub

    parentPath = fileSystem.GetParentFolderName(folderPath)

    If Len(parentPath) > 0 Then

        If Not fileSystem.FolderExists(parentPath) Then
            EnsureFolderExists parentPath
        End If

    End If

    fileSystem.CreateFolder folderPath

End Sub

'============================================================
' ファイルが存在するか確認
'============================================================

Private Function FileExists(ByVal filePath As String) As Boolean

    Dim fileSystem As Object

    If Len(filePath) = 0 Then Exit Function

    Set fileSystem = _
        CreateObject("Scripting.FileSystemObject")

    FileExists = fileSystem.FileExists(filePath)

End Function

'============================================================
' ファイル名に使用できない文字を除去・置換
'============================================================

Private Function CleanFileComponent( _
    ByVal sourceText As String, _
    ByVal fallbackText As String) As String

    Dim invalidCharacters As Variant
    Dim invalidCharacter As Variant
    Dim characterCode As Long
    Dim cleanedText As String

    cleanedText = sourceText

    ' ASCII制御文字を空白へ置換する
    For characterCode = 0 To 31

        cleanedText = _
            Replace( _
                cleanedText, _
                Chr$(characterCode), _
                " ")

    Next characterCode

    ' Windowsのファイル名で使用できない文字
    invalidCharacters = _
        Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For Each invalidCharacter In invalidCharacters

        cleanedText = _
            Replace( _
                cleanedText, _
                CStr(invalidCharacter), _
                "_")

    Next invalidCharacter

    cleanedText = Trim$(cleanedText)

    ' 連続した空白を1文字へまとめる
    Do While InStr(cleanedText, "  ") > 0
        cleanedText = Replace(cleanedText, "  ", " ")
    Loop

    ' ファイル名末尾のピリオドと空白を削除する
    Do While Len(cleanedText) > 0

        If Right$(cleanedText, 1) = "." _
            Or Right$(cleanedText, 1) = " " Then

            cleanedText = _
                Left$(cleanedText, Len(cleanedText) - 1)

        Else

            Exit Do

        End If

    Loop

    If Len(cleanedText) = 0 Then
        cleanedText = fallbackText
    End If

    CleanFileComponent = cleanedText

End Function

'============================================================
' パス全体が長くなりすぎないようファイル名を短縮
'============================================================

Private Function FitFileStemToPath( _
    ByVal folderPath As String, _
    ByVal fileStem As String, _
    ByVal extensionText As String) As String

    Dim maximumStemLength As Long
    Dim adjustedStem As String

    maximumStemLength = _
        MAX_PATH_LENGTH _
        - Len(folderPath) _
        - 1 _
        - Len(extensionText)

    If maximumStemLength < 30 Then

        Err.Raise _
            vbObjectError + 1002, _
            "FitFileStemToPath", _
            "保存先フォルダのパスが長すぎます。"

    End If

    adjustedStem = fileStem

    If Len(adjustedStem) > maximumStemLength Then

        adjustedStem = _
            Left$(adjustedStem, maximumStemLength)

    End If

    ' 短縮後の末尾にピリオドや空白が残らないようにする
    Do While Len(adjustedStem) > 0

        If Right$(adjustedStem, 1) = "." _
            Or Right$(adjustedStem, 1) = " " Then

            adjustedStem = _
                Left$(adjustedStem, Len(adjustedStem) - 1)

        Else

            Exit Do

        End If

    Loop

    FitFileStemToPath = adjustedStem

End Function

'============================================================
' 保存先パスをメール内部に記録
'============================================================

Private Sub SetRegisteredSavePath( _
    ByVal mail As Outlook.MailItem, _
    ByVal savePath As String)

    Dim userProperty As Outlook.UserProperty

    Set userProperty = _
        mail.UserProperties.Find( _
            SAVE_PATH_PROPERTY, _
            True)

    If userProperty Is Nothing Then

        Set userProperty = _
            mail.UserProperties.Add( _
                SAVE_PATH_PROPERTY, _
                olText, _
                False)

    End If

    userProperty.Value = savePath

End Sub

'============================================================
' メール内部に記録された保存先パスを取得
'============================================================

Private Function GetRegisteredSavePath( _
    ByVal mail As Outlook.MailItem) As String

    Dim userProperty As Outlook.UserProperty

    Set userProperty = _
        mail.UserProperties.Find( _
            SAVE_PATH_PROPERTY, _
            True)

    If Not userProperty Is Nothing Then
        GetRegisteredSavePath = CStr(userProperty.Value)
    End If

End Function

'============================================================
' メール内部に記録された保存先パスを削除
'============================================================

Private Sub ClearRegisteredSavePath( _
    ByVal mail As Outlook.MailItem)

    Dim userProperty As Outlook.UserProperty

    Set userProperty = _
        mail.UserProperties.Find( _
            SAVE_PATH_PROPERTY, _
            True)

    If Not userProperty Is Nothing Then
        userProperty.Delete
    End If

End Sub

'============================================================
' Outlookのマスター分類項目一覧に分類項目を作成
'============================================================

Private Sub EnsureMasterCategory( _
    ByVal categoryName As String)

    Dim outlookCategory As Outlook.Category

    On Error Resume Next

    Set outlookCategory = _
        Application.Session.Categories.Item(categoryName)

    If outlookCategory Is Nothing Then

        Set outlookCategory = _
            Application.Session.Categories.Add(categoryName)

    End If

    On Error GoTo 0

End Sub

'============================================================
' 指定された分類項目が設定されているか確認
'============================================================

Private Function HasCategory( _
    ByVal categoryText As String, _
    ByVal categoryName As String) As Boolean

    Dim separatorText As String
    Dim categoryArray As Variant
    Dim categoryValue As Variant

    If Len(Trim$(categoryText)) = 0 Then Exit Function

    separatorText = GetCategorySeparator()
    categoryArray = Split(categoryText, separatorText)

    For Each categoryValue In categoryArray

        If StrComp( _
            Trim$(CStr(categoryValue)), _
            categoryName, _
            vbTextCompare) = 0 Then

            HasCategory = True
            Exit Function

        End If

    Next categoryValue

End Function

'============================================================
' 分類項目を追加
'
' 既存の分類項目は削除しない。
'============================================================

Private Function AddCategoryIfMissing( _
    ByVal mail As Outlook.MailItem, _
    ByVal categoryName As String) As Boolean

    Dim separatorText As String

    If HasCategory(mail.Categories, categoryName) Then
        Exit Function
    End If

    separatorText = GetCategorySeparator()

    If Len(Trim$(mail.Categories)) = 0 Then

        mail.Categories = categoryName

    Else

        mail.Categories = _
            mail.Categories & _
            separatorText & _
            categoryName

    End If

    AddCategoryIfMissing = True

End Function

'============================================================
' 指定された分類項目を削除
'
' それ以外の分類項目は残す。
'============================================================

Private Sub RemoveCategoryIfPresent( _
    ByVal mail As Outlook.MailItem, _
    ByVal categoryName As String)

    Dim separatorText As String
    Dim categoryArray As Variant
    Dim categoryValue As Variant
    Dim rebuiltCategories As String
    Dim currentCategory As String

    If Len(Trim$(mail.Categories)) = 0 Then Exit Sub

    separatorText = GetCategorySeparator()
    categoryArray = Split(mail.Categories, separatorText)

    For Each categoryValue In categoryArray

        currentCategory = Trim$(CStr(categoryValue))

        If Len(currentCategory) > 0 Then

            If StrComp( _
                currentCategory, _
                categoryName, _
                vbTextCompare) <> 0 Then

                If Len(rebuiltCategories) > 0 Then

                    rebuiltCategories = _
                        rebuiltCategories & separatorText

                End If

                rebuiltCategories = _
                    rebuiltCategories & currentCategory

            End If

        End If

    Next categoryValue

    mail.Categories = rebuiltCategories

End Sub

'============================================================
' Windowsのリスト区切り記号を取得
'
' 日本の一般的な環境では「,」
'============================================================

Private Function GetCategorySeparator() As String

    Dim registryShell As Object
    Dim separatorText As String

    On Error Resume Next

    Set registryShell = CreateObject("WScript.Shell")

    separatorText = _
        registryShell.RegRead( _
            "HKCU\Control Panel\International\sList")

    On Error GoTo 0

    If Len(separatorText) = 0 Then
        separatorText = ","
    End If

    GetCategorySeparator = separatorText

End Function

'============================================================
' エラーログを出力
'
' 保存先：
' Documents\受信メール\保存エラー.log
'============================================================

Private Sub WriteErrorLog( _
    ByVal mail As Outlook.MailItem, _
    ByVal processingStage As String, _
    ByVal errorNumber As Long, _
    ByVal errorDescription As String, _
    ByVal savePath As String)

    On Error Resume Next

    Dim logFolderPath As String
    Dim logFilePath As String
    Dim fileNumber As Integer
    Dim mailSubject As String
    Dim mailEntryID As String
    Dim senderName As String

    logFolderPath = GetSaveFolderPath()

    EnsureFolderExists logFolderPath

    logFilePath = _
        logFolderPath & "\保存エラー.log"

    If Not mail Is Nothing Then

        mailSubject = CleanLogText(mail.Subject)
        mailEntryID = CleanLogText(mail.EntryID)
        senderName = CleanLogText(mail.SenderName)

    Else

        mailSubject = ""
        mailEntryID = ""
        senderName = ""

    End If

    fileNumber = FreeFile

    Open logFilePath For Append As #fileNumber

    Print #fileNumber, _
        Format$(Now, "yyyy/mm/dd hh:nn:ss") & vbTab & _
        processingStage & vbTab & _
        "エラー番号=" & CStr(errorNumber) & vbTab & _
        "内容=" & CleanLogText(errorDescription) & vbTab & _
        "差出人=" & senderName & vbTab & _
        "件名=" & mailSubject & vbTab & _
        "保存先=" & CleanLogText(savePath) & vbTab & _
        "EntryID=" & mailEntryID

    Close #fileNumber

    On Error GoTo 0

End Sub

'============================================================
' 同名ファイルのため保存をスキップした履歴を出力
'
' 保存先：
' Documents\受信メール\重複スキップ.log
'============================================================

Private Sub WriteDuplicateLog( _
    ByVal mail As Outlook.MailItem, _
    ByVal existingFilePath As String)

    On Error Resume Next

    Dim logFolderPath As String
    Dim logFilePath As String
    Dim fileNumber As Integer

    logFolderPath = GetSaveFolderPath()

    EnsureFolderExists logFolderPath

    logFilePath = _
        logFolderPath & "\重複スキップ.log"

    fileNumber = FreeFile

    Open logFilePath For Append As #fileNumber

    Print #fileNumber, _
        Format$(Now, "yyyy/mm/dd hh:nn:ss") & vbTab & _
        "同名ファイルのためMSG保存をスキップ" & vbTab & _
        "受信日時=" & _
            Format$(mail.ReceivedTime, "yyyy/mm/dd hh:nn:ss") & vbTab & _
        "差出人=" & CleanLogText(mail.SenderName) & vbTab & _
        "件名=" & CleanLogText(mail.Subject) & vbTab & _
        "既存ファイル=" & CleanLogText(existingFilePath) & vbTab & _
        "EntryID=" & CleanLogText(mail.EntryID)

    Close #fileNumber

    On Error GoTo 0

End Sub

'============================================================
' ログに改行やタブが入らないようにする
'============================================================

Private Function CleanLogText( _
    ByVal sourceText As String) As String

    Dim cleanedText As String

    cleanedText = sourceText

    cleanedText = Replace(cleanedText, vbCr, " ")
    cleanedText = Replace(cleanedText, vbLf, " ")
    cleanedText = Replace(cleanedText, vbTab, " ")

    CleanLogText = cleanedText

End Function



{
  "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json",
  "elmType": "a",
  "style": {
    "display": "=if(@currentField == '', 'none', 'inline-flex')",
    "align-items": "center",
    "gap": "6px",
    "text-decoration": "none",
    "font-weight": "600"
  },
  "attributes": {
    "href": "@currentField",
    "target": "_blank",
    "title": "リンク先を新しいタブで開きます"
  },
  "children": [
    {
      "elmType": "span",
      "attributes": {
        "iconName": "OpenInNewWindow"
      }
    },
    {
      "elmType": "span",
      "txtContent": "リンクを開く"
    }
  ]
}




{
  "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json",
  "elmType": "span",
  "txtContent": "=if([$利用者名] != '', [$利用者名], if(indexOf([$Author.title], ' ') >= 0, substring([$Author.title], 0, indexOf([$Author.title], ' ')), [$Author.title]))"
}





let
    PDFファイルパス =
        "C:\資料\対象ファイル.pdf",

    // PDF内のTable・Page一覧を取得
    ソース =
        Pdf.Tables(
            File.Contents(PDFファイルパス),
            [
                Implementation = "1.3",
                MultiPageTables = false
            ]
        ),

    // Pageだけを抽出
    Pageのみ =
        Table.SelectRows(
            ソース,
            each [Kind] = "Page"
        ),

    // ページ順に並べる
    Page順 =
        Table.Sort(
            Pageのみ,
            {
                {"Id", Order.Ascending}
            }
        ),

    // 各Pageに連番を付ける
    Page番号追加 =
        Table.AddIndexColumn(
            Page順,
            "PageNo",
            1,
            1,
            Int64.Type
        ),

    // 各ページの列名にPage番号を付ける
    列名変更 =
        Table.AddColumn(
            Page番号追加,
            "加工後Data",
            each
                let
                    現在Page番号 =
                        Text.PadStart(
                            Text.From([PageNo]),
                            3,
                            "0"
                        ),

                    元Data =
                        [Data],

                    元列名 =
                        Table.ColumnNames(元Data),

                    新列名 =
                        List.Transform(
                            元列名,
                            each
                                "Page"
                                & 現在Page番号
                                & "_"
                                & _
                        ),

                    列名対応 =
                        List.Zip(
                            {
                                元列名,
                                新列名
                            }
                        ),

                    名前変更後 =
                        Table.RenameColumns(
                            元Data,
                            列名対応,
                            MissingField.Ignore
                        )
                in
                    名前変更後
        ),

    // 各Pageのテーブルをリストとして取得
    Pageテーブル一覧 =
        列名変更[加工後Data],

    // 最大行数を取得
    最大行数 =
        if List.Count(Pageテーブル一覧) = 0 then
            0
        else
            List.Max(
                List.Transform(
                    Pageテーブル一覧,
                    each Table.RowCount(_)
                )
            ),

    // 各Pageに行番号を付ける
    行番号付き =
        List.Transform(
            Pageテーブル一覧,
            each
                Table.AddIndexColumn(
                    _,
                    "結合用行番号",
                    1,
                    1,
                    Int64.Type
                )
        ),

    // 先頭Pageを基準に横結合
    横結合 =
        if List.Count(行番号付き) = 0 then
            #table({}, {})
        else
            List.Accumulate(
                List.Skip(行番号付き, 1),
                List.First(行番号付き),
                (結合済み, 次Page) =>
                    Table.Join(
                        結合済み,
                        "結合用行番号",
                        次Page,
                        "結合用行番号",
                        JoinKind.FullOuter
                    )
            ),

    // 行番号で並べる
    行順整理 =
        Table.Sort(
            横結合,
            {
                {"結合用行番号", Order.Ascending}
            }
        ),

    // 結合用行番号を削除
    行番号削除 =
        Table.RemoveColumns(
            行順整理,
            {"結合用行番号"}
        )
in
    行番号削除








let
    //----------------------------------------
    // PDFファイル
    //----------------------------------------
    PDFファイルパス =
        "C:\資料\対象ファイル.pdf",

    //----------------------------------------
    // PDFを読み込む
    //----------------------------------------
    ソース =
        Pdf.Tables(
            File.Contents(PDFファイルパス),
            [
                Implementation = "1.3",
                MultiPageTables = false
            ]
        ),

    //----------------------------------------
    // Pageだけを抽出
    //----------------------------------------
    Pageのみ =
        Table.SelectRows(
            ソース,
            each [Kind] = "Page"
        ),

    //----------------------------------------
    // Page名で並べる
    //----------------------------------------
    Page順 =
        Table.Sort(
            Pageのみ,
            {
                {"Id", Order.Ascending}
            }
        ),

    //----------------------------------------
    // ページ番号を付ける
    //----------------------------------------
    Page番号追加 =
        Table.AddIndexColumn(
            Page順,
            "PageNo",
            1,
            1,
            Int64.Type
        ),

    //----------------------------------------
    // Pageごとに列データと列名を取得
    //----------------------------------------
    Page列情報 =
        List.Transform(
            Table.ToRecords(Page番号追加),

            (現在Page as record) =>
                let
                    Page番号 =
                        Text.PadStart(
                            Text.From(現在Page[PageNo]),
                            3,
                            "0"
                        ),

                    元テーブル =
                        現在Page[Data],

                    元列名 =
                        Table.ColumnNames(元テーブル),

                    // 各列をリストとして取得
                    列データ =
                        Table.ToColumns(元テーブル),

                    // 列名が重複しないようPage番号を付加
                    新列名 =
                        List.Transform(
                            元列名,
                            each
                                "Page"
                                & Page番号
                                & "_"
                                & _
                        )
                in
                    [
                        Columns = 列データ,
                        Names = 新列名
                    ]
        ),

    //----------------------------------------
    // 全ページの列データをまとめる
    //----------------------------------------
    全列データ =
        List.Combine(
            List.Transform(
                Page列情報,
                each [Columns]
            )
        ),

    //----------------------------------------
    // 全ページの列名をまとめる
    //----------------------------------------
    全列名 =
        List.Combine(
            List.Transform(
                Page列情報,
                each [Names]
            )
        ),

    //----------------------------------------
    // 横方向に結合
    //----------------------------------------
    横結合 =
        if List.Count(全列データ) = 0 then
            #table({}, {})
        else
            Table.FromColumns(
                全列データ,
                全列名
            )
in
    横結合
