Option Explicit

'============================================================
' 設定
'============================================================

' 受信トレイ直下の一時保存フォルダ名
Private Const TEMP_FOLDER_NAME As String = "一時保存"

' ドキュメントフォルダ内の保存先フォルダ名
Private Const SAVE_FOLDER_NAME As String = "受信メール"

' ファイル名に使用する日時
'
' 例：
' 20260807_172530_山田太郎_資料送付.msg
'
' VBAでは「分」は nn
Private Const DATE_FORMAT As String = "yyyymmdd_hhnnss"

' パスが長くなりすぎることを防ぐための上限
Private Const MAX_PATH_LENGTH As Long = 240

'============================================================
' イベント監視用
'============================================================

Private WithEvents mTemporaryItems As Outlook.Items

' 二重処理防止
Private mIsProcessing As Boolean


'============================================================
' Outlook起動時
'============================================================

Private Sub Application_Startup()

    InitializeMailSaver

End Sub


'============================================================
' 初期化
'
' Outlook起動時に自動実行される。
'
' Alt + F8
' → InitializeMailSaver
'
' から手動実行することも可能。
'============================================================

Public Sub InitializeMailSaver()

    Dim tempFolder As Outlook.Folder

    On Error GoTo ErrorHandler

    ' 一時保存フォルダを取得
    Set tempFolder = GetTemporaryFolder()

    ' フォルダへのメール追加を監視
    Set mTemporaryItems = tempFolder.Items

    ' 保存先フォルダがなければ作成
    EnsureFolderExists GetSaveFolderPath()

    ' Outlook停止中などに一時保存フォルダへ
    ' 残っていたメールも処理
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
' 一時保存フォルダ内のメールをすべて処理
'
' 手動実行も可能
'
' Alt + F8
' → ProcessTemporaryFolderNow
'============================================================

Public Sub ProcessTemporaryFolderNow()

    Dim tempFolder As Outlook.Folder
    Dim currentItem As Object
    Dim itemIndex As Long

    Dim errorNumber As Long
    Dim errorDescription As String

    ' すでに処理中なら実行しない
    If mIsProcessing Then Exit Sub

    mIsProcessing = True

    On Error GoTo ErrorHandler

    Set tempFolder = GetTemporaryFolder()

    ' 保存先がなければ作成
    EnsureFolderExists GetSaveFolderPath()

    ' メールを処理中にフォルダから移動するので
    ' 下から上へ処理する
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
' メール1件を処理
'============================================================

Private Sub ProcessOneMail(ByVal mail As Outlook.MailItem)

    Dim saveFolderPath As String
    Dim savePath As String

    Dim senderText As String
    Dim subjectText As String
    Dim fileStem As String

    Dim inboxFolder As Outlook.Folder

    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo ErrorHandler


    '--------------------------------------------------------
    ' 保存先取得
    '--------------------------------------------------------

    saveFolderPath = GetSaveFolderPath()

    EnsureFolderExists saveFolderPath

    Set inboxFolder = GetInboxFolder()


    '--------------------------------------------------------
    ' 差出人名
    '--------------------------------------------------------

    senderText = CleanFileComponent( _
        mail.SenderName, _
        "差出人不明")

    ' 差出人名が極端に長い場合は短縮
    If Len(senderText) > 60 Then

        senderText = Left$(senderText, 60)

    End If


    '--------------------------------------------------------
    ' 件名
    '--------------------------------------------------------

    subjectText = CleanFileComponent( _
        mail.Subject, _
        "件名なし")


    '--------------------------------------------------------
    ' ファイル名作成
    '
    ' 例：
    '
    ' 20260807_172530_山田太郎_資料送付.msg
    '--------------------------------------------------------

    fileStem = _
        Format$(mail.ReceivedTime, DATE_FORMAT) & "_" & _
        senderText & "_" & _
        subjectText


    '--------------------------------------------------------
    ' パス全体が長くなりすぎないように調整
    '--------------------------------------------------------

    fileStem = FitFileStemToPath( _
        saveFolderPath, _
        fileStem, _
        ".msg")


    savePath = _
        saveFolderPath & "\" & fileStem & ".msg"


    '--------------------------------------------------------
    ' 同名ファイルがすでに存在する場合
    '
    ' ・MSG保存しない
    ' ・重複スキップログへ記録
    ' ・メールを受信トレイへ移動
    '--------------------------------------------------------

    If FileExists(savePath) Then

        WriteDuplicateLog _
            mail, _
            savePath

        mail.Move inboxFolder

        Exit Sub

    End If


    '--------------------------------------------------------
    ' MSG形式で保存
    '--------------------------------------------------------

    mail.SaveAs _
        savePath, _
        olMSGUnicode


    '--------------------------------------------------------
    ' MSG保存成功後
    '
    ' 受信トレイへ移動
    '--------------------------------------------------------

    mail.Move inboxFolder

    Exit Sub


ErrorHandler:

    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next

    WriteErrorLog _
        mail, _
        "メール保存処理", _
        errorNumber, _
        errorDescription, _
        savePath

    On Error GoTo 0

    ' エラーになったメールは受信トレイへ移動しない。
    ' 一時保存フォルダへ残す。

End Sub


'============================================================
' 一時保存フォルダ取得
'
' 想定：
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
            "受信トレイ直下に「" & _
            TEMP_FOLDER_NAME & _
            "」フォルダが見つかりません。"

    End If


    Set GetTemporaryFolder = tempFolder

End Function


'============================================================
' 受信トレイ取得
'============================================================

Private Function GetInboxFolder() As Outlook.Folder

    Set GetInboxFolder = _
        Application.Session.GetDefaultFolder(olFolderInbox)

End Function


'============================================================
' 保存先フォルダ取得
'
' Windowsが認識している実際の
' 「ドキュメント」フォルダを取得する。
'
' 会社PCの例：
'
' C:\Users\ユーザー名\
' OneDrive - 会社名\
' Documents\
' 受信メール
'
'============================================================

Private Function GetSaveFolderPath() As String

    Dim shellObject As Object
    Dim documentsPath As String

    Set shellObject = _
        CreateObject("WScript.Shell")

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
' 実際の保存先を画面表示
'
' Alt + F8
' → ShowSaveFolderPath
'
' で確認可能
'============================================================

Public Sub ShowSaveFolderPath()

    MsgBox _
        "メールの保存先は次のフォルダです。" & _
        vbCrLf & vbCrLf & _
        GetSaveFolderPath(), _
        vbInformation, _
        "MSG保存先"

End Sub


'============================================================
' 保存先フォルダをエクスプローラーで開く
'
' Alt + F8
' → OpenSaveFolder
'============================================================

Public Sub OpenSaveFolder()

    Dim saveFolderPath As String

    saveFolderPath = GetSaveFolderPath()

    EnsureFolderExists saveFolderPath

    Shell _
        "explorer.exe """ & saveFolderPath & """", _
        vbNormalFocus

End Sub


'============================================================
' フォルダが存在しない場合は作成
'============================================================

Private Sub EnsureFolderExists(ByVal folderPath As String)

    Dim fileSystem As Object
    Dim parentPath As String

    Set fileSystem = _
        CreateObject("Scripting.FileSystemObject")


    If fileSystem.FolderExists(folderPath) Then

        Exit Sub

    End If


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
' ファイルが存在するか確認
'============================================================

Private Function FileExists(ByVal filePath As String) As Boolean

    Dim fileSystem As Object

    If Len(filePath) = 0 Then

        FileExists = False

        Exit Function

    End If


    Set fileSystem = _
        CreateObject("Scripting.FileSystemObject")


    FileExists = _
        fileSystem.FileExists(filePath)

End Function


'============================================================
' ファイル名に使用できない文字を処理
'============================================================

Private Function CleanFileComponent( _
    ByVal sourceText As String, _
    ByVal fallbackText As String) As String

    Dim invalidCharacters As Variant
    Dim invalidCharacter As Variant

    Dim characterCode As Long
    Dim cleanedText As String


    cleanedText = sourceText


    '--------------------------------------------------------
    ' 制御文字を空白へ置換
    '--------------------------------------------------------

    For characterCode = 0 To 31

        cleanedText = _
            Replace( _
                cleanedText, _
                Chr$(characterCode), _
                " ")

    Next characterCode


    '--------------------------------------------------------
    ' Windowsのファイル名で使用できない文字
    '
    ' \ / : * ? " < > |
    '
    ' を「_」へ置換
    '--------------------------------------------------------

    invalidCharacters = _
        Array( _
            "\", _
            "/", _
            ":", _
            "*", _
            "?", _
            """", _
            "<", _
            ">", _
            "|")


    For Each invalidCharacter In invalidCharacters

        cleanedText = _
            Replace( _
                cleanedText, _
                CStr(invalidCharacter), _
                "_")

    Next invalidCharacter


    ' 前後の空白を削除
    cleanedText = Trim$(cleanedText)


    '--------------------------------------------------------
    ' 連続した空白を1文字へまとめる
    '--------------------------------------------------------

    Do While InStr(cleanedText, "  ") > 0

        cleanedText = _
            Replace(cleanedText, "  ", " ")

    Loop


    '--------------------------------------------------------
    ' ファイル名末尾の
    '
    ' ・空白
    ' ・ピリオド
    '
    ' を削除
    '--------------------------------------------------------

    Do While Len(cleanedText) > 0

        If Right$(cleanedText, 1) = "." _
            Or Right$(cleanedText, 1) = " " Then

            cleanedText = _
                Left$( _
                    cleanedText, _
                    Len(cleanedText) - 1)

        Else

            Exit Do

        End If

    Loop


    '--------------------------------------------------------
    ' 空文字になった場合
    '--------------------------------------------------------

    If Len(cleanedText) = 0 Then

        cleanedText = fallbackText

    End If


    CleanFileComponent = cleanedText

End Function


'============================================================
' パス全体が長くなりすぎないように調整
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
            Left$( _
                adjustedStem, _
                maximumStemLength)

    End If


    '--------------------------------------------------------
    ' 短縮した結果、
    ' 最後が空白またはピリオドなら削除
    '--------------------------------------------------------

    Do While Len(adjustedStem) > 0

        If Right$(adjustedStem, 1) = "." _
            Or Right$(adjustedStem, 1) = " " Then

            adjustedStem = _
                Left$( _
                    adjustedStem, _
                    Len(adjustedStem) - 1)

        Else

            Exit Do

        End If

    Loop


    FitFileStemToPath = adjustedStem

End Function


'============================================================
' エラーログ
'
' 保存場所：
'
' Documents
' └ 受信メール
'    └ 保存エラー.log
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

        mailSubject = _
            CleanLogText(mail.Subject)

        mailEntryID = _
            CleanLogText(mail.EntryID)

        senderName = _
            CleanLogText(mail.SenderName)

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
' 重複スキップログ
'
' 同じファイル名のMSGがすでに存在していた場合に記録。
'
' 保存場所：
'
' Documents
' └ 受信メール
'    └ 重複スキップ.log
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
' ログ用文字列から改行・タブを削除
'============================================================

Private Function CleanLogText( _
    ByVal sourceText As String) As String

    Dim cleanedText As String


    cleanedText = sourceText


    cleanedText = _
        Replace(cleanedText, vbCr, " ")

    cleanedText = _
        Replace(cleanedText, vbLf, " ")

    cleanedText = _
        Replace(cleanedText, vbTab, " ")


    CleanLogText = cleanedText

End Function
