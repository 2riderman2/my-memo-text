Option Explicit

'============================================================
' 設定
'============================================================

'------------------------------------------------------------
' 個人メール
'------------------------------------------------------------

' 受信トレイ直下の一時保存フォルダ
Private Const TEMP_FOLDER_NAME As String = "一時保存"

' Documents配下の保存先
Private Const PERSONAL_RECEIVE_FOLDER As String = "受信メール"
Private Const PERSONAL_SENT_FOLDER As String = "送信メール"

' 受信メール保存後に付ける分類
Private Const CATEGORY_NAME As String = "保存済み"


'------------------------------------------------------------
' 係共有メールボックス
'------------------------------------------------------------

' ★ここは実際にOutlook左側に表示されている
'   共有メールボックス名に変更してください
Private Const SHARED_STORE_NAME As String = "係共有メールボックス"

' ★共有メールボックスのメールアドレス
'
' 例：
' accounting@example.co.jp
'
' 共有アドレスから送ったメールを
' 個人の「送信メール」へ保存しないための判定にも使用します。
Private Const SHARED_SMTP_ADDRESS As String = ""

' 受信トレイ・送信済みアイテム直下の対象フォルダ
Private Const ACCOUNTING_FOLDER_NAME As String = "会計係"

' Documents配下の保存先
Private Const SHARED_RECEIVE_FOLDER As String = "係受信メール"
Private Const SHARED_SENT_FOLDER As String = "係送信メール"


'------------------------------------------------------------
' 共通
'------------------------------------------------------------

' ファイル名本体の最大文字数
Private Const MAX_FILE_NAME_LENGTH As Long = 180

' 共有メール監視種別
Private Const MODE_SHARED_RECEIVE As Long = 1
Private Const MODE_SHARED_SENT As Long = 2


'============================================================
' イベント監視用
'============================================================

' 個人受信メール
Private WithEvents TempItems As Outlook.Items

' 共有メールボックスの各フォルダ監視
Private mSharedWatchers As Collection

' 二重登録防止
Private mSharedWatcherKeys As Object


'============================================================
' Outlook起動時
'============================================================

Private Sub Application_Startup()

    InitializeAllMailSaving

End Sub


'============================================================
' 全機能初期化
'============================================================

Private Sub InitializeAllMailSaving()

    ' 個人受信
    InitializePersonalReceive

    ' 共有メール
    InitializeSharedMailbox

End Sub


'============================================================
' 手動再初期化用
'
' 必要な場合、VBE上でこのプロシージャ内にカーソルを置き
' F5で実行できます。
'============================================================

Public Sub ReinitializeMailSaving()

    InitializeAllMailSaving

    MsgBox _
        "メール自動保存機能を再初期化しました。", _
        vbInformation, _
        "メール自動保存"

End Sub


'************************************************************
' 個人受信メール
'************************************************************

'============================================================
' 個人受信メールの初期化
'============================================================

Private Sub InitializePersonalReceive()

    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim tempFolder As Outlook.Folder

    On Error GoTo ErrHandler

    Set ns = Application.Session

    Set inbox = _
        ns.GetDefaultFolder(olFolderInbox)

    Set tempFolder = _
        GetSubFolder( _
            inbox, _
            TEMP_FOLDER_NAME _
        )

    If tempFolder Is Nothing Then

        MsgBox _
            "個人メールボックスの受信トレイ直下に「" & _
            TEMP_FOLDER_NAME & _
            "」フォルダが見つかりません。", _
            vbExclamation, _
            "メール自動保存"

        Exit Sub

    End If

    ' リアルタイム監視
    Set TempItems = tempFolder.Items

    ' 分類を準備
    EnsureCategoryExists CATEGORY_NAME

    ' Outlook停止中に入った分を処理
    ProcessExistingPersonalReceivedItems tempFolder

    Exit Sub


ErrHandler:

    Debug.Print _
        "InitializePersonalReceive Error " & _
        Err.Number & " : " & Err.Description

End Sub


'============================================================
' 一時保存フォルダへの追加
'============================================================

Private Sub TempItems_ItemAdd(ByVal Item As Object)

    On Error GoTo ErrHandler

    If Not TypeOf Item Is Outlook.MailItem Then
        Exit Sub
    End If

    ProcessPersonalReceivedMail Item

    Exit Sub


ErrHandler:

    Debug.Print _
        "TempItems_ItemAdd Error " & _
        Err.Number & " : " & Err.Description

End Sub


'============================================================
' 起動時の既存受信メール処理
'============================================================

Private Sub ProcessExistingPersonalReceivedItems( _
    ByVal tempFolder As Outlook.Folder _
)

    Dim i As Long
    Dim obj As Object

    On Error GoTo ErrHandler

    ' Moveで件数が変化するため後ろから処理
    For i = tempFolder.Items.Count To 1 Step -1

        Set obj = tempFolder.Items.Item(i)

        If TypeOf obj Is Outlook.MailItem Then

            ProcessPersonalReceivedMail obj

        End If

        Set obj = Nothing

    Next i

    Exit Sub


ErrHandler:

    Debug.Print _
        "ProcessExistingPersonalReceivedItems Error " & _
        Err.Number & " : " & Err.Description

End Sub


'============================================================
' 個人受信メール1件を処理
'============================================================

Private Sub ProcessPersonalReceivedMail( _
    ByVal mail As Outlook.MailItem _
)

    Dim saveFolder As String
    Dim fileName As String
    Dim inbox As Outlook.Folder

    On Error GoTo ErrHandler

    saveFolder = _
        GetDocumentsSubFolder( _
            PERSONAL_RECEIVE_FOLDER _
        )

    fileName = _
        BuildReceivedFileName(mail)

    ' 保存成功または同名ファイル既存の場合のみ
    ' 分類付与・受信トレイへ戻す
    If SaveMailAsMsg( _
        mail, _
        saveFolder, _
        fileName _
    ) Then

        AddCategory _
            mail, _
            CATEGORY_NAME

        mail.Save

        Set inbox = _
            Application.Session.GetDefaultFolder( _
                olFolderInbox _
            )

        mail.Move inbox

    End If

    Exit Sub


ErrHandler:

    Debug.Print _
        "ProcessPersonalReceivedMail Error " & _
        Err.Number & " : " & Err.Description

End Sub


'************************************************************
' 個人送信メール
'************************************************************

'============================================================
' メール送信時
'============================================================

Private Sub Application_ItemSend( _
    ByVal Item As Object, _
    Cancel As Boolean _
)

    Dim mail As Outlook.MailItem

    On Error GoTo ErrHandler

    If Not TypeOf Item Is Outlook.MailItem Then
        Exit Sub
    End If

    Set mail = Item

    '--------------------------------------------------------
    ' 係共有メールボックスから送信したメールは
    ' 個人の「送信メール」には保存しない
    '--------------------------------------------------------
    If IsSharedMailboxSend(mail) Then
        Exit Sub
    End If

    SavePersonalSentMail mail

    Exit Sub


ErrHandler:

    ' 保存エラーでもメール送信は止めない
    Debug.Print _
        "Application_ItemSend Error " & _
        Err.Number & " : " & Err.Description

End Sub


'============================================================
' 個人送信メール保存
'============================================================

Private Sub SavePersonalSentMail( _
    ByVal mail As Outlook.MailItem _
)

    Dim saveFolder As String
    Dim fileName As String

    On Error GoTo ErrHandler

    saveFolder = _
        GetDocumentsSubFolder( _
            PERSONAL_SENT_FOLDER _
        )

    fileName = _
        BuildPersonalSentFileName(mail)

    Call SaveMailAsMsg( _
        mail, _
        saveFolder, _
        fileName _
    )

    Exit Sub


ErrHandler:

    Debug.Print _
        "SavePersonalSentMail Error " & _
        Err.Number & " : " & Err.Description

End Sub


'============================================================
' 共有メールボックスからの送信か判定
'============================================================

Private Function IsSharedMailboxSend( _
    ByVal mail As Outlook.MailItem _
) As Boolean

    Dim account As Outlook.Account
    Dim behalfName As String
    Dim smtp As String
    Dim accountName As String

    On Error Resume Next

    '--------------------------------------------------------
    ' 代理送信名
    '--------------------------------------------------------
    behalfName = _
        Trim(mail.SentOnBehalfOfName)

    If behalfName <> "" Then

        If StrComp( _
            behalfName, _
            SHARED_STORE_NAME, _
            vbTextCompare _
        ) = 0 Then

            IsSharedMailboxSend = True
            Exit Function

        End If

        If SHARED_SMTP_ADDRESS <> "" Then

            If StrComp( _
                behalfName, _
                SHARED_SMTP_ADDRESS, _
                vbTextCompare _
            ) = 0 Then

                IsSharedMailboxSend = True
                Exit Function

            End If

        End If

    End If

    '--------------------------------------------------------
    ' SendUsingAccount
    '--------------------------------------------------------
    Set account = mail.SendUsingAccount

    If Not account Is Nothing Then

        smtp = Trim(account.SmtpAddress)
        accountName = Trim(account.DisplayName)

        If SHARED_SMTP_ADDRESS <> "" Then

            If StrComp( _
                smtp, _
                SHARED_SMTP_ADDRESS, _
                vbTextCompare _
            ) = 0 Then

                IsSharedMailboxSend = True
                Exit Function

            End If

        End If

        If StrComp( _
            accountName, _
            SHARED_STORE_NAME, _
            vbTextCompare _
        ) = 0 Then

            IsSharedMailboxSend = True
            Exit Function

        End If

    End If

    On Error GoTo 0

    IsSharedMailboxSend = False

End Function


'************************************************************
' 係共有メールボックス
'************************************************************

'============================================================
' 係共有メールボックス初期化
'============================================================

Private Sub InitializeSharedMailbox()

    Dim sharedStore As Outlook.Store

    Dim sharedInbox As Outlook.Folder
    Dim sharedSent As Outlook.Folder

    Dim accountingReceive As Outlook.Folder
    Dim accountingSent As Outlook.Folder

    Dim errorMessage As String

    On Error GoTo ErrHandler

    ' 監視オブジェクトを作り直す
    Set mSharedWatchers = New Collection

    Set mSharedWatcherKeys = _
        CreateObject("Scripting.Dictionary")

    '--------------------------------------------------------
    ' 共有メールボックス取得
    '--------------------------------------------------------
    Set sharedStore = _
        GetSharedStore()

    If sharedStore Is Nothing Then

        MsgBox _
            "係共有メールボックス「" & _
            SHARED_STORE_NAME & _
            "」が見つかりません。" & vbCrLf & vbCrLf & _
            "コード先頭の SHARED_STORE_NAME を、" & _
            "Outlook左側に表示されている名前と一致させてください。", _
            vbExclamation, _
            "共有メール自動保存"

        Exit Sub

    End If

    '--------------------------------------------------------
    ' 共有メールボックスの既定フォルダ取得
    '--------------------------------------------------------
    Set sharedInbox = _
        sharedStore.GetDefaultFolder( _
            olFolderInbox _
        )

    Set sharedSent = _
        sharedStore.GetDefaultFolder( _
            olFolderSentMail _
        )

    '--------------------------------------------------------
    ' 受信トレイ\会計係
    '--------------------------------------------------------
    If Not sharedInbox Is Nothing Then

        Set accountingReceive = _
            GetSubFolder( _
                sharedInbox, _
                ACCOUNTING_FOLDER_NAME _
            )

    End If

    '--------------------------------------------------------
    ' 送信済みアイテム\会計係
    '--------------------------------------------------------
    If Not sharedSent Is Nothing Then

        Set accountingSent = _
            GetSubFolder( _
                sharedSent, _
                ACCOUNTING_FOLDER_NAME _
            )

    End If

    '--------------------------------------------------------
    ' 受信側
    '--------------------------------------------------------
    If accountingReceive Is Nothing Then

        errorMessage = _
            errorMessage & _
            "・受信トレイ直下の「" & _
            ACCOUNTING_FOLDER_NAME & _
            "」が見つかりません。" & vbCrLf

    Else

        ' 会計係自身＋全サブフォルダを監視
        ' 既存メールも処理
        RegisterSharedFolderTree _
            accountingReceive, _
            MODE_SHARED_RECEIVE, _
            True

    End If

    '--------------------------------------------------------
    ' 送信側
    '--------------------------------------------------------
    If accountingSent Is Nothing Then

        errorMessage = _
            errorMessage & _
            "・送信済みアイテム直下の「" & _
            ACCOUNTING_FOLDER_NAME & _
            "」が見つかりません。" & vbCrLf

    Else

        RegisterSharedFolderTree _
            accountingSent, _
            MODE_SHARED_SENT, _
            True

    End If

    If errorMessage <> "" Then

        MsgBox _
            "共有メールボックスの一部を初期化できませんでした。" & _
            vbCrLf & vbCrLf & _
            errorMessage, _
            vbExclamation, _
            "共有メール自動保存"

    End If

    Exit Sub


ErrHandler:

    MsgBox _
        "係共有メールボックスの初期化中にエラーが発生しました。" & _
        vbCrLf & vbCrLf & _
        "エラー番号：" & Err.Number & vbCrLf & _
        "内容：" & Err.Description, _
        vbExclamation, _
        "共有メール自動保存"

End Sub


'============================================================
' 共有フォルダを再帰的に監視登録
'
' クラスモジュールからも呼ぶためPublic
'============================================================

Public Sub RegisterSharedFolderTree( _
    ByVal rootFolder As Outlook.Folder, _
    ByVal mode As Long, _
    Optional ByVal processExisting As Boolean = False _
)

    Dim watcher As CSharedFolderWatcher
    Dim subFolder As Outlook.Folder

    Dim watcherKey As String

    On Error GoTo ErrHandler

    If mSharedWatchers Is Nothing Then
        Set mSharedWatchers = New Collection
    End If

    If mSharedWatcherKeys Is Nothing Then

        Set mSharedWatcherKeys = _
            CreateObject("Scripting.Dictionary")

    End If

    '--------------------------------------------------------
    ' フォルダごとの一意キー
    '--------------------------------------------------------
    watcherKey = _
        CStr(mode) & "|" & _
        rootFolder.EntryID

    '--------------------------------------------------------
    ' まだ監視していなければ登録
    '--------------------------------------------------------
    If Not mSharedWatcherKeys.Exists(watcherKey) Then

        Set watcher = _
            New CSharedFolderWatcher

        watcher.Initialize _
            rootFolder, _
            mode, _
            Me

        mSharedWatchers.Add watcher

        mSharedWatcherKeys.Add _
            watcherKey, _
            True

    End If

    '--------------------------------------------------------
    ' 現在フォルダ内の既存メールを処理
    '--------------------------------------------------------
    If processExisting Then

        ProcessExistingSharedItems _
            rootFolder, _
            mode

    End If

    '--------------------------------------------------------
    ' 下位フォルダへ再帰
    '--------------------------------------------------------
    For Each subFolder In rootFolder.Folders

        RegisterSharedFolderTree _
            subFolder, _
            mode, _
            processExisting

    Next subFolder

    Exit Sub


ErrHandler:

    Debug.Print _
        "RegisterSharedFolderTree Error " & _
        Err.Number & " : " & Err.Description

End Sub


'============================================================
' 共有フォルダにメールが追加されたとき
'
' クラスモジュールから呼ぶためPublic
'============================================================

Public Sub HandleSharedFolderItem( _
    ByVal Item As Object, _
    ByVal mode As Long _
)

    On Error GoTo ErrHandler

    If Not TypeOf Item Is Outlook.MailItem Then
        Exit Sub
    End If

    Select Case mode

        Case MODE_SHARED_RECEIVE

            SaveSharedReceivedMail Item

        Case MODE_SHARED_SENT

            SaveSharedSentMail Item

    End Select

    Exit Sub


ErrHandler:

    Debug.Print _
        "HandleSharedFolderItem Error " & _
        Err.Number & " : " & Err.Description

End Sub


'============================================================
' 共有フォルダ内の既存アイテム処理
'============================================================

Private Sub ProcessExistingSharedItems( _
    ByVal targetFolder As Outlook.Folder, _
    ByVal mode As Long _
)

    Dim i As Long
    Dim obj As Object

    On Error GoTo ErrHandler

    For i = 1 To targetFolder.Items.Count

        Set obj = _
            targetFolder.Items.Item(i)

        If TypeOf obj Is Outlook.MailItem Then

            HandleSharedFolderItem _
                obj, _
                mode

        End If

        Set obj = Nothing

    Next i

    Exit Sub


ErrHandler:

    Debug.Print _
        "ProcessExistingSharedItems Error " & _
        Err.Number & " : " & Err.Description

End Sub


'============================================================
' 係共有・受信メール保存
'============================================================

Private Sub SaveSharedReceivedMail( _
    ByVal mail As Outlook.MailItem _
)

    Dim saveFolder As String
    Dim fileName As String

    On Error GoTo ErrHandler

    saveFolder = _
        GetDocumentsSubFolder( _
            SHARED_RECEIVE_FOLDER _
        )

    fileName = _
        BuildReceivedFileName(mail)

    Call SaveMailAsMsg( _
        mail, _
        saveFolder, _
        fileName _
    )

    ' 分類付与なし
    ' メール移動なし

    Exit Sub


ErrHandler:

    Debug.Print _
        "SaveSharedReceivedMail Error " & _
        Err.Number & " : " & Err.Description

End Sub


'============================================================
' 係共有・送信メール保存
'============================================================

Private Sub SaveSharedSentMail( _
    ByVal mail As Outlook.MailItem _
)

    Dim saveFolder As String
    Dim fileName As String

    On Error GoTo ErrHandler

    saveFolder = _
        GetDocumentsSubFolder( _
            SHARED_SENT_FOLDER _
        )

    fileName = _
        BuildStoredSentFileName(mail)

    Call SaveMailAsMsg( _
        mail, _
        saveFolder, _
        fileName _
    )

    ' 分類付与なし
    ' メール移動なし

    Exit Sub


ErrHandler:

    Debug.Print _
        "SaveSharedSentMail Error " & _
        Err.Number & " : " & Err.Description

End Sub


'************************************************************
' ファイル名作成
'************************************************************

'============================================================
' 受信メール
'
' yyyyMMdd_HHmm_差出人_件名.msg
'============================================================

Private Function BuildReceivedFileName( _
    ByVal mail As Outlook.MailItem _
) As String

    Dim dateText As String
    Dim senderText As String
    Dim subjectText As String
    Dim result As String
    Dim receivedDate As Date

    On Error GoTo ErrHandler

    receivedDate = _
        GetReceivedDate(mail)

    dateText = _
        Format( _
            receivedDate, _
            "yyyymmdd_HHmm" _
        )

    '--------------------------------------------------------
    ' 差出人
    '--------------------------------------------------------
    senderText = _
        Trim(mail.SenderName)

    If senderText = "" Then
        senderText = "差出人不明"
    End If

    senderText = _
        RemoveParenthesisPart(senderText)

    senderText = _
        CleanFileName(senderText)

    If senderText = "" Then
        senderText = "差出人不明"
    End If

    '--------------------------------------------------------
    ' 件名
    '--------------------------------------------------------
    subjectText = _
        CleanSubject(mail.Subject)

    '--------------------------------------------------------
    ' 組み立て
    '--------------------------------------------------------
    result = _
        dateText & "_" & _
        senderText & "_" & _
        subjectText

    result = _
        LimitFileNameLength(result)

    BuildReceivedFileName = _
        result & ".msg"

    Exit Function


ErrHandler:

    BuildReceivedFileName = _
        Format(Now, "yyyymmdd_HHmm") & _
        "_受信メール.msg"

End Function


'============================================================
' 個人送信メール
'
' ItemSend時点なので現在時刻
'============================================================

Private Function BuildPersonalSentFileName( _
    ByVal mail As Outlook.MailItem _
) As String

    Dim dateText As String
    Dim recipientText As String
    Dim subjectText As String
    Dim result As String

    On Error GoTo ErrHandler

    dateText = _
        Format( _
            Now, _
            "yyyymmdd_HHmm" _
        )

    recipientText = _
        GetRecipientFileName(mail)

    subjectText = _
        CleanSubject(mail.Subject)

    result = _
        dateText & "_" & _
        recipientText & "_" & _
        subjectText

    result = _
        LimitFileNameLength(result)

    BuildPersonalSentFileName = _
        result & ".msg"

    Exit Function


ErrHandler:

    BuildPersonalSentFileName = _
        Format(Now, "yyyymmdd_HHmm") & _
        "_送信メール.msg"

End Function


'============================================================
' 係共有送信メール
'
' 実際のSentOnを使用
'============================================================

Private Function BuildStoredSentFileName( _
    ByVal mail As Outlook.MailItem _
) As String

    Dim dateText As String
    Dim recipientText As String
    Dim subjectText As String
    Dim result As String
    Dim sentDate As Date

    On Error GoTo ErrHandler

    sentDate = _
        GetSentDate(mail)

    dateText = _
        Format( _
            sentDate, _
            "yyyymmdd_HHmm" _
        )

    recipientText = _
        GetRecipientFileName(mail)

    subjectText = _
        CleanSubject(mail.Subject)

    result = _
        dateText & "_" & _
        recipientText & "_" & _
        subjectText

    result = _
        LimitFileNameLength(result)

    BuildStoredSentFileName = _
        result & ".msg"

    Exit Function


ErrHandler:

    BuildStoredSentFileName = _
        Format(Now, "yyyymmdd_HHmm") & _
        "_係送信メール.msg"

End Function


'============================================================
' 宛先名
'
' Toが1名：
' 山田太郎
'
' Toが複数：
' 山田太郎他
'
' CC・BCCは人数に含めない
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

    For Each recipient In mail.Recipients

        If recipient.Type = olTo Then

            toCount = _
                toCount + 1

            If toCount = 1 Then

                firstName = _
                    recipient.Name

            End If

        End If

    Next recipient

    If Trim(firstName) = "" Then

        GetRecipientFileName = _
            "宛先なし"

        Exit Function

    End If

    ' ( / （ 以降を削除
    firstName = _
        RemoveParenthesisPart(firstName)

    firstName = _
        CleanFileName(firstName)

    If Trim(firstName) = "" Then
        firstName = "宛先不明"
    End If

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
' 受信日時
'============================================================

Private Function GetReceivedDate( _
    ByVal mail As Outlook.MailItem _
) As Date

    Dim d As Date

    On Error Resume Next

    d = mail.ReceivedTime

    On Error GoTo 0

    If d = 0 Then
        d = Now
    End If

    GetReceivedDate = d

End Function


'============================================================
' 送信日時
'============================================================

Private Function GetSentDate( _
    ByVal mail As Outlook.MailItem _
) As Date

    Dim d As Date

    On Error Resume Next

    d = mail.SentOn

    On Error GoTo 0

    If d = 0 Then
        d = Now
    End If

    GetSentDate = d

End Function


'============================================================
' 件名の整形
'============================================================

Private Function CleanSubject( _
    ByVal subjectText As String _
) As String

    subjectText = _
        Trim(subjectText)

    If subjectText = "" Then
        subjectText = "件名なし"
    End If

    subjectText = _
        CleanFileName(subjectText)

    If subjectText = "" Then
        subjectText = "件名なし"
    End If

    CleanSubject = subjectText

End Function


'============================================================
' 「(」または「（」以降を削除
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

    If posHalf > 0 And posFull > 0 Then

        If posHalf < posFull Then
            cutPos = posHalf
        Else
            cutPos = posFull
        End If

    ElseIf posHalf > 0 Then

        cutPos = posHalf

    ElseIf posFull > 0 Then

        cutPos = posFull

    End If

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
' Windowsファイル名禁止文字処理
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

    For Each ch In invalidChars

        text = _
            Replace( _
                text, _
                CStr(ch), _
                "_" _
            )

    Next ch

    text = _
        Replace(text, vbCr, " ")

    text = _
        Replace(text, vbLf, " ")

    text = _
        Replace(text, vbTab, " ")

    text = Trim(text)

    ' 末尾の空白・ピリオドを除去
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
' 長すぎるファイル名を短縮
'============================================================

Private Function LimitFileNameLength( _
    ByVal text As String _
) As String

    If Len(text) > MAX_FILE_NAME_LENGTH Then

        text = _
            Left( _
                text, _
                MAX_FILE_NAME_LENGTH _
            )

    End If

    LimitFileNameLength = text

End Function


'************************************************************
' ファイル保存
'************************************************************

'============================================================
' MSG保存
'
' 成功：True
' 同名既存：True
' エラー：False
'============================================================

Private Function SaveMailAsMsg( _
    ByVal mail As Outlook.MailItem, _
    ByVal saveFolder As String, _
    ByVal fileName As String _
) As Boolean

    Dim fullPath As String

    On Error GoTo ErrHandler

    EnsureFolderExists saveFolder

    fullPath = _
        saveFolder & "\" & fileName

    '--------------------------------------------------------
    ' 同名ファイルが存在する場合は保存しない
    '--------------------------------------------------------
    If Dir(fullPath) <> "" Then

        Debug.Print _
            "重複のため保存スキップ：" & _
            fullPath

        SaveMailAsMsg = True

        Exit Function

    End If

    mail.SaveAs _
        fullPath, _
        olMSG

    SaveMailAsMsg = True

    Exit Function


ErrHandler:

    Debug.Print _
        "SaveMailAsMsg Error " & _
        Err.Number & " : " & _
        Err.Description

    SaveMailAsMsg = False

End Function


'============================================================
' Documents配下のフォルダパス
'============================================================

Private Function GetDocumentsSubFolder( _
    ByVal subFolderName As String _
) As String

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

    GetDocumentsSubFolder = _
        documentsPath & "\" & _
        subFolderName

    Set shell = Nothing

End Function


'============================================================
' Windowsフォルダ作成
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


'************************************************************
' Outlookフォルダ関連
'************************************************************

'============================================================
' サブフォルダ取得
'============================================================

Private Function GetSubFolder( _
    ByVal parentFolder As Outlook.Folder, _
    ByVal folderName As String _
) As Outlook.Folder

    Dim targetFolder As Outlook.Folder

    On Error Resume Next

    Set targetFolder = _
        parentFolder.Folders(folderName)

    On Error GoTo 0

    Set GetSubFolder = _
        targetFolder

End Function


'============================================================
' 係共有メールボックスのStoreを取得
'============================================================

Private Function GetSharedStore() As Outlook.Store

    Dim store As Outlook.Store
    Dim rootFolder As Outlook.Folder

    On Error GoTo ErrHandler

    For Each store In Application.Session.Stores

        '----------------------------------------------------
        ' Storeの表示名
        '----------------------------------------------------
        If StrComp( _
            Trim(store.DisplayName), _
            Trim(SHARED_STORE_NAME), _
            vbTextCompare _
        ) = 0 Then

            Set GetSharedStore = store
            Exit Function

        End If

        '----------------------------------------------------
        ' ルートフォルダ名も確認
        '----------------------------------------------------
        Set rootFolder = Nothing

        On Error Resume Next

        Set rootFolder = _
            store.GetRootFolder

        On Error GoTo ErrHandler

        If Not rootFolder Is Nothing Then

            If StrComp( _
                Trim(rootFolder.Name), _
                Trim(SHARED_STORE_NAME), _
                vbTextCompare _
            ) = 0 Then

                Set GetSharedStore = store
                Exit Function

            End If

        End If

    Next store

    Set GetSharedStore = Nothing

    Exit Function


ErrHandler:

    Set GetSharedStore = Nothing

End Function


'************************************************************
' 分類項目
'************************************************************

'============================================================
' 分類が無ければ作成
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
' 分類を追加
'============================================================

Private Sub AddCategory( _
    ByVal mail As Outlook.MailItem, _
    ByVal categoryName As String _
)

    Dim existingCategories As String

    On Error GoTo ErrHandler

    existingCategories = _
        Trim(mail.Categories)

    If HasCategory( _
        existingCategories, _
        categoryName _
    ) Then

        Exit Sub

    End If

    If existingCategories = "" Then

        mail.Categories = _
            categoryName

    Else

        mail.Categories = _
            existingCategories & _
            "," & _
            categoryName

    End If

    Exit Sub


ErrHandler:

    Debug.Print _
        "AddCategory Error " & _
        Err.Number & " : " & Err.Description

End Sub


'============================================================
' 分類が付いているか確認
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
