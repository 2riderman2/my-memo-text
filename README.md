Option Explicit

'============================================================
' 個人の「送信済みアイテム」にある過去メールを一括保存
'
' 保存先：
'   Windowsのドキュメント\受信メール
'
' ファイル名：
'   yyyyMMdd_HHmm_最初の宛先_件名.msg
'
' Toが複数の場合：
'   yyyyMMdd_HHmm_最初の宛先他_件名.msg
'
' 名前の加工：
'   ・「.共有メールボックス」を削除
'   ・「(」または「（」以降を削除
'============================================================


'------------------------------------------------------------
' 設定
'------------------------------------------------------------

Private Const SAVE_FOLDER_NAME As String = "受信メール"

Private Const MAX_FILE_NAME_LENGTH As Long = 180

Private Const SHARED_MAILBOX_TEXT As String = _
    ".共有メールボックス"


'============================================================
' メイン処理
'
' Alt + F8 から
' SaveAllPastSentMails
' を実行
'============================================================

Public Sub SaveAllPastSentMails()

    Dim ns As Outlook.NameSpace
    Dim sentFolder As Outlook.Folder

    Dim obj As Object
    Dim mail As Outlook.MailItem

    Dim saveFolder As String
    Dim fileName As String
    Dim fullPath As String

    Dim totalCount As Long
    Dim savedCount As Long
    Dim skippedCount As Long
    Dim errorCount As Long

    Dim result As VbMsgBoxResult

    On Error GoTo FatalError

    '--------------------------------------------------------
    ' 実行確認
    '--------------------------------------------------------
    result = MsgBox( _
        "個人メールボックスの「送信済みアイテム」にある" & _
        "すべてのメールを保存します。" & vbCrLf & vbCrLf & _
        "保存先：" & vbCrLf & _
        "ドキュメント\" & SAVE_FOLDER_NAME & vbCrLf & vbCrLf & _
        "すでに同名ファイルがあるメールはスキップします。" & _
        vbCrLf & vbCrLf & _
        "実行しますか？", _
        vbYesNo + vbQuestion, _
        "過去の送信メール一括保存" _
    )

    If result <> vbYes Then
        Exit Sub
    End If

    '--------------------------------------------------------
    ' Outlook
    '--------------------------------------------------------
    Set ns = Application.Session

    ' 個人の既定「送信済みアイテム」
    Set sentFolder = _
        ns.GetDefaultFolder(olFolderSentMail)

    '--------------------------------------------------------
    ' 保存先
    '--------------------------------------------------------
    saveFolder = _
        GetPastSentSaveFolderPath()

    EnsurePastSentFolderExists _
        saveFolder

    '--------------------------------------------------------
    ' 件数初期化
    '--------------------------------------------------------
    totalCount = 0
    savedCount = 0
    skippedCount = 0
    errorCount = 0

    '--------------------------------------------------------
    ' 送信済みアイテムを全件処理
    '--------------------------------------------------------
    For Each obj In sentFolder.Items

        ' メールだけ対象
        If TypeOf obj Is Outlook.MailItem Then

            Set mail = obj

            totalCount = _
                totalCount + 1

            On Error GoTo ItemError

            '------------------------------------------------
            ' ファイル名作成
            '------------------------------------------------
            fileName = _
                BuildPastSentFileName(mail)

            fullPath = _
                saveFolder & "\" & fileName

            '------------------------------------------------
            ' すでに同名ファイルがある場合
            '------------------------------------------------
            If Dir(fullPath) <> "" Then

                skippedCount = _
                    skippedCount + 1

            Else

                '--------------------------------------------
                ' MSG形式で保存
                '--------------------------------------------
                mail.SaveAs _
                    fullPath, _
                    olMSG

                savedCount = _
                    savedCount + 1

            End If

NextItem:

            On Error GoTo FatalError

            Set mail = Nothing

        End If

    Next obj

    '--------------------------------------------------------
    ' 完了
    '--------------------------------------------------------
    MsgBox _
        "過去の送信メールの保存が完了しました。" & _
        vbCrLf & vbCrLf & _
        "対象メール：" & totalCount & "件" & vbCrLf & _
        "新規保存：" & savedCount & "件" & vbCrLf & _
        "同名のためスキップ：" & skippedCount & "件" & vbCrLf & _
        "エラー：" & errorCount & "件" & vbCrLf & vbCrLf & _
        "保存先：" & vbCrLf & _
        saveFolder, _
        vbInformation, _
        "過去の送信メール一括保存"

    Exit Sub


'------------------------------------------------------------
' 1メールだけでエラーになった場合
' 次のメールへ進む
'------------------------------------------------------------
ItemError:

    errorCount = _
        errorCount + 1

    Debug.Print _
        "SaveAllPastSentMails Item Error " & _
        Err.Number & " : " & _
        Err.Description

    Err.Clear

    Resume NextItem


'------------------------------------------------------------
' 全体エラー
'------------------------------------------------------------
FatalError:

    MsgBox _
        "処理中にエラーが発生しました。" & _
        vbCrLf & vbCrLf & _
        "エラー番号：" & Err.Number & vbCrLf & _
        "内容：" & Err.Description, _
        vbExclamation, _
        "過去の送信メール一括保存"

End Sub


'============================================================
' ファイル名作成
'
' yyyyMMdd_HHmm_宛先_件名.msg
'============================================================

Private Function BuildPastSentFileName( _
    ByVal mail As Outlook.MailItem _
) As String

    Dim dateText As String
    Dim recipientText As String
    Dim subjectText As String

    Dim result As String
    Dim sentDate As Date

    On Error GoTo ErrHandler

    '--------------------------------------------------------
    ' 送信日時
    '--------------------------------------------------------
    sentDate = _
        GetPastSentDate(mail)

    dateText = _
        Format( _
            sentDate, _
            "yyyymmdd_HHmm" _
        )

    '--------------------------------------------------------
    ' 宛先
    '--------------------------------------------------------
    recipientText = _
        GetPastSentRecipientName(mail)

    '--------------------------------------------------------
    ' 件名
    '--------------------------------------------------------
    subjectText = _
        Trim(mail.Subject)

    If subjectText = "" Then
        subjectText = "件名なし"
    End If

    subjectText = _
        CleanPastSentFileName(subjectText)

    If subjectText = "" Then
        subjectText = "件名なし"
    End If

    '--------------------------------------------------------
    ' ファイル名
    '--------------------------------------------------------
    result = _
        dateText & "_" & _
        recipientText & "_" & _
        subjectText

    '--------------------------------------------------------
    ' 長すぎる場合
    '--------------------------------------------------------
    If Len(result) > MAX_FILE_NAME_LENGTH Then

        result = _
            Left( _
                result, _
                MAX_FILE_NAME_LENGTH _
            )

    End If

    BuildPastSentFileName = _
        result & ".msg"

    Exit Function


ErrHandler:

    BuildPastSentFileName = _
        Format( _
            Now, _
            "yyyymmdd_HHmm" _
        ) & _
        "_送信メール.msg"

End Function


'============================================================
' ファイル名用の宛先名
'
' To 1名：
'   山田太郎
'
' To 複数：
'   山田太郎他
'
' CC・BCCは人数に含めない
'============================================================

Private Function GetPastSentRecipientName( _
    ByVal mail As Outlook.MailItem _
) As String

    Dim recipient As Outlook.Recipient

    Dim firstName As String
    Dim toCount As Long

    On Error GoTo ErrHandler

    firstName = ""
    toCount = 0

    '--------------------------------------------------------
    ' Toのみ確認
    '--------------------------------------------------------
    For Each recipient In mail.Recipients

        If recipient.Type = olTo Then

            toCount = _
                toCount + 1

            ' 最初の1名
            If toCount = 1 Then

                firstName = _
                    recipient.Name

            End If

        End If

    Next recipient

    '--------------------------------------------------------
    ' Toが取得できない場合
    '--------------------------------------------------------
    If Trim(firstName) = "" Then

        GetPastSentRecipientName = _
            "宛先なし"

        Exit Function

    End If

    '--------------------------------------------------------
    ' 「.共有メールボックス」を削除
    '--------------------------------------------------------
    firstName = _
        RemovePastSharedMailboxText( _
            firstName _
        )

    '--------------------------------------------------------
    ' 「(」「（」以降を削除
    '--------------------------------------------------------
    firstName = _
        RemovePastParenthesisPart( _
            firstName _
        )

    '--------------------------------------------------------
    ' Windowsファイル名用に整形
    '--------------------------------------------------------
    firstName = _
        CleanPastSentFileName( _
            firstName _
        )

    If Trim(firstName) = "" Then

        firstName = _
            "宛先不明"

    End If

    '--------------------------------------------------------
    ' Toが複数なら「他」
    '--------------------------------------------------------
    If toCount >= 2 Then

        GetPastSentRecipientName = _
            firstName & "他"

    Else

        GetPastSentRecipientName = _
            firstName

    End If

    Exit Function


ErrHandler:

    GetPastSentRecipientName = _
        "宛先不明"

End Function


'============================================================
' 「.共有メールボックス」を削除
'
' 例：
' 山田太郎.共有メールボックス
' ↓
' 山田太郎
'============================================================

Private Function RemovePastSharedMailboxText( _
    ByVal text As String _
) As String

    text = _
        Replace( _
            text, _
            SHARED_MAILBOX_TEXT, _
            "", _
            1, _
            -1, _
            vbTextCompare _
        )

    RemovePastSharedMailboxText = _
        Trim(text)

End Function


'============================================================
' 「(」または「（」以降を削除
'
' 例：
' 山田太郎(株式会社ABC)
' ↓
' 山田太郎
'============================================================

Private Function RemovePastParenthesisPart( _
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

    '--------------------------------------------------------
    ' 半角と全角の両方
    '--------------------------------------------------------
    If posHalf > 0 And posFull > 0 Then

        If posHalf < posFull Then

            cutPos = posHalf

        Else

            cutPos = posFull

        End If

    '--------------------------------------------------------
    ' 半角のみ
    '--------------------------------------------------------
    ElseIf posHalf > 0 Then

        cutPos = posHalf

    '--------------------------------------------------------
    ' 全角のみ
    '--------------------------------------------------------
    ElseIf posFull > 0 Then

        cutPos = posFull

    End If

    '--------------------------------------------------------
    ' 括弧以降を削除
    '--------------------------------------------------------
    If cutPos > 0 Then

        text = _
            Left( _
                text, _
                cutPos - 1 _
            )

    End If

    RemovePastParenthesisPart = _
        Trim(text)

End Function


'============================================================
' Windowsファイル名として使用できるように整形
'============================================================

Private Function CleanPastSentFileName( _
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

    '--------------------------------------------------------
    ' 禁止文字を「_」へ
    '--------------------------------------------------------
    For Each ch In invalidChars

        text = _
            Replace( _
                text, _
                CStr(ch), _
                "_" _
            )

    Next ch

    '--------------------------------------------------------
    ' 改行・タブ
    '--------------------------------------------------------
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

    text = Trim(text)

    '--------------------------------------------------------
    ' 末尾のピリオドまたは空白を削除
    '--------------------------------------------------------
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

    CleanPastSentFileName = text

End Function


'============================================================
' 実際の送信日時を取得
'============================================================

Private Function GetPastSentDate( _
    ByVal mail As Outlook.MailItem _
) As Date

    Dim d As Date

    On Error Resume Next

    d = mail.SentOn

    On Error GoTo 0

    If d = 0 Then
        d = Now
    End If

    GetPastSentDate = d

End Function


'============================================================
' 保存先
'
' Windowsのドキュメント
'     ↓
' 受信メール
'============================================================

Private Function GetPastSentSaveFolderPath() As String

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

    GetPastSentSaveFolderPath = _
        documentsPath & "\" & _
        SAVE_FOLDER_NAME

    Set shell = Nothing

End Function


'============================================================
' 保存先フォルダが無ければ作成
'============================================================

Private Sub EnsurePastSentFolderExists( _
    ByVal folderPath As String _
)

    Dim fso As Object

    Set fso = _
        CreateObject( _
            "Scripting.FileSystemObject" _
        )

    If Not fso.FolderExists(folderPath) Then

        fso.CreateFolder _
            folderPath

    End If

    Set fso = Nothing

End Sub
