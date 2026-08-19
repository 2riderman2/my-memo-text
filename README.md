Option Explicit

'============================================================
' 過去の受信メールを選択フォルダから一括保存
'
' 保存先：
'   Documents\受信メール
'
' ファイル名：
'   yyyyMMdd_HHmm_差出人_件名.msg
'
' 差出人名：
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
' SaveAllPastReceivedMails
' を実行
'============================================================

Public Sub SaveAllPastReceivedMails()

    Dim sourceFolder As Outlook.Folder

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
    ' 保存元フォルダを選択
    '--------------------------------------------------------
    Set sourceFolder = _
        Application.Session.PickFolder

    ' キャンセル
    If sourceFolder Is Nothing Then
        Exit Sub
    End If

    '--------------------------------------------------------
    ' 実行確認
    '--------------------------------------------------------
    result = MsgBox( _
        "次のOutlookフォルダにあるメールをすべて保存します。" & _
        vbCrLf & vbCrLf & _
        sourceFolder.FolderPath & _
        vbCrLf & vbCrLf & _
        "保存先：" & vbCrLf & _
        "ドキュメント\" & SAVE_FOLDER_NAME & _
        vbCrLf & vbCrLf & _
        "すでに同名ファイルがあるメールはスキップします。" & _
        vbCrLf & vbCrLf & _
        "実行しますか？", _
        vbYesNo + vbQuestion, _
        "過去の受信メール一括保存" _
    )

    If result <> vbYes Then
        Exit Sub
    End If

    '--------------------------------------------------------
    ' 保存先
    '--------------------------------------------------------
    saveFolder = _
        GetPastReceivedSaveFolderPath()

    EnsurePastReceivedFolderExists _
        saveFolder

    '--------------------------------------------------------
    ' 件数初期化
    '--------------------------------------------------------
    totalCount = 0
    savedCount = 0
    skippedCount = 0
    errorCount = 0

    '--------------------------------------------------------
    ' 選択したフォルダ内を全件処理
    '--------------------------------------------------------
    For Each obj In sourceFolder.Items

        ' メールだけ対象
        If TypeOf obj Is Outlook.MailItem Then

            Set mail = obj

            totalCount = _
                totalCount + 1

            On Error GoTo ItemError

            '------------------------------------------------
            ' ファイル名
            '------------------------------------------------
            fileName = _
                BuildPastReceivedFileName(mail)

            fullPath = _
                saveFolder & "\" & fileName

            '------------------------------------------------
            ' 同名ファイルあり
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
        "過去の受信メールの保存が完了しました。" & _
        vbCrLf & vbCrLf & _
        "選択フォルダ：" & vbCrLf & _
        sourceFolder.FolderPath & _
        vbCrLf & vbCrLf & _
        "対象メール：" & totalCount & "件" & vbCrLf & _
        "新規保存：" & savedCount & "件" & vbCrLf & _
        "同名のためスキップ：" & skippedCount & "件" & vbCrLf & _
        "エラー：" & errorCount & "件" & vbCrLf & vbCrLf & _
        "保存先：" & vbCrLf & _
        saveFolder, _
        vbInformation, _
        "過去の受信メール一括保存"

    Exit Sub


'------------------------------------------------------------
' 1メールだけエラー
'------------------------------------------------------------
ItemError:

    errorCount = _
        errorCount + 1

    Debug.Print _
        "SaveAllPastReceivedMails Item Error " & _
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
        "過去の受信メール一括保存"

End Sub


'============================================================
' ファイル名作成
'
' yyyyMMdd_HHmm_差出人_件名.msg
'============================================================

Private Function BuildPastReceivedFileName( _
    ByVal mail As Outlook.MailItem _
) As String

    Dim receivedDate As Date

    Dim dateText As String
    Dim senderText As String
    Dim subjectText As String

    Dim result As String

    On Error GoTo ErrHandler

    '--------------------------------------------------------
    ' 受信日時
    '--------------------------------------------------------
    receivedDate = _
        GetPastReceivedDate(mail)

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

    ' 「.共有メールボックス」を削除
    senderText = _
        RemovePastReceivedSharedMailboxText( _
            senderText _
        )

    ' ( または （ 以降を削除
    senderText = _
        RemovePastReceivedParenthesisPart( _
            senderText _
        )

    ' ファイル名用に整形
    senderText = _
        CleanPastReceivedFileName( _
            senderText _
        )

    If senderText = "" Then
        senderText = "差出人不明"
    End If

    '--------------------------------------------------------
    ' 件名
    '--------------------------------------------------------
    subjectText = _
        Trim(mail.Subject)

    If subjectText = "" Then
        subjectText = "件名なし"
    End If

    subjectText = _
        CleanPastReceivedFileName( _
            subjectText _
        )

    If subjectText = "" Then
        subjectText = "件名なし"
    End If

    '--------------------------------------------------------
    ' 組み立て
    '--------------------------------------------------------
    result = _
        dateText & "_" & _
        senderText & "_" & _
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

    BuildPastReceivedFileName = _
        result & ".msg"

    Exit Function


ErrHandler:

    BuildPastReceivedFileName = _
        Format( _
            Now, _
            "yyyymmdd_HHmm" _
        ) & _
        "_受信メール.msg"

End Function


'============================================================
' 「.共有メールボックス」を削除
'============================================================

Private Function RemovePastReceivedSharedMailboxText( _
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

    RemovePastReceivedSharedMailboxText = _
        Trim(text)

End Function


'============================================================
' 「(」または「（」以降を削除
'============================================================

Private Function RemovePastReceivedParenthesisPart( _
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

    RemovePastReceivedParenthesisPart = _
        Trim(text)

End Function


'============================================================
' Windowsファイル名として使用できるように整形
'============================================================

Private Function CleanPastReceivedFileName( _
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
    ' 禁止文字
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
    ' 末尾のピリオド・空白
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

    CleanPastReceivedFileName = _
        text

End Function


'============================================================
' 受信日時取得
'============================================================

Private Function GetPastReceivedDate( _
    ByVal mail As Outlook.MailItem _
) As Date

    Dim d As Date

    On Error Resume Next

    d = mail.ReceivedTime

    On Error GoTo 0

    If d = 0 Then
        d = Now
    End If

    GetPastReceivedDate = d

End Function


'============================================================
' 保存先
'
' Documents\受信メール
'============================================================

Private Function GetPastReceivedSaveFolderPath() As String

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

    GetPastReceivedSaveFolderPath = _
        documentsPath & "\" & _
        SAVE_FOLDER_NAME

    Set shell = Nothing

End Function


'============================================================
' 保存先フォルダが無ければ作成
'============================================================

Private Sub EnsurePastReceivedFolderExists( _
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
