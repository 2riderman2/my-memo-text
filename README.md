'============================================================
' 「.共有メールボックス」を削除
'
' 例：
' 山田太郎.共有メールボックス
' → 山田太郎
'============================================================

Private Function RemoveSharedMailboxText( _
    ByVal text As String _
) As String

    text = Replace( _
        text, _
        ".共有メールボックス", _
        "", _
        1, _
        -1, _
        vbTextCompare _
    )

    RemoveSharedMailboxText = Trim(text)

End Function
