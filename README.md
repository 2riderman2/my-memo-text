let
    txt = [摘要],
    pos = if txt = null then -1 else Text.PositionOf(txt, "購入"),
    result =
        if pos >= 0 then
            let
                after = Text.Range(txt, pos + 2),
                digits = "０１２３４５６７８９"
            in
                if Text.Length(after) >= 2
                    and Text.Start(after, 1) = "　"
                    and Text.Contains(digits, Text.Middle(after, 1, 1))
                then
                    Text.Start(txt, pos + 2)
                else
                    txt
        else
            txt
in
    result
