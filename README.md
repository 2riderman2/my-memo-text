let
    txt = [摘要],
    pos = if txt = null then -1 else Text.PositionOf(txt, "保険料"),
    result =
        if pos >= 0 then
            let
                after = Text.Range(txt, pos + 3)
            in
                if Text.StartsWith(after, "　") then
                    Text.Start(txt, pos + 3)
                else
                    txt
        else
            txt
in
    result
