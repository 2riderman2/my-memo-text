let
    txt = [摘要],
    pos = if txt = null then -1 else Text.PositionOf(txt, "）")
in
    if txt <> null
        and Text.Contains(txt, "加除式図書")
        and pos >= 0
    then
        Text.Start(txt, pos + 1)
    else
        txt
