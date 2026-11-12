let
    txt = [摘要],
    pos = if txt = null then -1 else Text.PositionOf(txt, "（")
in
    if txt <> null
        and Text.Contains(txt, "健康診断")
        and pos >= 0
    then
        Text.Start(txt, pos)
    else
        txt
