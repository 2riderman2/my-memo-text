let
    txt = [摘要],
    digits = "０１２３４５６７８９",
    pos =
        if txt = null then -1
        else Text.PositionOfAny(txt, Text.ToList(digits))
in
    if txt <> null
        and Text.Contains(txt, "非常勤賞与")
        and pos >= 0
    then
        Text.Start(txt, pos)
    else
        txt
