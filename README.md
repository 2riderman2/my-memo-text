let
    txt = [摘要],
    pos = if txt = null then -1 else Text.PositionOf(txt, "月分"),
    before = if pos >= 0 then Text.Start(txt, pos) else "",
    digits = "０１２３４５６７８９",
    n = List.Count(
        List.FirstN(
            Text.ToList(Text.Reverse(before)),
            each Text.Contains(digits, _)
        )
    )
in
    if pos >= 0 and n > 0 then
        Text.Start(txt, pos - n) & Text.Range(txt, pos + 2)
    else
        txt
