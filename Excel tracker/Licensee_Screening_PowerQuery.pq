let
    Source = Excel.CurrentWorkbook(){[Name="tblRawShipments"]}[Content],
    Typed = Table.TransformColumnTypes(Source, {{"Month", type date}, {"Company", type text}, {"Region", type text}, {"Units", Int64.Type}}),
    Cleaned = Table.TransformColumns(Typed, {{"Company", each Text.Proper(Text.Trim(_)), type text}}),
    Valid = Table.SelectRows(Cleaned, each [Units] <> null and [Units] > 0),
    Grouped = Table.Group(Valid, {"Company"}, {{"Units (12m)", each List.Sum([Units]), Int64.Type}, {"Months reported", each Table.RowCount(_), Int64.Type}}),
    Sorted = Table.Sort(Grouped, {{"Units (12m)", Order.Descending}})
in
    Sorted
