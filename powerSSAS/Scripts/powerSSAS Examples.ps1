#add-PSSnapin powerssas

Invoke-AsMdx -ServerName localhost\sql08 -DatabaseName "Adventure Works DW 2008" -Query "SELECT [Measures].[Internet Sales Amount] on 0, [Geography].[Country].members ON 1  FROM [Adventure Works]"

get-aslock localhost\sql08

$svr = get-AsServer localhost\sql08

get-AsDimension localhost\sql08 "Adventure Works DW 2008" "Geography" | new-AsScript $dim > c:\data\geog.xmla

