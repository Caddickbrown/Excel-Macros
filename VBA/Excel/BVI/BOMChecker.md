

=IFERROR(SUM(SUMIFS(Released!M:M,Released!D:D,FILTER(ManStru!C:C,A5=ManStru!N:N))*INDEX(ManStru!V:V,MATCH(A5&FILTER(ManStru!C:C,A5=ManStru!N:N),ManStru!N:N&ManStru!C:C,0))),"-")