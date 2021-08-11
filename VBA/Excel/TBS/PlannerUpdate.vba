'Not sure'

Sub Planner_Update()

Dim wbPlan As Workbook
Dim wbSFWB As Workbook

Application.ScreenUpdating = False

Set wbSFWB = Workbooks.Open("M:\Supply Chain\Planning\FG Planning\Pre-Pack Schedule\TBSUKShopFloorWorkbench.csv")
Set wbPlan = ThisWorkbook

With wbPlan.Worksheets("SFW-DATA").Cells.ClearContents

End With
         With wbSFWB.Sheets(1).Rows(1)
        Set a = .Find("Shop Order No")
        If Not a Is Nothing Then
            Columns(a.Column).EntireColumn.Copy _
            Destination:=wbPlan.Worksheets("SFW-DATA").Range("A1")
        Else: MsgBox "Shop Order No Not Found"
        End If
        End With

         With wbSFWB.Sheets(1).Rows(1)
        Set b = .Find("Part No")
        If Not b Is Nothing Then
            Columns(b.Column).EntireColumn.Copy _
            Destination:=wbPlan.Worksheets("SFW-DATA").Range("B1")
        Else: MsgBox "Part No Not Found"
        End If
        End With

          With wbSFWB.Sheets(1).Rows(1)
        Set c = .Find("Executable Qty")
        If Not c Is Nothing Then
            Columns(c.Column).EntireColumn.Copy _
            Destination:=wbPlan.Worksheets("SFW-DATA").Range("C1")
        Else: MsgBox "Executable Qty Not Found"
        End If
        End With

        With wbSFWB.Sheets(1).Rows(1)
        Set d = .Find("Remaining Qty")
        If Not d Is Nothing Then
            Columns(d.Column).EntireColumn.Copy _
            Destination:=wbPlan.Worksheets("SFW-DATA").Range("D1")
        Else: MsgBox "Remaining Qty Not Found"
        End If
        End With

        With wbSFWB.Sheets(1).Rows(1)
        Set e = .Find("Vial")
        If Not e Is Nothing Then
            Columns(e.Column).EntireColumn.Copy _
            Destination:=wbPlan.Worksheets("SFW-DATA").Range("E1")
        Else: MsgBox "Vial Not Found"
        End If
        End With

         With wbSFWB.Sheets(1).Rows(1)
        Set f = .Find("Label")
        If Not f Is Nothing Then
            Columns(f.Column).EntireColumn.Copy _
            Destination:=wbPlan.Worksheets("SFW-DATA").Range("F1")
        Else: MsgBox "Label Not Found"
        End If
        End With

         With wbSFWB.Sheets(1).Rows(1)
        Set g = .Find("Cap")
        If Not g Is Nothing Then
            Columns(g.Column).EntireColumn.Copy _
            Destination:=wbPlan.Worksheets("SFW-DATA").Range("G1")
        Else: MsgBox "Cap Not Found"
        End If
        End With

        Windows("TBSUKShopFloorWorkbench.csv").Close

    Application.ScreenUpdating = True

End Sub
