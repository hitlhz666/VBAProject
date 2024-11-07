Sub DrawRectanglesWithConnectionsInVisio()
    Dim VisioApp As Object
    Dim VisioDoc As Object
    Dim VisioPage As Object
    
    Dim ExcelSheet As Worksheet
    Dim lastRow As Long
    Dim ShapesDict As Object
    Dim i As Integer
    Dim rectId As Integer
    Dim xPos As Double
    Dim yPos As Double
    Dim width As Double
    Dim height As Double
    Dim rectText As String
    
    Dim Shape As Object
    Dim j As Integer
    Dim connectToList As Variant
    Dim connectToID As Integer
    Dim targetShape As Object

    Dim connectLine As Visio.Shape
    Dim vsoCell1 As Visio.Cell
    Dim vsoCell2 As Visio.Cell
    
    ' Create Visio application object
    On Error Resume Next                                 ' 之后程序出现“运行时错误”时，不中断继续运行
    Set VisioApp = GetObject(, "Visio.Application")      ' Try to hook Visio if it's already running
    If VisioApp Is Nothing Then
        Set VisioApp = CreateObject("Visio.Application") ' If Visio isn't running, create a new instance
    End If
    On Error GoTo 0                                      ' 之后程序出现“运行时错误”时，中断并报错
    
    VisioApp.Visible = True ' Make Visio visible
    
    ' Create a new document if no document is open
    If VisioApp.Documents.Count = 0 Then
        Set VisioDoc = VisioApp.Documents.Add("")
    Else
        Set VisioDoc = VisioApp.ActiveDocument
    End If
    
    ' Get the active page in Visio
    Set VisioPage = VisioDoc.Pages(1)
    
    ' 设置Excel工作表
    Set ExcelSheet = ThisWorkbook.Sheets("Sheet1")
    
    ' 获取工作表中的最后一行
    lastRow = ExcelSheet.UsedRange.Rows.Count
    
    ' Create a dictionary to store shapes by their ID for easy reference
    Set ShapesDict = CreateObject("Scripting.Dictionary")
    
    ' Step 1: Loop through the rows in Excel and draw all the rectangles
    For i = 2 To lastRow ' Assume data starts from row 2
        rectId = ExcelSheet.Cells(i, 1).value
        xPos = ExcelSheet.Cells(i, 2).value
        yPos = ExcelSheet.Cells(i, 3).value
        width = ExcelSheet.Cells(i, 4).value
        height = ExcelSheet.Cells(i, 5).value
        rectText = ExcelSheet.Cells(i, 6).value ' Text from the new column
        
        Set Shape = VisioPage.DrawRectangle(xPos, yPos, xPos + width, yPos + height)
        Shape.Text = rectText
        
        ' Store the shape in the dictionary with its RectID as the key
        ShapesDict.Add rectId, Shape
    Next i
    
    ' Step 2: Now loop through each rectangle and create the connections
    For i = 2 To lastRow ' Again start from row 2
        rectId = ExcelSheet.Cells(i, 1).value
        connectToList = Split(ThisWorkbook.Sheets(1).Cells(i, 7).value, ";") ' Split ConnectTo into an array
        
        ' Loop through each entry in the connectToList (splitted by semicolon)
        For j = LBound(connectToList) To UBound(connectToList)
            connectToID = Trim(connectToList(j)) ' Trim用于去除字符串两端的空格
            If connectToID <> 0 Then
                If ShapesDict.Exists(connectToID) Then
                    Set targetShape = ShapesDict(connectToID)
                    Set Shape = ShapesDict(rectId)
                    
                    Set connectLine = VisioPage.Drop(VisioApp.ConnectorToolDataObject, 0#, 0#)
                    VisioPage.Shapes.ItemFromID(connectLine.id).CellsSRC(visSectionObject, visRowLine, visLineEndArrow).FormulaU = "13" ' 设置连接线的箭头
                    
                    Set vsoCell1 = VisioPage.Shapes.ItemFromID(connectLine.id).CellsU("BeginX")
                    Set vsoCell2 = VisioPage.Shapes.ItemFromID(Shape.id).CellsSRC(1, 1, 0)
                    vsoCell1.GlueTo vsoCell2
                    Set vsoCell1 = VisioPage.Shapes.ItemFromID(connectLine.id).CellsU("EndX")
                    Set vsoCell2 = VisioPage.Shapes.ItemFromID(targetShape.id).CellsSRC(1, 1, 0)
                    vsoCell1.GlueTo vsoCell2
                Else
                    MsgBox "没有序号" & connectToID & "的矩形，无法连接线！"
                End If
            End If
        Next j
    Next i
    MsgBox "矩形和连接线绘制完成！"
    
    If Not VisioPage Is Nothing Then Set VisioPage = Nothing
    If Not VisioDoc Is Nothing Then Set VisioDoc = Nothing
    If Not VisioApp Is Nothing Then
        ' Optionally close Visio if it was started by the script (this is commented out for safety)
        ' VisioApp.Quit
        VisioApp.Quit ' This will close Visio completely
        Set VisioApp = Nothing
    End If
End Sub
