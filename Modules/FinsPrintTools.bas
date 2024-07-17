Attribute VB_Name = "FinsPrintTools"
'Print Tools
'http://www.arief.in/

Public Sub PrintCell(ByVal objectPrint As Object, ByVal xPos As Long, ByVal yPos As Long, ByVal cellWidth As Long, ByVal cellHeight As Long, Optional str As String, Optional cellFontSize As Long, Optional line As Boolean, Optional hAlign As Integer, Optional vAlign As Integer, Optional paddingLeft As Integer, Optional paddingTop As Integer, Optional paddingRight As Integer, Optional paddingBottom As Integer)
    cellWidth = cellWidth + paddingLeft + paddingRight
    cellHeight = cellHeight + paddingTop + paddingBottom
    If cellFontSize = 0 Then
        cellFontSize = 10
    End If
    If line = True Then
        objectPrint.Line (xPos, yPos)-Step(cellWidth, cellHeight), , B
    End If
    objectPrint.FontSize = cellFontSize
    If hAlign = 1 Then 'horizontal align : right
        objectPrint.CurrentX = xPos + cellWidth - objectPrint.TextWidth(str)
    ElseIf hAlign = 2 Then 'horizontal align : center
        objectPrint.CurrentX = xPos + (cellWidth / 2) - (objectPrint.TextWidth(str) / 2) - paddingRight
    Else 'horizontal align : left
        objectPrint.CurrentX = xPos + paddingLeft
    End If
    If vAlign = 1 Then 'vertical align : bottom
        objectPrint.CurrentY = yPos + cellHeight - objectPrint.TextHeight(str) - paddingBottom
    ElseIf vAlign = 2 Then 'vertical align : middle
        objectPrint.CurrentY = yPos + (cellHeight / 2) - (objectPrint.TextHeight(str) / 2)
    Else 'vertical align : top
        objectPrint.CurrentY = yPos + paddingTop
    End If
    objectPrint.Print str
End Sub
