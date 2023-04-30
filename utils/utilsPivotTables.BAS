Option Explicit


Public Function makePivotTable( _
    ByRef ws As Worksheet, _
    ByRef sourceRange As Range, _
    tn As String, _
    firstCol As Long, _
    lastCol As Long)
    
    ws.Activate
    
    Dim PTCache As PivotCache
    Dim PT As PivotTable
    Set PT = PTCache.CreatePivotTable( _
        TableDestination:=ws.Cells(firstCol, lastCol), tableName:=tn)
    PT.ManualUpdate = True
    
    
End Function


