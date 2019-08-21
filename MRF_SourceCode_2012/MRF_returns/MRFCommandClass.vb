Imports System
Imports System.Type
Imports System.CLSCompliantAttribute
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices

Public Class MRFCommandClass
    <CommandMethod("EO")>
    Public Sub EngraveRun()
        EngraveOffsetRun()
    End Sub
    <CommandMethod("ED")>
    Public Sub NewEdit()
        EditOffset()
    End Sub

    <CommandMethod("AB")>
    Public Sub NewAlignBlocks()
        NewAlignBlock()
    End Sub

    <CommandMethod("GXD")> _
    Public Shared Sub GetXData()

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        ' Ask the user to select an entity 

        ' for which to retrieve XData 
        Dim opt As New PromptEntityOptions(vbLf & "Select entity: ")
        Dim res As PromptEntityResult = ed.GetEntity(opt)
        If res.Status = PromptStatus.OK Then
            Dim tr As Transaction = doc.TransactionManager.StartTransaction()
            Using tr
                Dim obj As DBObject = tr.GetObject(res.ObjectId, OpenMode.ForRead)
                Dim rb As ResultBuffer = obj.XData
                If rb Is Nothing Then
                    ed.WriteMessage(vbLf & "Entity does not have XData attached.")
                Else
                    Dim n As Integer = 0
                    For Each tv As TypedValue In rb
                        ed.WriteMessage(vbLf & "TypedValue {0} - type: {1}, value: {2}", System.Math.Max(System.Threading.Interlocked.Increment(n), n - 1), tv.TypeCode, tv.Value)
                    Next
                    rb.Dispose()
                End If
            End Using
        End If
    End Sub

End Class
