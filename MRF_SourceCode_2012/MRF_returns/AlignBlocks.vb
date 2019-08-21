Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common

Public Class AlignBlocks
    Dim AcadApp As AcadApplication
    Dim AcadDoc As AcadDocument
    Dim kobj As AcadObject
    Dim kBlk, kentity, kblk3 As AcadBlockReference
    Dim kSubE, xdataout, xdatatype, xdataout1, xdatatype1, xdatatype2, xdataout2 As Array
    Dim kPnt, sd As Object
    Dim i, j As Integer
    Dim kCollection As New Collection
    Dim kBlk2 As AcadBlockReference
    Dim StartAngle, EndAngle, kRadius, BlkX, BlkY As Double
    Dim BlkMin, BlkMax As Object
    Dim appname(0 To 3), TA, TW, BA, KN As String
    Dim TypArray(0 To 6) As Int16
    Dim ValueArray(0 To 6) As Object
    Dim k As VariantType
    Dim kpnt1 As Object
    Dim iptime, optime As String
    Dim entity As AcadObject
    Dim entities As Object
    Dim midangle, stangle1, autocenter, width1 As Double
    Dim d(0 To 100) As Double

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        AcadApp = GetObject(, "AutoCAD.Application") : AcadDoc = AcadApp.ActiveDocument
        If Editcheck.Checked = True Then
            CHECKTRUE()
        Else
            CHECKFALSE()
        End If

    End Sub

    Sub CHECKTRUE()
        Me.Hide()
        For Each entity In AcadDoc.ModelSpace
            entity.GetXData("MRF-KKM1", xdatatype2, xdataout2)
            If xdataout2 Is Nothing Then
                GoTo 2
            ElseIf xdataout2(5).ToString = "MRFBLOCK" And xdataout2(1).ToString = "Input" Then
                If strTime = xdataout2(3).ToString Then
                    kobj = entity
                    Exit For
                End If
                Continue For
            End If
2:      Next
        operation()
    End Sub
    Sub CHECKFALSE()
        Static cnt As Long

        Me.Hide()
        Try
1:          AcadDoc.Utility.GetEntity(kobj, kPnt, "Select entity")

        Catch
            MsgBox("Nothing selected.")
            Me.Show()
            Exit Sub
            'cnt = cnt + 1
            'If cnt = 3 Then End
            'GoTo 1
        End Try
        operation()
    End Sub
    Sub operation()
        iptime = genID()
        optime = genID()

        If Not kobj.EntityName = "AcDbBlockReference" Then MsgBox("Selected entity is not a Block Reference") : Me.Show() : Exit Sub

        If kobj.EntityName = "AcDbBlockReference" Then
            kBlk = kobj
            kSubE = kBlk.Explode
            If (kSubE(0).entityname()) = "AcDbBlockReference" Then

                'CODE FOR VERIFYING XDATA AVAILABILITY
                'kBlk.GetXData("", xdatatype, xdataout)
                'If xdataout.Length = 0 Then
                'CreateXData(kobj, "MRFBLOCK", "Input", iptime, optime, StAngle.Text, EnAngle.Text, WrapRad.Text, Equispace.Checked)
                'Else
                '    iptime = xdataout(2).ToString
                'End If

                CreateXData(kobj, "MRF-KKM1", "Input", iptime, optime, StAngle.Text, "MRFBLOCK", WrapRad.Text, Equispace.Checked, TextBox1.Text, RadioButton1.Checked)

                'CODE FOR TO GET BOUNDING BOX
                kBlk.GetBoundingBox(BlkMin, BlkMax)
                '************************CODE FOR AUTOCENTER***********************************
                kRadius = Val(WrapRad.Text) : StartAngle = Val(StAngle.Text)

                If RadioButton2.Checked = True Then
                    StartAngle = 180
                ElseIf RadioButton1.Checked = True Then
                    StartAngle = Val(StAngle.Text)
                End If

                Dim d(0 To kSubE.Length - 1) As Double

                For i = 0 To kSubE.Length - 1
                    kBlk2 = kSubE(i)
                    width1 = kBlk2.InsertionPoint(0) - BlkMin(0)
                    d(i) = Rad2Ang((width1) / kRadius)
                Next

                'midangle = (d(0) - d(kSubE.Length - 2)) / 2

                'Array.Sort(d) Commented by prabu 07/07/2010
                'midangle = (d(kSubE.Length - 1) - d(0)) / 2

                midangle = Rad2Ang(((BlkMax(0) - BlkMin(0)) / 2) / kRadius) 'Resolve by prabu

                'stangle1 = midangle - d(0) - StartAngle

                If RadioButton2.Checked = True Then
                    autocenter = Val(TextBox1.Text) + midangle '+ 7
                    StartAngle = 0
                    'StartAngle = 10
                Else
                    autocenter = 0
                End If
                width1 = 0

                '******************************************************************************   
                For j = 0 To kSubE.Length - 1
                    kBlk2 = kSubE(j)
                    kCollection.Add(kBlk2)
                    BlkX = kBlk2.InsertionPoint(0) - BlkMin(0)
                    BlkY = kBlk2.InsertionPoint(1) - BlkMin(1)
                    kCollection(j + 1).InsertionPoint = PointAtAngle(kRadius, Rad2Ang((-BlkX) / kRadius) + StartAngle + autocenter)
                    kCollection(j + 1).Rotation = Ang2Rad(Rad2Ang((-BlkX) / kRadius) + StartAngle + 270 + autocenter)
                    kblk3 = kCollection(j + 1)
                    entities = kblk3.Explode
                    For i = LBound(entities) To UBound(entities)
                        CreateXData(entities(i), "MRF-KKM1", "Output", iptime, optime, StAngle.Text, "MRFBLOCK", WrapRad.Text, Equispace.Checked, TextBox1.Text, RadioButton1.Checked)
                    Next
                    kCollection(j + 1).delete()
                Next
            End If
        End If
        kCollection.Clear()

        If Editcheck.Checked = True Then
            Editcheck.Checked = False
            'Me.Hide()
            Me.Close()
        Else
            Me.Show()
            'Me.Close()
        End If
    End Sub

    Sub CreateXData(ByVal kobj As AcadObject, ByVal Appname As String, ByVal InROut As String, ByVal InputID As String, ByVal OutputID As String, ByVal StrtAng As String, ByVal ObjectName As String, ByVal WRad As String, ByVal AttachPt As String, ByVal ACenterAt As String, ByVal IsStAngle As String)
        Dim TypArray(0 To 9) As Int16
        Dim ValueArray(0 To 9) As Object
        TypArray(0) = 1001 : ValueArray(0) = Appname
        TypArray(1) = 1000 : ValueArray(1) = InROut
        TypArray(2) = 1000 : ValueArray(2) = InputID
        TypArray(3) = 1000 : ValueArray(3) = OutputID
        TypArray(4) = 1000 : ValueArray(4) = StrtAng
        TypArray(5) = 1000 : ValueArray(5) = ObjectName
        TypArray(6) = 1000 : ValueArray(6) = WRad
        TypArray(7) = 1000 : ValueArray(7) = AttachPt
        TypArray(8) = 1000 : ValueArray(8) = ACenterAt
        TypArray(9) = 1000 : ValueArray(9) = IsStAngle

        kobj.SetXData(TypArray, ValueArray)
    End Sub
    Function PointAtAngle(ByVal kDist As Double, ByVal kAng As Double) As Object
        Dim kPt(0 To 2) As Double
        kPt(0) = kDist * Math.Cos(Ang2Rad(kAng))
        kPt(1) = kDist * Math.Sin(Ang2Rad(kAng))
        kPt(2) = 0
        PointAtAngle = kPt
    End Function
    Function Ang2Rad(ByVal Ang As Double) As Double
        Ang2Rad = Ang * (Math.PI / 180)
    End Function
    Function Rad2Ang(ByVal Rad As Double) As Double
        Rad2Ang = Rad * (180 / Math.PI)
    End Function
    Function genID() As String
        genID = DateTime.Now.Day & "-" & DateTime.Now.Month & "-" & DateTime.Now.Year & "-" & DateTime.Now.Hour & "-" & DateTime.Now.Minute & "-" & DateTime.Now.Second
    End Function
End Class