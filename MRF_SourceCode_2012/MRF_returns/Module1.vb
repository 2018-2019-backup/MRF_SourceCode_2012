Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports System.IO
Imports System.Runtime.InteropServices.COMException
Imports System.Data

Module Module1
    Public currenttext As String
    Public frmInsertBlock As InsertBlock
    Public frmAlignBlocks As AlignBlocks
    Public Const AngleDeviationCorrectionFactor As Double = 0 '0.15
    Public strInRout, strTime As String
    Public acad As AcadDocument
    Public previous As String

    Dim kobj1 As AcadObject : Dim kpnt1 As Object
    Dim AcadApp As AcadApplication : Dim AcadDoc As AcadDocument
    Dim xdatatype1, xdataout1, xdatatype2, xdataout2 As Object
    Dim entity As AcadObject : Dim b1 As AcadBlockReference
    Dim i, j, k As Integer

    Public Sub EDIT()
        AcadApp = GetObject(, "AutoCAD.Application") : AcadDoc = AcadApp.ActiveDocument
        Try
1:          AcadDoc.Utility.GetEntity(kobj1, kpnt1, "Select entity")
        Catch
            MsgBox("Nothing selected.")
            Exit Sub
            'MsgBox("Select Again")
            'GoTo 1
        End Try
        Dim startangle1, radius1 As Double

        Try
            kobj1.GetXData("MRF-KKM1", xdatatype1, xdataout1)

            strInRout = xdataout1(1)

            strTime = xdataout1(3)
        Catch ex As Exception
            MsgBox("Selected entity can not be editable")
            Exit Sub
        End Try

        startangle1 = xdataout1(4) : radius1 = Val(xdataout1(6))
        'kobj1.GetXData("", xdatatype1, xdataout1)

        If xdataout1.length >= 5 Then
            ''**********code for delete block*************
            'frmAlignBlocks = New AlignBlocks

            'If xdataout1(5) = "MRFBLOCK" Then
            '    frmAlignBlocks.StAngle.Text = startangle1
            '    frmAlignBlocks.WrapRad.Text = radius1
            '    frmAlignBlocks.TextBox1.Text = xdataout1(8)

            '    If xdataout1(9) = "True" Then
            '        frmAlignBlocks.RadioButton1.Checked = True
            '    Else
            '        frmAlignBlocks.RadioButton2.Checked = True
            '    End If

            '    objdelete("MRFBLOCK")
            '    frmAlignBlocks.Editcheck.Checked = True
            '    frmAlignBlocks.Show()
            '    '************code for delete the wrap text**********
            'ElseIf xdataout1(5) = "MRFWRAP" Then
            '    frmWrapSketch = New WrapText

            '    frmWrapSketch.StAngle.Text = startangle1
            '    frmWrapSketch.WrapRad.Text = radius1
            '    objdelete("MRFWRAP")
            '    frmWrapSketch.CheckBox1.Checked = True
            '    frmWrapSketch.Show()
            '    '************code for delete the mtext**********
            'ElseIf xdataout1(5) = "MRFTEXT" Then

            '    frmAlignText = New AlignText

            '    Dim AttachPt() As String = xdataout1(7).ToString.Split("/")

            '    frmAlignText.StAngle.Text = startangle1
            '    frmAlignText.WrapRad.Text = radius1

            '    frmAlignText.ComboBox1.Text = AttachPt(0).ToString
            '    frmAlignText.ComboBox2.Text = AttachPt(1).ToString
            '    frmAlignText.TextBox1.Text = xdataout1(8).ToString

            '    If xdataout1(9) = True Then
            '        frmAlignText.RadioButton1.Checked = True
            '        frmAlignText.RadioButton2.Checked = False
            '    Else
            '        frmAlignText.RadioButton1.Checked = False
            '        frmAlignText.RadioButton2.Checked = True
            '    End If

            '    objdelete("MRFTEXT")
            '    frmAlignText.CheckBox1.Checked = True

            '    frmAlignText.Show()

            'ElseIf xdataout1(5) = "ALPHA" Then

            '    frmInsertTextBlock = New InsertTextBlock

            '    frmInsertTextBlock.TextBox1.Text = xdataout1(8)
            '    frmInsertTextBlock.TxtHt.Text = xdataout1(9)
            '    frmInsertTextBlock.StAngle.Text = xdataout1(4)
            '    frmInsertTextBlock.EndAngle.Text = xdataout1(6)
            '    frmInsertTextBlock.GrvWd.Text = xdataout1(10)
            '    frmInsertTextBlock.SpFactor.Text = xdataout1(11)
            '    frmInsertTextBlock.Radius.Text = xdataout1(7)

            '    If xdataout1(12) = "True" Then
            '        frmInsertTextBlock.RadioButton1.Checked = True
            '    Else
            '        frmInsertTextBlock.RadioButton2.Checked = True
            '    End If


            '    objdelete("ALPHA")
            '    frmInsertTextBlock.CheckBox1.Checked = True

            '    frmInsertTextBlock.Show()

            'ElseIf xdataout1(5) = "MRFKNURLING" Then
            '    frmKnurling = New AddKnurling
            '    frmKnurling.StAngletxt.Text = xdataout1(4)
            '    frmKnurling.MaxRad.Text = xdataout1(6)
            '    frmKnurling.EnAngletxt.Text = xdataout1(7)
            '    If xdataout1(8) = "True" Then
            '        frmKnurling.FullRotn.Checked = True
            '    Else
            '        frmKnurling.AngleRotn.Checked = True
            '    End If

            '    If xdataout1(9) = "True" Then
            '        frmKnurling.NumberDist.Checked = True
            '    Else
            '        frmKnurling.AngularDist.Checked = True
            '    End If

            '    frmKnurling.SpaceAngle.Text = xdataout1(10)
            '    frmKnurling.SpaceDist.Text = xdataout1(11)

            '    objdelete("MRFKNURLING")

            '    frmKnurling.Show()

            'End If

        End If
    End Sub

    Public Sub EditOffset()
        AcadApp = GetObject(, "AutoCAD.Application") : AcadDoc = AcadApp.ActiveDocument
        Try
1:          AcadDoc.Utility.GetEntity(kobj1, kpnt1, "Select entity")
        Catch
            MsgBox("Nothing selected.")
            Exit Sub
            'MsgBox("Select Again")
            'GoTo 1
        End Try
        Dim startangle1, radius1 As Double

        Try
            kobj1.GetXData("MRF-KKM1", xdatatype1, xdataout1)


            strInRout = xdataout1(1)

            strTime = xdataout1(3)
        Catch ex As Exception
            MsgBox("Selected entity can not be editable")
            Exit Sub
        End Try

        startangle1 = xdataout1(4) : radius1 = Val(xdataout1(6))
        'kobj1.GetXData("", xdatatype1, xdataout1)

        If xdataout1.length >= 5 Then
            '**********code for delete block*************
            If xdataout1(5) = "MRFBLOCK" Then
                frmAlignBlocks = New AlignBlocks
                frmAlignBlocks.StAngle.Text = startangle1
                frmAlignBlocks.WrapRad.Text = radius1
                frmAlignBlocks.TextBox1.Text = xdataout1(8)

                If xdataout1(9) = "True" Then
                    frmAlignBlocks.RadioButton1.Checked = True
                Else
                    frmAlignBlocks.RadioButton2.Checked = True
                End If

                objdelete("MRFBLOCK")
                frmAlignBlocks.Editcheck.Checked = True
                frmAlignBlocks.Show()

            ElseIf xdataout1(5) = "ALPHA" Then
                'currenttext = ""

                ' frmInsertTextBlock = New InsertTextBlock
                frmInsertBlock = New InsertBlock
                currenttext = xdataout1(13)
                frmInsertBlock.TxtText_KeyUp()
                frmInsertBlock.TxtText.Text = xdataout1(8)
                frmInsertBlock.TxtHt.Text = xdataout1(9).ToString.Split("&")(0)
                frmInsertBlock.StAngle.Text = xdataout1(4)
                frmInsertBlock.EndAngle.Text = xdataout1(6)
                frmInsertBlock.GrvWd.Text = xdataout1(10)
                frmInsertBlock.SpFactor.Text = xdataout1(11)
                frmInsertBlock.Radius.Text = xdataout1(7)
                frmInsertBlock.txt_Angle.Text = xdataout1(9).ToString.Split("&")(2)
                frmInsertBlock.txtDist.Text = xdataout1(9).ToString.Split("&")(3)
                frmInsertBlock.TXT_L_Angle.Text = xdataout1(9).ToString.Split("&")(4)
                If xdataout1(12) = "True" Then
                    frmInsertBlock.RadioButton1.Checked = True
                Else
                    frmInsertBlock.RadioButton2.Checked = True
                End If


                objdelete("ALPHA")
                frmInsertBlock.CheckBox1.Checked = True

                frmInsertBlock.Show()
                frmInsertBlock.cmb_Font.SelectedIndex = Val(xdataout1(9).ToString.Split("&")(1))
            ElseIf xdataout1(5) = "MRFKNURL" Then

            End If

        End If
    End Sub

    Public Sub objdelete(ByVal objname As String)
        'For Each entity As Object In AcadDoc.ModelSpace

        '    entity.GetXData("MRF-KKM1", xdatatype2, xdataout2)
        '    If xdataout2 Is Nothing Then
        '        i = i + 1
        '        Continue For
        '    ElseIf xdataout2(5).ToString = objname And xdataout2(3).ToString = xdataout1(3).ToString And xdataout2(1) = "Output" Then
        '        k = k + 1
        '        entity.Delete()
        '    End If


        'Next
        For Each entity In AcadDoc.ModelSpace
            entity.GetXData("MRF-KKM1", xdatatype2, xdataout2)
            If xdataout2 Is Nothing Then
                i = i + 1
                GoTo 2
            ElseIf xdataout2(5).ToString = objname And xdataout2(3).ToString = xdataout1(3).ToString And xdataout2(1) = "Output" Then
                k = k + 1
                entity.Delete()
            End If
2:      Next

    End Sub

    Public Sub objdelete(ByVal objname As String, ByVal AcadDoc As AcadDocument)
        For Each entity In AcadDoc.ModelSpace
            entity.GetXData("MRF-KKM1", xdatatype2, xdataout2)
            If xdataout2 Is Nothing Then
                i = i + 1
                GoTo 2
            Else
                k = k + 1
                entity.Delete()
            End If
2:      Next
    End Sub

    Public Sub EngraveOffsetRun()
        Dim inser As New InsertBlock
        inser.Show()
    End Sub

    Public Sub NewAlignBlock()
        Dim frmAlign As New AlignBlocks
        frmAlign.Show()
    End Sub
    Function GetMaxX(ByVal kBlk As AcadBlock) As Double
        Dim minpnt, maxpnt As Object
        Dim kmax(0 To kBlk.Count - 1) As Double
        For i As Integer = 0 To kBlk.Count - 1
            kBlk.Item(i).GetBoundingBox(minpnt, maxpnt)
            kmax(i) = maxpnt(0)
        Next
        Array.Sort(kmax)
        GetMaxX = kmax(kmax.Length - 1)
    End Function

    Function GetMinX(ByVal kBlk As AcadBlock) As Double
        Dim minpnt, maxpnt As Object
        Dim kmin(0 To kBlk.Count - 1) As Double
        For i As Integer = 0 To kBlk.Count - 1
            kBlk.Item(i).GetBoundingBox(minpnt, maxpnt)
            kmin(i) = minpnt(0)
        Next
        Array.Sort(kmin)
        GetMinX = kmin(0)
    End Function

    Function GetBlkname(ByVal Str As String) As String
        Select Case Str
            Case " ", "" : Return "sp"
            Case "?" : Return "qu"
            Case ":" : Return "c1"
            Case ";" : Return "c2"
            Case ">" : Return "gr"
            Case "<" : Return "le"
            Case "-" : Return "mi"
            Case "_" : Return "ud"
            Case "/" : Return "s1"
            Case "\" : Return "s2"
            Case "*" : Return "st"
            Case "|" : Return "pip"
            Case "," : Return "comma"
            Case "=" : Return "eq"
            Case "'" : Return "qo"
            Case "." : Return "dot"
            Case "A" To "Z" : Return "c" & Str
            Case Else
                Return Str
        End Select
    End Function

    Function GetBlkname(ByVal Str As String, ByVal txt As ComboBox) As String
        Select Case Str
            Case " ", "" : Return "font" + txt.Text + "_" + "sp"
            Case "?" : Return "font" + txt.Text + "_" + "qu"
            Case ":" : Return "font" + txt.Text + "_" + "c1"
            Case ";" : Return "font" + txt.Text + "_" + "c2"
            Case ">" : Return "font" + txt.Text + "_" + "gr"
            Case "<" : Return "font" + txt.Text + "_" + "le"
            Case "-" : Return "font" + txt.Text + "_" + "mi"
            Case "_" : Return "font" + txt.Text + "_" + "ud"
            Case "/" : Return "font" + txt.Text + "_" + "s1"
            Case "\" : Return "font" + txt.Text + "_" + "s2"
            Case "*" : Return "font" + txt.Text + "_" + "st"
            Case "|" : Return "font" + txt.Text + "_" + "pip"
            Case "," : Return "font" + txt.Text + "_" + "comma"
            Case "=" : Return "font" + txt.Text + "_" + "eq"
            Case "'" : Return "font" + txt.Text + "_" + "qo"
            Case "." : Return "font" + txt.Text + "_" + "dot"
            Case "A" To "Z" : Return "font" + txt.Text + "_#" & Str
            Case Else
                Return "font" + txt.Text + "_" + Str
        End Select
    End Function

    Sub CreateXData(ByVal kobj As AcadObject, ByVal Appname As String, ByVal InROut As String, ByVal InputID As String, ByVal OutputID As String, _
 ByVal StrtAng As String, ByVal MidAngle As String, ByVal frmName As String, ByVal WRad As String, ByVal ipTxt As String, ByVal TxtScale As String, _
 ByVal GrooveWd As String, ByVal SpaceFact As String, ByVal stAngChecked As String)
        Dim TypArray(0 To 12) As Int16
        Dim ValueArray(0 To 12) As Object
        TypArray(0) = 1001 : ValueArray(0) = Appname
        TypArray(1) = 1000 : ValueArray(1) = InROut
        TypArray(2) = 1000 : ValueArray(2) = InputID
        TypArray(3) = 1000 : ValueArray(3) = OutputID
        TypArray(4) = 1000 : ValueArray(4) = StrtAng
        TypArray(5) = 1000 : ValueArray(5) = frmName
        TypArray(6) = 1000 : ValueArray(6) = MidAngle
        TypArray(7) = 1000 : ValueArray(7) = WRad
        TypArray(8) = 1000 : ValueArray(8) = ipTxt
        TypArray(9) = 1000 : ValueArray(9) = TxtScale
        TypArray(10) = 1000 : ValueArray(10) = GrooveWd
        TypArray(11) = 1000 : ValueArray(11) = SpaceFact
        TypArray(12) = 1000 : ValueArray(12) = stAngChecked
        kobj.SetXData(TypArray, ValueArray)
    End Sub

    Sub CreateXDataoffset(ByVal kobj As AcadObject, ByVal Appname As String, ByVal InROut As String, ByVal InputID As String, ByVal OutputID As String, _
ByVal StrtAng As String, ByVal MidAngle As String, ByVal frmName As String, ByVal WRad As String, ByVal ipTxt As String, ByVal TxtScale As String, _
ByVal GrooveWd As String, ByVal SpaceFact As String, ByVal stAngChecked As String, ByVal datagridtext As String)
        Dim TypArray(0 To 13) As Int16
        Dim ValueArray(0 To 13) As Object
        TypArray(0) = 1001 : ValueArray(0) = Appname
        TypArray(1) = 1000 : ValueArray(1) = InROut
        TypArray(2) = 1000 : ValueArray(2) = InputID
        TypArray(3) = 1000 : ValueArray(3) = OutputID
        TypArray(4) = 1000 : ValueArray(4) = StrtAng
        TypArray(5) = 1000 : ValueArray(5) = frmName
        TypArray(6) = 1000 : ValueArray(6) = MidAngle
        TypArray(7) = 1000 : ValueArray(7) = WRad
        TypArray(8) = 1000 : ValueArray(8) = ipTxt
        TypArray(9) = 1000 : ValueArray(9) = TxtScale
        TypArray(10) = 1000 : ValueArray(10) = GrooveWd
        TypArray(11) = 1000 : ValueArray(11) = SpaceFact
        TypArray(12) = 1000 : ValueArray(12) = stAngChecked
        TypArray(13) = 1000 : ValueArray(13) = datagridtext
        kobj.SetXData(TypArray, ValueArray)
    End Sub

    Function PointAtAngle(ByVal kDist As Double, ByVal kAng As Double, ByVal sr As Double) As Object
        Dim kPt(0 To 2) As Double
        kPt(0) = kDist * Math.Cos(Ang2Rad(kAng))
        ' If kPt(0 Then
        kPt(1) = kDist * Math.Sin(Ang2Rad(kAng)) + sr
        kPt(2) = 0
        PointAtAngle = kPt
    End Function

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

End Module
