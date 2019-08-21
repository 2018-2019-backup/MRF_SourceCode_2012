Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Drawing
Imports System.Data
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.ApplicationServices

Public Class InsertBlock

    Public presskey As String
    Private frmExpanded As Boolean

    Dim yes As Boolean
    Dim needed As Integer
    Dim col As Integer
    Dim ro As Integer
    Dim key As Char
    Dim AcadApp As AcadApplication
    Dim ThisDrawing As AcadDocument
    Dim kColl As New Collection
    Dim j, k, m, lk, n As Integer
    Dim i As Double
    Dim IntPts, kPnt, intpts2 As Object
    Dim explodedObjects As Object
    Dim StAng, EnAng, AngStep, NumLines, OutRad As Double
    Dim RSt(0 To 2), REn(0 To 2) As Double
    Dim kRay As AcadRay
    Dim kLine As AcadLine
    Dim kselect As AcadSelectionSet
    Dim blockrefobj As AcadBlockReference
    Dim dirvector As Object
    Dim insertblock As String
    Dim insertionPnt(0 To 2) As Double
    Dim radrotation As Double
    Dim angle, rotation, vectorx, vectory As Double
    Dim circenterPoint(0 To 2) As Double
    Dim cirradius, temp1 As Double
    Dim typeflag As Boolean = False
    Dim previousText As String


    Private Sub InsertBlock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Size = New Size(269, 664)
        Me.Panel3.Size = New Size(245, 620)
        If CheckBox1.Checked = False Then
            Label6.Text = ""
            lbl_TotTextHight.Text = Val(TxtHt.Text) + Val(GrvWd.Text)
            TxtText.Text = ""
        End If
        Dim AcadApp As Autodesk.AutoCAD.Interop.AcadApplication : Dim AcadDoc As Autodesk.AutoCAD.Interop.AcadDocument
        AcadApp = GetObject(, "AutoCAD.Application")
        AcadDoc = AcadApp.ActiveDocument
        cmb_Font.Items.Clear()
        For Each B As AcadBlock In AcadDoc.Blocks
            If B.Name.Contains("_") And B.Name.ToLower.Contains("font") Then
                If Not cmb_Font.Items.Contains(B.Name.Split("_")(1).ToString.Replace("font", "")) Then
                    cmb_Font.Items.Add(B.Name.Split("_")(1).ToString.Replace("font", ""))
                End If
            End If
        Next
        If cmb_Font.Items.Count <> 0 Then
            cmb_Font.SelectedIndex = 0
        End If
        txtDist_TextChanged(sender, e)
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        DataGridView1.BeginEdit(True)
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        DataGridView1.BeginEdit(False)
    End Sub

    Private Sub DataGridView1_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles DataGridView1.CellValidating
        Dim cell As DataGridViewCell = DataGridView1.Item(e.ColumnIndex, e.RowIndex)
        col = e.ColumnIndex
        ro = e.RowIndex
        'Dim d As System.Windows.Forms.DataGridViewCellEventArgs


        If cell.IsInEditMode Then
            Dim c As Control = DataGridView1.EditingControl



            Select Case DataGridView1.Columns(e.ColumnIndex).Name
                Case "Gap"
                    '  If c.Text.Length = 1 Then
                    c.Text = CleanInputNumber(c.Text)
                    ' Else
                    ' c.Text = ""
                    ' End If

                    If c.Text = "" Then cell.Selected = True

                Case "Text"

                    If c.Text.Length = 1 Then
                        c.Text = CleanInputAlphabet(c.Text)
                    Else
                        c.Text = ""
                        'DataGridView1.CurrentCell = DataGridView1.Item(e.ColumnIndex, e.RowIndex)

                        'DataGridView1.BeginEdit(True)
                        'DataGridView1_CellClick(sender, e)
                        DataGridView1.Item(e.ColumnIndex, e.RowIndex).DataGridView.Focus()


                        ' MsgBox("Please enter single Character", MsgBoxStyle.Critical)



                        'cell.Selected = True
                    End If
                    'If c.Text = "" Then MsgBox("enter valied Data") 'Me.DataGridView1.EditMode = DataGridViewEditMode.EditOnKeystroke
            End Select

            'Me.DataGridView1.EditMode = DataGridViewEditMode.EditOnEnter
            'DataGridView1.Rows(0).Cells(0).Selected = True

        End If
        'DataGridView1.Rows(0).Cells(0).Selected = True
    End Sub

    Private Function CleanInputNumber(ByVal str As String) As String
        Return System.Text.RegularExpressions.Regex.Replace(str, "[a-zA-Z\b\s]", "")
    End Function

    Private Function CleanInputAlphabet(ByVal str As String) As String
        Return System.Text.RegularExpressions.Regex.Replace(str, "[0-9\b\s-]", "")
    End Function

    Sub insert()
        Try


            Dim AcadApp As Autodesk.AutoCAD.Interop.AcadApplication : Dim AcadDoc As Autodesk.AutoCAD.Interop.AcadDocument
            Dim hi As Double = 0.0
            ' Dim oApp As AcadApplication = Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication

            AcadApp = GetObject(, "AutoCAD.Application")
            AcadDoc = AcadApp.ActiveDocument
            Dim count As Integer = 0

            'Code by SPARTHIBAN

            Dim sBlockKey As String = "_"


            Dim blockRefObj As AcadBlockReference

            Dim minpnt, maxpnt As Object


            Dim iptime, optime As String

            Dim InputText As String
            Dim instring As String

            Dim iTextLength As Integer = TxtText.Text.Length - 1

            ''Array
            Dim k(0 To iTextLength) As String
            Dim insertionPnt(0 To 2) As Double
            Dim insertionPnt1(0 To 2) As Double
            ''Vars
            Dim width2, width1, midangle, stangle1, autocenter As Double
            Dim PlaceDia As Double = 0.0
            Dim vRadius, StartAngle, kang, ang, kang1 As Double

            width2 = 0.0 : width1 = 0.0 : midangle = 0.0 : stangle1 = 0.0 : autocenter = 0.0

            Dim txtScale As Double = CDbl(TxtHt.Text)
            Dim WdFactor As Double = CDbl(SpFactor.Text)
            Dim GrWidth As Double = CDbl(GrvWd.Text)

            '            'counters
            Dim i As Integer = 0
            InputText = TxtText.Text
            vRadius = CDbl(Radius.Text)

            Dim nocount(0 To iTextLength - 1) As Double

            iptime = genID() : optime = genID()

            Dim spaceratio(0 To iTextLength + 1) As Double
            Dim instring1(0 To iTextLength) As String
            Dim TotalAngleSpaces As Double = Nothing
            Dim FirstTxtWid As Double = Nothing
            Dim kMax, kMin As Double
            Dim SPMax, SPMin, TotalLetterSpace, TotalSpace As Double
            Dim ref As Double = Nothing
            TotalLetterSpace = 0
            TotalSpace = 0
            kMax = GetMaxX(AcadDoc.Blocks.Item(sBlockKey & GetBlkname(k(0), cmb_Font)))
            kMin = GetMinX(AcadDoc.Blocks.Item(sBlockKey & GetBlkname(k(0), cmb_Font)))
            For Each ds As System.Windows.Forms.DataGridViewRow In DataGridView1.Rows
                SPMax = GetMaxX(AcadDoc.Blocks.Item(sBlockKey & GetBlkname(k(i), cmb_Font)))
                SPMin = GetMinX(AcadDoc.Blocks.Item(sBlockKey & GetBlkname(k(i), cmb_Font)))

                k(i) = ds.Cells(0).Value   'Mid(InputText, i, 1)

                instring1(i) = k(i)
                TotalLetterSpace = TotalLetterSpace + (((SPMax - SPMin) * txtScale)) '+ CDbl(txta.Text) * 2
                '                '************CODE TO GET SPACE RATIO FROM XL******************
                '  If i = 0 Then
                spaceratio(i) = ds.Cells(1).Value
                'Else
                '    spaceratio(i) = ds.Cells(1).Value + CDbl(txta.Text) * 2
                'End If
                '* WdFactor  'Gap
                ref = (spaceratio(i)) + ref
                If i > 1 Then
                    TotalAngleSpaces = (spaceratio(i)) + TotalAngleSpaces
                End If
                TotalSpace = (spaceratio(i)) + TotalSpace
                i = i + 1

            Next
            '            '*************CODE TO FIND THE WIDTH OF THE TEXT*********************

            If Not RadioButton2.Checked Then 'For Start angel
                width2 = (((kMax - kMin - 0.2) * txtScale))  ' + GrWidth) / 2 '(spaceratio(1) / 2) '+ ( ' *********** TO COMPENSATE FOR STARTING ANGLE*********
            Else
                width2 = 0
            End If
            vRadius = CDbl(Radius.Text) ' + (GrWidth / 2))
            For i = 0 To iTextLength 'InputText.Length
                instring = GetBlkname(k(i), cmb_Font) 'Modified by gad
                '                

                Try
                    '                    ' instring = k(i)                    
                    AcadDoc.Blocks.Item(sBlockKey & instring).Item(0).GetBoundingBox(minpnt, maxpnt)
                Catch ex As Exception
                    minpnt(0) = 0 : maxpnt(0) = 1
                End Try

                width1 = 0 'spaceratio(i) '+ width1

            Next

            '            '************Autocenter**********************

            midangle = ((TotalSpace + TotalLetterSpace) * 360) / (2 * Math.PI * vRadius)

            stangle1 = midangle

            If Not RadioButton2.Checked Then
                StartAngle = CDbl(StAngle.Text) '+ Rad2Ang(-GrWidth / CDbl(Radius.Text))  '- Rad2Ang((GrWidth / 6) / vRadius) '- 0.15788459688762677 ' '+ (Rad2Ang(-GrWidth / vRadius) / 4) : autocenter = 0
                midangle = 0
            End If

            If Not RadioButton1.Checked Then
                autocenter = (EndAngle.Text) + (midangle / 2) '+ (Rad2Ang(((4 * TxtText.Text.Length * GrWidth)) / vRadius) / 2) + midangle : StartAngle = 0
                StartAngle = autocenter
            End If

            For i = 0 To iTextLength   'InputText.Length
                instring = GetBlkname(k(i), cmb_Font)
                If instring = "sp" Then
                    If i = 0 Then
                        kang = StartAngle - (Rad2Ang(CDbl(spaceratio(i))) / vRadius)
                    Else
                        kang = kang - (Rad2Ang(CDbl(spaceratio(i))) / vRadius)
                    End If

                    GoTo a
                End If
                Select Case StartAngle

                    Case 0

                        If kang <= -(355 - autocenter) Then MsgBox("Text Overlapping", MsgBoxStyle.SystemModal) : Exit Sub

                    Case Else

                        If kang <= -(355 - StartAngle) Then MsgBox("Text Overlapping", MsgBoxStyle.SystemModal) : Exit Sub

                End Select

                insertionPnt(0) = 0 : insertionPnt(1) = 0 : insertionPnt(2) = 0
                If i = 0 Then
                    kang = StartAngle - (Rad2Ang(CDbl(spaceratio(i))) / vRadius)
                    ang = StartAngle
                    hi = 0
                Else

                    kang = kang - (Rad2Ang(CDbl((spaceratio(i) + CDbl(Math.Round(hi, 2)) - (((Val(TxtHt.Text) * 0.2) - (CDbl(GrWidth.ToString))))) / vRadius))) ' - Rad2Ang(CDbl(Math.Round(hi, 2)) / vRadius)))


                End If

                insertionPnt = PointAtAngle(vRadius, kang)
                'insertionPnt(0) = insertionPnt(0) - (CDbl(GrWidth.ToString) / 2)
                'insertionPnt(1) = insertionPnt(1) - (CDbl(GrWidth.ToString) / 2)

                If instring = "" Then Continue For

                Dim atable As AcadTable


                Try
                    Dim obj As Object = AcadDoc.ModelSpace.Count
                    insertionPnt(0) = 0
                    insertionPnt(1) = 0
                    insertionPnt(2) = 0

                    Dim dou As Double = 0.3

                    blockRefObj = AcadDoc.ModelSpace.InsertBlock(insertionPnt, sBlockKey & instring, txtScale, txtScale, txtScale, 0)


                    blockRefObj.GetBoundingBox(minpnt, maxpnt)

                    kang1 = Rad2Ang(((-minpnt(0) + maxpnt(0)) - ((Val(TxtHt.Text) * 0.2) - CDbl(GrWidth.ToString))) / 2) / vRadius

                    insertionPnt1 = PointAtAngle((vRadius - (CDbl((Val(TxtHt.Text) * 0.2)) - CDbl(GrWidth.ToString))), (kang - kang1))

                    blockRefObj.Delete()

                    If i = 0 Then
                        Dim insertionTABLEPnt(0 To 2) As Double
                        insertionTABLEPnt(0) = insertionPnt1(0) + CDbl(txta.Text)
                        insertionTABLEPnt(1) = insertionPnt1(1)
                        insertionTABLEPnt(2) = insertionPnt1(2)
                        atable = AcadDoc.ModelSpace.AddTable(insertionTABLEPnt, 6, 2, 0.5, 1)
                        atable.SetCellValueFromText(0, 0, "Details", AcParseOption.acPreserveMtextFormat)
                        atable.SetCellValueFromText(1, 0, "H", AcParseOption.acPreserveMtextFormat)
                        atable.SetCellValueFromText(1, 1, TxtHt.Text, AcParseOption.acPreserveMtextFormat)
                        atable.SetCellValueFromText(2, 0, "W", AcParseOption.acPreserveMtextFormat)
                        atable.SetCellValueFromText(2, 1, GrWidth.ToString, AcParseOption.acPreserveMtextFormat)
                        atable.SetCellValueFromText(3, 0, "D", AcParseOption.acPreserveMtextFormat)
                        atable.SetCellValueFromText(3, 1, txtDist.Text, AcParseOption.acPreserveMtextFormat)
                        atable.SetCellValueFromText(4, 0, "A", AcParseOption.acPreserveMtextFormat)
                        atable.SetCellValueFromText(4, 1, txt_Angle.Text.ToString + "°", AcParseOption.acPreserveMtextFormat)
                        atable.SetCellValueFromText(5, 0, TXT_L_Angle.Text + "° " + Label15.Text, AcParseOption.acPreserveMtextFormat)

                        atable.TitleSuppressed = True
                        atable.HeaderSuppressed = True

                        atable.MergeCells(5, 5, 0, 1)
                    End If


                    blockRefObj = AcadDoc.ModelSpace.InsertBlock(insertionPnt1, sBlockKey & instring, txtScale, txtScale, txtScale, 0)

                    'blockRefObj.Move(insertionPnt, insertionPnt1)
                    blockRefObj.Rotation = Ang2Rad((kang - kang1) + 270)

                Catch ex As Exception

                    MsgBox(sBlockKey & instring & " Block not found.", MsgBoxStyle.SystemModal) : Continue For

                End Try
                previous = blockRefObj.ObjectName.ToString
b:              If instring = "sp" Then blockRefObj.Delete() : GoTo a

                Dim BlkEntities As Object
                Dim GrvPoly As AcadObject
                Dim OffsetPoly As Object
                Dim NewOffsetPoly() As Object
                Dim gpoly() As AcadObject
                Dim outer As Boolean = True

                BlkEntities = blockRefObj.Explode
                blockRefObj.Delete()
                Dim xdata As String = ""
                For f As Integer = 0 To DataGridView1.Rows.Count - 1
                    xdata = xdata + "$" + DataGridView1.Rows(f).Cells(0).Value.ToString + "@" + DataGridView1.Rows(f).Cells(1).Value.ToString
                Next
                Dim j As Integer

                For j = 0 To BlkEntities.Length - 1


                    If BlkEntities(j).ObjectName = "AcDbPolyline" Or BlkEntities(j).ObjectName = "AcDb2dPolyline" Then

                        GrvPoly = BlkEntities(j)
                        CreateXDataoffset(GrvPoly, "MRF-KKM1", "Output", iptime, optime, StAngle.Text, EndAngle.Text, "ALPHA", vRadius, InputText, txtScale.ToString + "&" + cmb_Font.SelectedIndex.ToString + "&" + txt_Angle.Text + "&" + txtDist.Text + "&" + TXT_L_Angle.Text, GrWidth, WdFactor, RadioButton1.Checked, xdata)
                        Dim pointsFORLINE(0 To 100) As Double
                        If GrvPoly.layer = "blk" Then
                            NewOffsetPoly = DirectCast(GrvPoly, Object).Offset(txta.Text)
                            CreateXDataoffset(NewOffsetPoly(0), "MRF-KKM1", "Output", iptime, optime, StAngle.Text, EndAngle.Text, "ALPHA", vRadius, InputText, txtScale.ToString + "&" + cmb_Font.SelectedIndex.ToString + "&" + txt_Angle.Text + "&" + txtDist.Text + "&" + TXT_L_Angle.Text, GrWidth, WdFactor, RadioButton1.Checked, xdata)
                            CreateXDataoffset(atable, "MRF-KKM1", "Output", iptime, optime, StAngle.Text, EndAngle.Text, "ALPHA", vRadius, InputText, txtScale.ToString + "&" + cmb_Font.SelectedIndex.ToString + "&" + txt_Angle.Text + "&" + txtDist.Text + "&" + TXT_L_Angle.Text, GrWidth, WdFactor, RadioButton1.Checked, xdata)

                            DirectCast(NewOffsetPoly(0), Object).color = Autodesk.AutoCAD.Interop.Common.ACAD_COLOR.acRed
                        End If

                        If (GrWidth - (0.2 * txtScale)) <> 0 Then
                            Try
                                If GrvPoly.layer = "blk" Then
                                    OffsetPoly = GrvPoly.offset((GrWidth - (0.2 * txtScale)) / 2)      'OffsetPoly = GrvPoly.Offset(0.1)
                                    NewOffsetPoly = GrvPoly.offset((GrWidth - (0.4 * txtScale)) / 2)
                                Else
                                    OffsetPoly = GrvPoly.offset((GrWidth - (0.2 * txtScale)) / 2) '   OffsetPoly = GrvPoly.Offset(0.1) '(GrWidth - (0.2 * txtScale)) / 2)
                                    NewOffsetPoly = GrvPoly.offset((GrWidth - (0.4 * txtScale)) / 2)
                                End If
                                CreateXDataoffset(OffsetPoly(0), "MRF-KKM1", "Output", iptime, optime, StAngle.Text, EndAngle.Text, "ALPHA", vRadius, InputText, txtScale.ToString + "&" + cmb_Font.SelectedIndex.ToString + "&" + txt_Angle.Text + "&" + txtDist.Text + "&" + TXT_L_Angle.Text, GrWidth, WdFactor, RadioButton1.Checked, xdata)
                                'CreateXDataoffset(NewOffsetPoly(0), "MRF-KKM1", "Output", iptime, optime, StAngle.Text, EndAngle.Text, "ALPHA", vRadius, InputText, txtScale.ToString + "&" + cmb_Font.SelectedIndex.ToString, GrWidth, WdFactor, RadioButton1.Checked, xdata)

                                GrvPoly.Delete()
                            Catch ex As Exception
                                MessageBox.Show("Width too high, retry with lower value", "Engrave Text")
                                Exit Sub
                            End Try
                        End If
                        hi = (maxpnt(0) - minpnt(0)) '+ CDbl(txta.Text) * 2



                    Else

                        CreateXDataoffset(BlkEntities(j), "MRF-KKM1", "Output", iptime, optime, StAngle.Text, EndAngle.Text, "ALPHA", vRadius, InputText, txtScale.ToString + "&" + cmb_Font.SelectedIndex.ToString + "&" + txt_Angle.Text + "&" + txtDist.Text + "&" + TXT_L_Angle.Text, GrWidth, WdFactor, RadioButton1.Checked, xdata)

                    End If

                Next


a:              width2 = width2 + spaceratio(i + 1)

            Next
            AcadApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
            ' acad = AcadApp.ActiveDocument
        Catch ex As Exception
            MsgBox("No Blocks Available")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        insert()
        Me.Close()
        Me.Dispose()


    End Sub

    Public Sub TxtText_KeyUp()
        If currenttext = "" Then
            If TxtText.Text.Length = 1 Or TxtText.Text.Length = 0 Then
                DataGridView1.Rows.Clear()
                DataGridView1.Rows.Add(presskey.ToString, 0)


                GoTo sx
            ElseIf TxtText.Text.Length = 2 Then
                DataGridView1.Rows.Clear()
                DataGridView1.Rows.Add(TxtText.Text.Chars(0).ToString, 0)
                DataGridView1.Rows.Add(TxtText.Text.Chars(TxtText.Text.Length - 1).ToString, (SpFactor.Text * TxtHt.Text).ToString)
            Else
                Dim str1 As String = ""
                Dim str(), str2() As String

                For l As Integer = 0 To DataGridView1.Rows.Count - 1
                    str1 = str1 + "$" + DataGridView1.Rows(l).Cells(0).Value.ToString + "@" + DataGridView1.Rows(l).Cells(1).Value.ToString
                Next
                DataGridView1.Rows.Clear()
                Try
                    str = str1.Split("$")
                    For i As Integer = 1 To TxtText.Text.Length - 1
                        str2 = str(i).Split("@")
                        DataGridView1.Rows.Add(str2(0).ToString, str2(1).ToString)
                    Next
                    DataGridView1.Rows.Add(TxtText.Text.Chars(TxtText.Text.Length - 1).ToString, (SpFactor.Text * TxtHt.Text).ToString)
                Catch ex As Exception
                    MsgBox("Type slowly")
                End Try
sx:
                previousText = TxtText.Text
            End If
        Else

            DataGridView1.Rows.Clear()
            Dim a(), b() As String
            a = currenttext.ToString.Split("$")
            For i As Integer = 1 To a.Length - 1
                b = a(i).Split("@")
                If frmInsertBlock.DataGridView1.ColumnCount = 0 Then frmInsertBlock.DataGridView1.Columns.Add("Text", "Text") : frmInsertBlock.DataGridView1.Columns.Add("Gap", "Gap")

                frmInsertBlock.DataGridView1.Rows.Add(b(0).ToString, b(1).ToString)
            Next
            previousText = TxtText.Text
        End If

    End Sub

    Private Sub SpFactor_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles SpFactor.TextChanged

        If SpFactor.Text = "" Then
            SpFactor.Text = 0
        ElseIf SpFactor.Text = "." Then
            GoTo e

        End If

        For i As Integer = 1 To TxtText.Text.Length - 1
            DataGridView1.Rows(i).Cells(1).Value = (SpFactor.Text * TxtHt.Text).ToString

        Next
e:      SpFactor.Text = SpFactor.Text
    End Sub

    Private Sub GrvWd_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        lbl_TotTextHight.Text = Val(TxtHt.Text) + Val(GrvWd.Text)
    End Sub

    Private Sub TxtHt_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtHt.TextChanged
        If currenttext = "" Then
            If TxtHt.Text = "" Then
                TxtHt.Text = 0
            End If
            'DataGridView1.Rows.Clear()
            'If TxtText.Text.Length = 0 Then
            '    DataGridView1.Rows.Clear()
            'Else
            '    DataGridView1.Rows.Add(TxtText.Text.Chars(0).ToString, 0)
            '    For i As Integer = 1 To TxtText.Text.Length - 1
            '        DataGridView1.Rows.Add(TxtText.Text.Chars(i).ToString, (SpFactor.Text * TxtHt.Text).ToString)

            '    Next
            'End If
            'For i As Integer = 1 To TxtText.Text.Length - 1
            '    DataGridView1.Rows(i).Cells(1).Value = (SpFactor.Text * TxtHt.Text).ToString

            'Next
            'GrvWd.Text = Val(TxtHt.Text) * 0.2
            Dim GrAngle As Double = Val(txt_Angle.Text)
            Dim GrWidth As Double = Val(TxtHt.Text) * 0.2
            Dim GrDepth As Double = Val(txtDist.Text)
            Try
                txta.Text = Math.Round(GrDepth * Math.Tan(Ang2Rad(GrAngle)), 3)
                ' GrWidth = Math.Tan(Ang2Rad(GrAngle / 2)) * GrDepth * 2
                txtWidth.Text = Math.Round(GrWidth, 3)
                GrvWd.Text = txtWidth.Text

                txtw2a.Text = Math.Round(GrWidth + (2 * Val(txta.Text)), 3).ToString
            Catch ex As Exception
            End Try
            lbl_TotTextHight.Text = Val(TxtHt.Text) + Val(GrvWd.Text)
        Else
            'TxtText_KeyUp()
            currenttext = ""
        End If

    End Sub

    Function genID() As String
        genID = DateTime.Now.Day & "-" & DateTime.Now.Month & "-" & DateTime.Now.Year & "-" & DateTime.Now.Hour & "-" & DateTime.Now.Minute & "-" & DateTime.Now.Second
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If frmExpanded Then
            Me.Width = 260
            Button2.Text = ">>"
        Else
            Me.Width = 520
            Button2.Text = "<<"
        End If
        frmExpanded = Not frmExpanded
    End Sub

    Private Sub SendVal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SendVal.Click
        GrvWd.Text = Val(GrW.Text)
    End Sub

    Public Sub OffsetObject(ByVal acDoc As AcadDocument)
        '' Get the current document and database
        Dim acCurDb As Database = acDoc.Database
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, _
            OpenMode.ForRead)
            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), _
            OpenMode.ForWrite)
            '' Create a lightweight polyline
            Dim acPoly As Polyline = New Polyline()
            acPoly.SetDatabaseDefaults()
            acPoly.AddVertexAt(0, New Point2d(1, 1), 0, 0, 0)
            acPoly.AddVertexAt(1, New Point2d(1, 2), 0, 0, 0)
            acPoly.AddVertexAt(2, New Point2d(2, 2), 0, 0, 0)
            acPoly.AddVertexAt(3, New Point2d(3, 2), 0, 0, 0)
            acPoly.AddVertexAt(4, New Point2d(4, 4), 0, 0, 0)
            acPoly.AddVertexAt(5, New Point2d(4, 1), 0, 0, 0)
            '' Add the new object to the block table record and the transaction
            acBlkTblRec.AppendEntity(acPoly)
            acTrans.AddNewlyCreatedDBObject(acPoly, True)
            '' Offset the polyline a given distance
            Dim acDbObjColl As DBObjectCollection = acPoly.GetOffsetCurves(0.25)
            '' Step through the new objects created
            For Each acEnt As Entity In acDbObjColl
                '' Add each offset object
                acBlkTblRec.AppendEntity(acEnt)
                acTrans.AddNewlyCreatedDBObject(acEnt, True)
            Next
            '' Save the new objects to the database
            acTrans.Commit()
        End Using
    End Sub

    Public Function XYZDistance(ByVal Point1 As Object, ByVal Point2 As Object) As Double
        'Returns the distance between two points
        Dim dblDist As Double
        Dim dblXSl As Double
        Dim dblYSl As Double
        Dim dblZSl As Double
        'Calc distance
        dblXSl = (Point1(0) - Point2(0)) ^ 2
        dblYSl = (Point1(1) - Point2(1)) ^ 2
        dblZSl = (Point1(2) - Point2(2)) ^ 2
        dblDist = Math.Sqrt(dblXSl + dblYSl + dblZSl)
        'Return Distance
        XYZDistance = dblDist
    End Function

    Private Sub BreakForBlock(ByVal ln As AcadObject, ByVal fpt As Object, ByVal spt As Object)

        Dim strp1 As String, strp2 As String

        Dim strh As String
        strh = ln.Handle
        strp1 = Replace(CStr(fpt(0)), ",", ".") & "," & _
                Replace(CStr(fpt(1)), ",", ".") & "," & _
                Replace(CStr(fpt(2)), ",", ".")
        strp2 = Replace(CStr(spt(0)), ",", ".") & "," & _
                Replace(CStr(spt(1)), ",", ".") & "," & _
                Replace(CStr(spt(2)), ",", ".")
        ThisDrawing.SendCommand("_BREAK " & _
                                "(handent " & Chr(34) & strh & Chr(34) & ")" & _
                                vbCr & strp1 & vbCr & strp2 & vbCr)
Exit_Here:
        'ss.Clear()
        With ThisDrawing
            .Regen(AcRegenType.acActiveViewport)
        End With

Err_Control:
        If Err.Number <> 0 Then
            MsgBox(Err.Description)
            Resume Exit_Here
        End If
    End Sub

    Private Sub TxtText_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtText.KeyDown
        typeflag = True
        Typefunction()
        previousText = TxtText.Text
    End Sub

    Private Function Typefunction() As Boolean
        Dim h As Integer
        If TxtText.Text.Length = 0 Then
            DataGridView1.Rows.Clear()
            GoTo sx
        ElseIf TxtText.Text.Length = 1 Then
            DataGridView1.Rows.Clear()
            DataGridView1.Rows.Add(TxtText.Text.Chars(0).ToString, 0)


            GoTo sx
        ElseIf TxtText.Text.Length = 2 Then
            DataGridView1.Rows.Clear()
            DataGridView1.Rows.Add(TxtText.Text.Chars(0).ToString, 0)
            DataGridView1.Rows.Add(TxtText.Text.Chars(TxtText.Text.Length - 1).ToString, (SpFactor.Text * TxtHt.Text).ToString)
        Else
            Dim str1 As String = ""
            Dim str(), str2() As String

            For l As Integer = 0 To DataGridView1.Rows.Count - 1
                str1 = str1 + "$" + DataGridView1.Rows(l).Cells(0).Value.ToString + "@" + DataGridView1.Rows(l).Cells(1).Value.ToString
            Next
            DataGridView1.Rows.Clear()
            Try
                'DataGridView1.Rows.Add(TxtText.Text.Chars(0).ToString, 0)
                str = str1.Split("$")
                For i As Integer = 1 To TxtText.Text.Length
                    str2 = str(i).Split("@")
                    DataGridView1.Rows.Add(str2(0).ToString, str2(1).ToString)
                    h = i
                Next
                ' DataGridView1.Rows.Add(TxtText.Text.Chars(TxtText.Text.Length - 1).ToString, (SpFactor.Text * TxtHt.Text).ToString)
            Catch ex As Exception
                For j As Integer = h To TxtText.Text.Length - 1
                    DataGridView1.Rows.Add(TxtText.Text.Chars(TxtText.Text.Length - 1).ToString, (SpFactor.Text * TxtHt.Text).ToString)
                Next
            End Try

sx:


        End If
        Return True
    End Function

    Private Sub TxtText_KeyUp1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtText.KeyUp
        Dim i As Integer
        ' If e.KeyCode = Keys.Back Or e.KeyCode = Keys.Delete Then

        If previousText.Count - 1 = TxtText.Text.Count Then
            Try
                For i = 0 To TxtText.Text.Count - 1
                    If Not TxtText.Text.Chars(i) = previousText.Chars(i) Then
                        DataGridView1.Rows.RemoveAt(i)
                        Exit Sub
                    End If
                Next
            Catch ex As Exception

            End Try
        ElseIf previousText.Count = TxtText.Text.Count - 1 Then
            Try
                For i = 0 To previousText.Count - 1
                    If Not TxtText.Text.Chars(i) = previousText.Chars(i) Then
                        DataGridView1.Rows.Insert(i, TxtText.Text.Chars(i), (SpFactor.Text * TxtHt.Text).ToString)
                        Exit Sub
                    End If
                Next
            Catch ex As Exception

            End Try
        ElseIf previousText.Count > TxtText.Text.Count Then
            Try
                For i = 0 To TxtText.Text.Count - 1
                    If Not TxtText.Text.Chars(i) = previousText.Chars(i) Then
                        DataGridView1.Rows.RemoveAt(i)

                    End If
                Next
            Catch ex As Exception

            End Try
            'ElseIf previousText.Count < TxtText.Text.Count Then
            '    Try
            '        For i = 0 To previousText.Count - 1
            '            If Not TxtText.Text.Chars(i) = previousText.Chars(i) Then
            '                DataGridView1.Rows.Insert(i, TxtText.Text.Chars(i), 0)

            '            End If
            '        Next
            '    Catch ex As Exception

            '    End Try

        End If




        If i = previousText.Count Or i = TxtText.Text.Count Then
            Typefunction()
        End If


        If typeflag = False Then
            Typefunction()
        End If
    End Sub

    Private Sub TxtText_MouseCaptureChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtText.MouseCaptureChanged

    End Sub

    Private Sub TxtText_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtText.MouseClick

    End Sub

    Private Sub TxtText_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtText.MouseDown
        TxtText.SelectionLength = 0
        TxtText.SelectedText = ""

    End Sub

    Private Sub TxtText_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtText.MouseEnter
        TxtText.SelectionLength = 0
        TxtText.SelectedText = ""
    End Sub

    Private Sub TxtText_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtText.MouseMove
        TxtText.SelectionLength = 0
        TxtText.SelectedText = ""
    End Sub

    Private Sub TxtText_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtText.MouseUp
        TxtText.SelectionLength = 0
        TxtText.SelectedText = ""
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton2.CheckedChanged, RadioButton1.CheckedChanged
        If TryCast(sender, RadioButton).Name = "RadioButton2" Then
            EndAngle.Enabled = True
            StAngle.Enabled = False
            Label15.Text = "Auto Center"
        Else
            StAngle.Enabled = True
            EndAngle.Enabled = False
            Label15.Text = "Start Angle"
        End If
    End Sub

    Private Sub txt_Angle_Leave(sender As System.Object, e As System.EventArgs) Handles txtWidth.Leave, txtw2a.Leave, txta.Leave, TextBox1.Leave, TextBox2.Leave
        Dim GrAngle As Double = Val(txt_Angle.Text)
        Dim GrWidth As Double = Val(TxtHt.Text) * 0.2
        Dim GrDepth As Double = Val(txtDist.Text)
        Try
            txta.Text = Math.Round(GrDepth * Math.Tan(Ang2Rad(GrAngle)), 3)
            ' GrWidth = Math.Tan(Ang2Rad(GrAngle / 2)) * GrDepth * 2
            txtWidth.Text = Math.Round(GrWidth, 3)
            GrvWd.Text = txtWidth.Text

            txtw2a.Text = Math.Round(GrWidth + (2 * Val(txta.Text)), 3).ToString
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnextend_Click(sender As System.Object, e As System.EventArgs) Handles btnextend.Click
        Me.Size = New Size(611, 664)
        Me.Panel3.Size = New Size(587, 620)
    End Sub

    Private Sub btnsend_Click(sender As Object, e As System.EventArgs) Handles btnsend.Click
        Me.Size = New Size(269, 664)
        Me.Panel3.Size = New Size(245, 620)
    End Sub

    Private Sub txtDist_TextChanged(sender As System.Object, e As System.EventArgs)
        Dim GrAngle As Double = Val(txt_Angle.Text)
        Dim GrWidth As Double = Val(TxtHt.Text) * 0.2
        Dim GrDepth As Double = Val(txtDist.Text)
        Try
            txta.Text = Math.Round(GrDepth * Math.Tan(Ang2Rad(GrAngle)), 3)
            ' GrWidth = Math.Tan(Ang2Rad(GrAngle / 2)) * GrDepth * 2
            txtWidth.Text = Math.Round(GrWidth, 3)
            GrvWd.Text = txtWidth.Text

            txtw2a.Text = Math.Round(GrWidth + (2 * Val(txta.Text)), 3).ToString
        Catch ex As Exception
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtDist.TextChanged, TextBox1.TextChanged
        txtDist.Text = TryCast(sender, TextBox).Text
        TextBox1.Text = TryCast(sender, TextBox).Text

    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_Angle.TextChanged, TextBox2.TextChanged
        txt_Angle.Text = TryCast(sender, TextBox).Text
        TextBox2.Text = TryCast(sender, TextBox).Text
    End Sub

    Private Sub txt_Angle_KeyDown(sender As System.Object, e As System.Windows.Forms.KeyEventArgs) Handles txt_Angle.KeyDown, txtDist.KeyDown, TextBox1.KeyDown, TextBox2.KeyDown
        Dim GrAngle As Double = Val(txt_Angle.Text)
        Dim GrWidth As Double = Val(TxtHt.Text) * 0.2
        Dim GrDepth As Double = Val(txtDist.Text)
        Try
            txta.Text = Math.Round(GrDepth * Math.Tan(Ang2Rad(GrAngle)), 3)
            ' GrWidth = Math.Tan(Ang2Rad(GrAngle / 2)) * GrDepth * 2
            txtWidth.Text = Math.Round(GrWidth, 3)
            GrvWd.Text = txtWidth.Text

            txtw2a.Text = Math.Round(GrWidth + (2 * Val(txta.Text)), 3).ToString
        Catch ex As Exception
        End Try

    End Sub

    Private Sub GrvWd_TextChanged_1(sender As System.Object, e As System.EventArgs) Handles txtWidth.TextChanged, GrvWd.TextChanged
        GrvWd.Text = TryCast(sender, TextBox).Text
        txtWidth.Text = TryCast(sender, TextBox).Text
    End Sub
End Class