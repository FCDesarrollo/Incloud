Imports System.Data.SqlClient
Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports Microsoft.Office.Interop
Imports ZXing
Module modImprimeXML

    Private ReadyState As Boolean
    Private Const maxTime As Long = 20
    Public Const tFactura As String = "FACTURA"
    Public Const tPoliza As String = "POLIZA"
    Public Const directContaXMLNUBE As String = "Contabilidad/Contabilidad/ExpedientesContables"

    Public dCarpetas As Dictionary(Of String, String)
    Public dCarpetasPol As Dictionary(Of String, String)
    Public sConCarpetas As Boolean
    Private Enum iColEnc
        iTipo = 1
        iFecha = 2
        iSerie = 3
        iFolio = 4
        iUUID = 5
        iLExp = 8
        iVers = 9
        iERfc = 2
        iENom = 4
        iMon = 6
        iusocfdi = 9
        iFormaP = 7
        iMetodP = 9
        iRRfc = 2
        iRNom = 4
    End Enum

    Private Enum iColMov
        iClavePro = 1
        iIDSat = 2
        iCant = 3
        iUnid = 4
        iDes = 1
        iImpor = 6
        iIva = 7
        iIeps = 8
        iTotal = 9
    End Enum

    Public Sub ConsultaXML(ByVal xVersion As Integer, ByVal xcon As SqlConnection,
                           ByVal cconEmpr As SqlConnection,
                           ByVal fechai As Date, fechaf As Date,
                           sEmpresa As String, plantilla As String, Optional UUIDXml As String = "")
        Dim cQue As String, cMov As String, cConten As String, tXml As String
        Dim tamano As Integer, posicion As Integer, posicion1 As Integer, posicion2 As Integer, longitud As Integer
        Dim dIeps As Double, dIva As Double, rutaExiste As String
        Dim dXML As CLXml
        Dim dXmlMov As CLMovXml
        Dim lGuidDocument As Guid

        'rutaExiste = "C:\Users\Arturo Gallegos\Desktop\MODULOS\ARCHIVOXML\" & sEmpresa & "\"
        rutaExiste = FC_RutaModulos & "\ArchivosIncloud\" & sEmpresa & "\COMPROBANTES\"

        If Not System.IO.Directory.Exists(rutaExiste) Then
            Exit Sub
        End If

        If xVersion = 0 Then
            cQue = "SELECT RFCEmisor,NombreEmisor,RegimenEmisor, RFCReceptor, NombreReceptor, GuidDocument,
                      Version, Serie, Folio, Fecha, FormaPago, CondicionesPago, Subtotal, Descuento, TipoCambio, Moneda, Total, TipoComprobante, MetodoPago, 
                      LugarExp, UUID, FechaTimbrado, NumeroCertificado, TipoDocumento, UsoCFDI FROM Comprobante
                    WHERE TipoComprobante<>'P' AND Cast(Fecha As Date)>=@fecha and Cast(Fecha As Date)<=@fechaF"
        Else
            cQue = "SELECT RFCEmisor,NombreEmisor,RegimenEmisor, RFCReceptor, NombreReceptor, GuidDocument,
                          Version, Serie, Folio, Fecha, FormaPago, CondicionesPago, Subtotal, Descuento, TipoCambio, Moneda, Total, TipoComprobante, MetodoPago, 
                          LugarExp, UUID, FechaTimbrado, NumeroCertificado, TipoDocumento, UsoCFDI FROM Comprobante
                           WHERE TipoComprobante<>'P' AND Cast(Fecha As Date)>=@fecha AND Cast(Fecha As Date)<=@fechaF AND UUID=@uuid"
        End If

        Using comsr = New SqlCommand(cQue, xcon)
            comsr.Parameters.AddWithValue("@fecha", fechai)
            comsr.Parameters.AddWithValue("@fechaF", fechaf)
            If xVersion <> 0 Then
                comsr.Parameters.AddWithValue("@uuid", UUIDXml)
            End If
            Using Rscon = comsr.ExecuteReader()
                Do While Rscon.Read()
                    If Not System.IO.File.Exists(rutaExiste & Rscon("UUID").ToString & ".xlsx") Or xVersion = 0 Then
                        dXML = New CLXml
                        With dXML
                            .SVersion = Rscon("Version")

                            .SNombreEmisor = Trim(Rscon("NombreEmisor"))
                            .SRFCEmisor = Trim(Rscon("RFCEmisor"))
                            .SRegimenFiscalE = Trim(Rscon("RegimenEmisor"))

                            .SNombreReceptor = Trim(Rscon("NombreReceptor"))
                            .SRFCReceptor = Trim(Rscon("RFCReceptor"))

                            .SUsoCFDI = Rscon("UsoCFDI")
                            .SFecha = Format(CDate(Rscon("Fecha")), "dd/MM/yyyy")
                            .SFolio = IIf(Not IsNothing(Rscon("Folio")), Rscon("Folio"), 0)
                            .SSerie = IIf(Not IsNothing(Rscon("Serie")), Rscon("Serie"), 0)

                            .SSubtotal = Rscon("SubTotal")
                            .SDescto = IIf(Rscon("Descuento") IsNot DBNull.Value, Rscon("Descuento"), 0)
                            .STotalXML = Rscon("Total")

                            .STipo = Rscon("TipoComprobante")
                            .SFormaPago = Rscon("FormaPago")
                            .SMoneda = Rscon("Moneda")
                            .STipoCambio = Rscon("TipoCambio")
                            .SMetodoPago = Rscon("MetodoPago")
                            .SLugarExpedicion = Trim(Rscon("LugarExp"))
                            .SNoCertificado = Rscon("NumeroCertificado")

                            .SUUID = Rscon("UUID")
                            .SFechaTimbrado = Trim(Rscon("FechaTimbrado"))
                            .SCerSAT = Trim(Rscon("NumeroCertificado"))
                            lGuidDocument = Rscon("GuidDocument")

                            .STotalIva = GetSumImpuesto(xcon, "Impuesto_Traslado", "IVA", Rscon("UUID").ToString)
                            .STotalIeps = GetSumImpuesto(xcon, "Impuesto_Traslado", "IEPS", Rscon("UUID").ToString)

                            .STotalRetIsr = GetSumImpuesto(xcon, "Impuesto_Retencion", "ISR", Rscon("UUID").ToString)
                            .STotalRetIva = GetSumImpuesto(xcon, "Impuesto_Retencion", "IVA", Rscon("UUID").ToString)
                            'GetSumImpuesto

                            cConten = "SELECT Content FROM DocumentContent WHERE GuidDocument = '" & lGuidDocument.ToString & "'"
                            Using comConten = New SqlCommand(cConten, DConexionesConten(sEmpresa))
                                Using rsConten = comConten.ExecuteReader()
                                    rsConten.Read()
                                    If rsConten.HasRows = True Then
                                        tXml = rsConten("Content")
                                        tamano = Len("SelloCFD=") + 1
                                        posicion1 = InStr(1, tXml, "SelloCFD", vbTextCompare) + tamano
                                        posicion2 = InStr(posicion1, tXml, Chr(34), vbTextCompare)
                                        longitud = posicion2 - posicion1
                                        .SSelloDig = Mid(tXml, posicion1, longitud)

                                        tamano = Len("SelloSAT=") + 1
                                        posicion1 = InStr(1, tXml, "SelloSAT", vbTextCompare) + tamano
                                        posicion2 = InStr(posicion1, tXml, Chr(34), vbTextCompare)
                                        longitud = posicion2 - posicion1
                                        .SSelloSAT = Mid(tXml, posicion1, longitud)

                                        tamano = Len("Version=") + 1
                                        posicion = InStr(1, tXml, "Complemento", vbTextCompare)
                                        posicion1 = InStr(posicion, tXml, "Version", vbTextCompare) + tamano
                                        posicion2 = InStr(posicion1, tXml, Chr(34), vbTextCompare)
                                        longitud = posicion2 - posicion1
                                        .SVersionSello = Mid(tXml, posicion1, longitud)
                                    End If
                                End Using
                            End Using
                            .SCodigoQr = "?re=" & .SRFCEmisor & "&rr=" & .SRFCReceptor & "&tt=" & .STotalXML & "&id=" & .SUUID.ToString
                        End With
                        dIeps = 0
                        dIva = 0
                        cMov = "SELECT IdConcepto, Importe, ValorUnitario, Descripcion, 
                                    NoIdentificacion, Unidad, Cantidad, Descuento, CveProdSer FROM Conceptos 
                                        WHERE GuidDocument=@GuiDoc"
                        Using comMov = New SqlCommand(cMov, xcon)
                            comMov.Parameters.AddWithValue("@GuiDoc", Rscon("GuidDocument"))
                            Using rsMov = comMov.ExecuteReader()
                                Do While rsMov.Read()
                                    dIva = 0
                                    dIeps = 0
                                    dXmlMov = New CLMovXml
                                    With dXmlMov
                                        .MImporte = rsMov("Importe")
                                        .MValorUni = rsMov("ValorUnitario")
                                        .MDescrip = Trim(rsMov("Descripcion"))
                                        .MNoIentifi = Trim(rsMov("NoIdentificacion"))
                                        .MUnidad = Trim(rsMov("Unidad"))
                                        .MCantidad = rsMov("Cantidad")
                                        .MDesc = IIf(rsMov("Descuento") Is DBNull.Value, 0, rsMov("Descuento"))
                                        .MCveProdSer = rsMov("CveProdSer")
                                        cConten = "SELECT ImpuestoDesc,Importe FROM Impuesto_Traslado_Concepto 
                                                WHERE IdConcepto=@idcon"
                                        Using comTras = New SqlCommand(cConten, xcon)
                                            comTras.Parameters.AddWithValue("@idcon", rsMov("IdConcepto"))
                                            Using rsTras = comTras.ExecuteReader()
                                                Do While rsTras.Read()
                                                    If rsTras("ImpuestoDesc") = "IVA" Then
                                                        dIva = dIva + rsTras("Importe")
                                                    ElseIf rsTras("ImpuestoDesc") = "IEPS" Then
                                                        dIeps = dIeps + rsTras("Importe")
                                                    End If
                                                Loop
                                            End Using
                                        End Using
                                        .MIva = dIva
                                        .MIeps = dIeps
                                    End With
                                    dXML.MovXml.Add(dXmlMov)
                                Loop
                                rsMov.Close()
                            End Using
                        End Using
                        AddXML(dXML)
                    End If
                Loop
                Rscon.Close()
            End Using
        End Using
        CrearArchivosXML(cconEmpr, sEmpresa, plantilla)
    End Sub

    Public Sub AddXML(ByRef nuevoXml As CLXml)
        Dim exists As Boolean
        Dim i As Integer

        For i = cgXML.Count To 1 Step -1
            If cgXML.Item(i).SUUID = Nothing Then
                cgXML.Remove(i)
            ElseIf cgXML.Item(i).SUUID = nuevoXml.SUUID Then
                exists = True
                Exit For
            End If
        Next i

        If Not exists Then cgXML.Add(nuevoXml)
    End Sub

    Private Sub CrearArchivosXML(ByVal conE As SqlConnection,
                                 ByVal cEmpresa As String, ByVal plantilla As String)
        Dim nXml As CLXml, sExiste As Boolean
        dCarpetas = New Dictionary(Of String, String)
        sExiste = False

        For Each nXml In cgXML
            sExiste = True
            If Creafactura(nXml, cEmpresa, plantilla) = True Then

            End If

        Next nXml
        'If sExiste = True Then
        '    Anexalink(cEmpresa)
        'End If
    End Sub


    Public Sub Anexalink(ByVal aemp As String)
        Dim appExcel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wbExcel As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim wsExcel As Microsoft.Office.Interop.Excel.Worksheet = Nothing

        Dim rutalibro As String, linkstring As String
        Dim clasebit As clBitacora
        Dim regbit As clRegistroBitacora
        Dim dDatos() As String
        Dim x As Long, aUUID As String, strmes As String, fecha As Date
        Dim mes As Integer
        appExcel = New Microsoft.Office.Interop.Excel.Application
        appExcel.Visible = False
        appExcel.DisplayAlerts = False
        Try
            clasebit = New clBitacora
            clasebit.Idsubmenu = 4
            clasebit.Tipodocumento = "COMPROBANTES"
            For t = 0 To dCarpetas.Count - 1
                rutalibro = FC_RutaModulos & "\ArchivosIncloud\" &
            aemp & "\COMPROBANTES\" & dCarpetas.Item(dCarpetas.Keys(t)) &
            "\pdfs\relacion\" & dCarpetas.Keys(t) & ".xlsx"

                wbExcel = appExcel.Workbooks.Open(rutalibro)
                wsExcel = wbExcel.ActiveSheet

                With wsExcel
                    dDatos = Split(dCarpetas.Item(dCarpetas.Keys(t)), "\")
                    .Cells(2, 1).value = .Cells(2, 1).value & " de " & UCase(MonthName(CInt(dDatos(1)))) & " del 20" & dDatos(0)
                    '.Cells(3, 2).value = CInt(dDatos(0))
                    '.Cells(3, 3).value = "Periodo"
                    '.Cells(3, 4).value = Format(dDatos(1), "00")
                    x = 5
                    Do While .Cells(x, 1).value <> ""
                        If .Cells(x, 8).value = "" And .Cells(x, 9).value <> "" Then
                            linkstring = getLinkCompartido(.Cells(x, 9).value)
                            If linkstring <> "" Then
                                .Cells(x, 8).value = linkstring
                                aUUID = .Cells(x, 7).value
                                .Hyperlinks.Add(Anchor:= .Range("G" & CStr(x)),
                                                        Address:=linkstring,
                                                        TextToDisplay:=aUUID)
                            End If
                        End If
                        x += 1
                    Loop
                    .Columns("B:H").AutoFit
                    .Columns("H:I").Hidden = True
                    .Columns("A:I").Sort(key1:= .Range("A4"),
      order1:=Excel.XlSortOrder.xlDescending, header:=Excel.XlYesNoGuess.xlYes)
                End With

                wbExcel.Save()

                appExcel.ScreenUpdating = False
                appExcel.DisplayAlerts = False

                wbExcel.ExportAsFixedFormat(Type:=Excel.XlFixedFormatType.xlTypePDF,
                                             Filename:=FC_RutaModulos & "\ArchivosIncloud\" &
            aemp & "\COMPROBANTES\" & dCarpetas.Item(dCarpetas.Keys(t)) &
            "\pdfs\relacion\" & dCarpetas.Keys(t) & "_CFDI",
                                             Quality:=Excel.XlFixedFormatQuality.xlQualityStandard,
                                             IncludeDocProperties:=True, IgnorePrintAreas:=False,
                                             OpenAfterPublish:=False)

                'dDatos = Split(dCarpetas.Item(dCarpetas.Keys(t)), "\")
                regbit = New clRegistroBitacora
                strmes = dDatos(1)
                fecha = CDate("01/" & strmes & "/" & CStr(Year(Date.Now.Date)))
                mes = Month(fecha)
                regbit.Periodo = mes
                regbit.Ejercicio = CInt(dDatos(0))
                regbit.Archivo = dCarpetas.Keys(t) & "_CFDI.pdf"
                regbit.Nombrearchivo = GlobalRFCEmpresa & "_" & Right(dDatos(0), 2) & Format(mes, "00") & "_CFDI"
                clasebit.Regbitacora.Add(regbit)


                wsExcel = Nothing
                wbExcel.Close()
                wbExcel = Nothing
                appExcel.Quit()
            Next
            clasebit.AgregaRegistro()
            regbit = Nothing
            clasebit = Nothing
        Catch ex As Exception
            My.Computer.FileSystem.WriteAllText(FC_RutaModulos & "\ArchivosIncloud\" & aemp & "\COMPROBANTES\errores.log", Format(Now, "dd/MM/yyy HH:mm") & " - " & ex.Message & vbCrLf, True)
        Finally
            releaseObject(wsExcel)
            releaseObject(wbExcel)
            releaseObject(appExcel)
        End Try


        dCarpetas = Nothing
    End Sub

    Public Sub AnexalinkPol(ByVal aemp As String)
        Dim appExcel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wbExcel As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim wsExcel As Microsoft.Office.Interop.Excel.Worksheet = Nothing

        Dim rutalibro As String, linkstring As String
        Dim clasebit As clBitacora
        Dim regbit As clRegistroBitacora
        Dim dDatos() As String
        Dim x As Long, aUUID As String, strmes As String, fecha As Date
        Dim mes As Integer
        appExcel = New Microsoft.Office.Interop.Excel.Application
        appExcel.Visible = False
        appExcel.DisplayAlerts = False
        Try
            clasebit = New clBitacora
            clasebit.Idsubmenu = 4
            clasebit.Tipodocumento = "POLIZAS"
            For t = 0 To dCarpetasPol.Count - 1
                rutalibro = FC_RutaModulos & "\ArchivosIncloud\" &
            aemp & "\POLIZAS\" & dCarpetasPol.Item(dCarpetasPol.Keys(t)) &
            "\relacion\" & dCarpetasPol.Keys(t) & ".xlsx"

                wbExcel = appExcel.Workbooks.Open(rutalibro)
                wsExcel = wbExcel.ActiveSheet

                With wsExcel
                    dDatos = Split(dCarpetasPol.Item(dCarpetasPol.Keys(t)), "\")
                    .Cells(3, 1).value = .Cells(2, 1).value & " de " & UCase(MonthName(CInt(dDatos(1)))) & " del 20" & dDatos(0)
                    '.Cells(3, 2).value = CInt(dDatos(0))
                    '.Cells(3, 3).value = "Periodo"
                    '.Cells(3, 4).value = dDatos(1)

                    x = 5
                    Do While .Cells(x, 1).value <> ""
                        If .Cells(x, 7).value = "" And .Cells(x, 8).value <> "" Then
                            linkstring = getLinkCompartido(.Cells(x, 8).value)
                            If linkstring <> "" Then
                                .Cells(x, 7).value = linkstring
                                aUUID = .Cells(x, 6).value
                                .Hyperlinks.Add(Anchor:= .Range("F" & CStr(x)),
                                                        Address:=linkstring,
                                                        TextToDisplay:=aUUID)
                            End If
                        End If
                        x += 1
                    Loop
                    .Columns("B:F").AutoFit
                    .Columns("G:I").Hidden = True
                    .Columns("A:I").Sort(key1:= .Range("A4"),
      order1:=Excel.XlSortOrder.xlDescending, header:=Excel.XlYesNoGuess.xlYes)
                End With

                wbExcel.Save()

                appExcel.ScreenUpdating = False
                appExcel.DisplayAlerts = False

                wbExcel.ExportAsFixedFormat(Type:=Excel.XlFixedFormatType.xlTypePDF,
                                             Filename:=FC_RutaModulos & "\ArchivosIncloud\" &
            aemp & "\POLIZAS\" & dCarpetasPol.Item(dCarpetasPol.Keys(t)) &
            "\relacion\" & dCarpetasPol.Keys(t) & "_POL",
                                             Quality:=Excel.XlFixedFormatQuality.xlQualityStandard,
                                             IncludeDocProperties:=True, IgnorePrintAreas:=False,
                                             OpenAfterPublish:=False)


                regbit = New clRegistroBitacora
                strmes = dDatos(1)
                fecha = CDate("01/" & strmes & "/" & CStr(Year(Date.Now.Date)))
                mes = Month(fecha)
                regbit.Periodo = mes
                regbit.Ejercicio = CInt(dDatos(0))
                regbit.Archivo = dCarpetasPol.Keys(t) & "_POL.pdf"
                regbit.Nombrearchivo = GlobalRFCEmpresa & "_" & Right(dDatos(0), 2) & Format(mes, "00") & "_POL"
                clasebit.Regbitacora.Add(regbit)


                wsExcel = Nothing
                wbExcel.Close()
                wbExcel = Nothing
                appExcel.Quit()
            Next
            clasebit.AgregaRegistro()
            regbit = Nothing
            clasebit = Nothing
        Catch ex As Exception
            My.Computer.FileSystem.WriteAllText(FC_RutaModulos & "\ArchivosIncloud\" & aemp & "\POLIZAS\errores.log", Format(Now, "dd/MM/yyy HH:mm") & " - " & ex.Message & vbCrLf, True)
        Finally
            releaseObject(wsExcel)
            releaseObject(wbExcel)
            releaseObject(appExcel)
        End Try


        dCarpetasPol = Nothing
    End Sub

    Public Function Creafactura(ByVal factu As CLXml,
                                ByVal cEmpresa As String, ByVal plantilla As String) As Boolean
        Dim m As CLMovXml
        Dim appXL As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wbXl As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim shXL As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim rutaQr As String
        Dim Celda As Object = Nothing
        Dim Izquierda As Single
        Dim Arriba As Single
        Dim Alto As Double
        Dim Fname As String = "", complementoruta As String, rutaGuarda As String
        Dim compleRelacion As String
        Creafactura = True
        Dim i As Integer = 11
        Try
            'Fname = "C:\Users\Arturo Gallegos\Desktop\MODULOS\plantillafactura.xlsx"
            Fname = FC_RutaModulos & "\ArchivosIncloud\" & cEmpresa & "\COMPROBANTES\" & plantilla

            appXL = New Microsoft.Office.Interop.Excel.Application
            appXL.Visible = False
            wbXl = appXL.Workbooks.Open(Fname)
            shXL = wbXl.ActiveSheet
            With shXL
                ''ENCABEZADO
                complementoruta = Right(Year(factu.SFecha), 2) & "\" & UCase(Format(Month(factu.SFecha), "00"))
                compleRelacion = Right(Year(factu.SFecha), 2) & UCase(Format(Month(factu.SFecha), "00"))
                .Cells(3, iColEnc.iTipo).value = factu.STipo
                .Cells(3, iColEnc.iFecha).value = "'" & factu.SFecha
                .Cells(3, iColEnc.iSerie).value = factu.SSerie
                .Cells(3, iColEnc.iFolio).value = factu.SFolio
                .Cells(3, iColEnc.iUUID).value = factu.SUUID.ToString
                .Cells(3, iColEnc.iLExp).value = factu.SLugarExpedicion
                .Cells(3, iColEnc.iVers).value = factu.SVersion

                .Cells(5, iColEnc.iERfc).value = factu.SRFCEmisor
                .Cells(5, iColEnc.iENom).value = factu.SNombreEmisor

                .Cells(6, iColEnc.iMon).value = factu.SMoneda
                .Cells(6, iColEnc.iFormaP).value = factu.SFormaPago
                .Cells(6, iColEnc.iMetodP).value = factu.SMetodoPago

                .Cells(8, iColEnc.iusocfdi).value = factu.SUsoCFDI

                .Cells(8, iColEnc.iRRfc).value = factu.SRFCReceptor
                .Cells(8, iColEnc.iRNom).value = factu.SNombreReceptor


                ''MOVIMIENTOS
                For Each m In factu.MovXml
                    .Cells(i, iColMov.iClavePro).value = m.MCveProdSer
                    .Cells(i, iColMov.iIDSat) = m.MNoIentifi
                    .Cells(i, iColMov.iCant).value = m.MCantidad
                    .Cells(i, iColMov.iUnid).value = m.MUnidad

                    .Cells(i, iColMov.iImpor).value = m.MImporte
                    .Cells(i, iColMov.iIva).value = m.MIva
                    .Cells(i, iColMov.iIeps) = m.MIeps
                    .Cells(i, iColMov.iTotal) = m.MImporte + m.MIva
                    i = i + 1
                    .Rows(i).Insert()
                    .Cells(i, iColMov.iDes).value = m.MDescrip
                    .Cells(i, iColMov.iDes).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                    i = i + 1
                    .Rows(i).Insert()
                Next m

                ''TOTALES
                .Cells(i + 3, 2).value = factu.SSubtotal
                .Cells(i + 3, 3).value = factu.SDescto
                .Cells(i + 3, 4).value = factu.STotalIva
                .Cells(i + 3, 5).value = factu.STotalIeps
                .Cells(i + 3, 6).value = factu.STotalRetIsr
                .Cells(i + 3, 7).value = factu.STotalRetIva
                .Cells(i + 3, 9).value = factu.STotalXML

                ''TOTAL EN LETRA
                If factu.SMoneda = "USD" Then
                    letr.MascaraSalidaDecimal = "00/100 USD"
                    letr.SeparadorDecimalSalida = "DOLARES"
                Else
                    letr.MascaraSalidaDecimal = "00/100 M.N."
                    letr.SeparadorDecimalSalida = "PESOS"
                End If
                .Cells(i + 4, 2).value = UCase(letr.ToCustomCardinal(factu.STotalXML))

                ''SELLOS
                .Cells(i + 7, 1).value = factu.SSelloDig
                .Cells(i + 10, 1).value = factu.SSelloSAT
                .Cells(i + 13, 3).value = Trim("||" &
                factu.SVersionSello & "|" & factu.SUUID.ToString & "|" &
                factu.SFechaTimbrado & "|" & factu.SSelloDig & "|" & factu.SCerSAT & "||")

                .Cells(i + 17, 3).value = "'" & factu.SCerSAT
                .Cells(i + 17, 6).value = factu.SFechaTimbrado


                ''CODIGO QR
                CrearCodigoQR(factu.SCodigoQr)
                Celda = .Cells(i + 17, 1)

                Izquierda = 1
                Alto = 70
                Arriba = Celda.Top - Alto

                'rutaQr = "C:\Users\Arturo Gallegos\Desktop\MODULOS\ARCHIVOXML\codigoQr.bmp"
                rutaQr = FC_RutaModulos & "\ArchivosIncloud\codigoQr.bmp"
                .Shapes.AddPicture(rutaQr, False, True, Izquierda, Arriba, 90, 80)

            End With

            appXL.DisplayAlerts = False
            '        wbXl.SaveAs("C:\Users\Arturo Gallegos\Desktop\MODULOS\ARCHIVOXML\" & cEmpresa & "\" & factu.SUUID.ToString & ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, False, False,
            '0, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing)
            rutaGuarda = FC_RutaModulos & "\ArchivosIncloud\" & cEmpresa & "\COMPROBANTES\" & complementoruta
            If Not System.IO.Directory.Exists(rutaGuarda) Then
                My.Computer.FileSystem.CreateDirectory(rutaGuarda)
                My.Computer.FileSystem.CreateDirectory(rutaGuarda & "\pdfs")
                My.Computer.FileSystem.CreateDirectory(rutaGuarda & "\pdfs\relacion")
            End If

            wbXl.SaveAs(rutaGuarda & "\" & UCase(factu.SUUID.ToString) & ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, False, False,
    0, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing)

            appXL.ScreenUpdating = False
            appXL.DisplayAlerts = False
            shXL.ExportAsFixedFormat(Type:=Excel.XlFixedFormatType.xlTypePDF,
                                     Filename:=rutaGuarda & "\pdfs\" & UCase(factu.SUUID.ToString),
                                     Quality:=Excel.XlFixedFormatQuality.xlQualityStandard,
                                     IncludeDocProperties:=True, IgnorePrintAreas:=True,
                                     OpenAfterPublish:=False)

            AsociacionesXML(factu, cEmpresa, rutaGuarda & "\pdfs\relacion\" & GlobalRFCEmpresa & "_" & compleRelacion)

            Celda = Nothing
            wbXl.Close()
            wbXl = Nothing
            appXL.Workbooks.Close()

        Catch ex As Exception
            Creafactura = False
            My.Computer.FileSystem.WriteAllText(FC_RutaModulos & "\ArchivosIncloud\" & cEmpresa & "\COMPROBANTES\errores.log", Format(Now, "dd/MM/yyy HH:mm") & " - " & ex.Message & vbCrLf, True)
        Finally
            releaseObject(Celda)
            releaseObject(shXL)
            releaseObject(wbXl)
            releaseObject(appXL)
        End Try
    End Function

    Public Sub AsociacionesXML(ByVal tdocum As CLXml, ByVal tEmpresa As String, ByVal rArchivo As String)
        Dim appExcel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wbExcel As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim wsExcel As Microsoft.Office.Interop.Excel.Worksheet = Nothing

        Dim rutalibro As String, f As Long, linkstring As String
        Dim x As Long
        appExcel = New Microsoft.Office.Interop.Excel.Application
        appExcel.Visible = False
        appExcel.DisplayAlerts = False
        f = 0
        Try
            If Not System.IO.File.Exists(rArchivo & ".xlsx") Then
                rutalibro = FC_RutaModulos & "\ArchivosIncloud\" & tEmpresa & "\COMPROBANTES\relacionxml.xlsx"

            Else
                rutalibro = rArchivo & ".xlsx"
            End If

            wbExcel = appExcel.Workbooks.Open(rutalibro)
            wsExcel = wbExcel.ActiveSheet

            f = getLastRow(wsExcel) + 1

            With wsExcel
                '.Cells(1, 1).value = IIf(tdocum.SRFCEmisor = GlobalRFCEmpresa, Trim(quitarSaltosLinea(tdocum.SNombreEmisor)), Trim(quitarSaltosLinea(tdocum.SNombreReceptor)))
                '.Cells(2, 1).value = "Relación de Comprobantes"
                If f > 4 Then
                    x = 4
                    Do While .Cells(x, 1).value <> ""
                        If .Cells(x, 7).value = tdocum.SUUID.ToString Then
                            f = x
                            Exit Do
                        End If
                        x += 1
                    Loop
                End If

                .Cells(f, 1) = "'" & tdocum.SFecha
                .Cells(f, 2) = IIf(tdocum.STipo = "I", "INGRESO", IIf(tdocum.STipo = "E", "EGRESO", "NOTA"))
                .Cells(f, 3) = tdocum.SFolio
                .Cells(f, 4) = tdocum.SNombreEmisor
                .Cells(f, 5) = tdocum.SNombreReceptor
                .Cells(f, 6) = tdocum.STotalXML
                .Cells(f, 7) = UCase(tdocum.SUUID.ToString)
                'If f = 2 Or .Cells(f, 8).value = "" Then
                linkstring = ""
                'linkstring = getLinkCompartido(directContaXMLNUBE & "/" &
                'Year(tdocum.SFecha) & "/" &
                'UCase(MonthName(Month(tdocum.SFecha))) & "/pdfs/" & tdocum.SUUID.ToString & ".pdf")
                .Cells(f, 9) = directContaXMLNUBE & "/COMPROBANTES/" &
                                                 Right(Year(tdocum.SFecha), 2) & "/" &
                                                 UCase(Format(Month(tdocum.SFecha), "00")) & "/pdfs/" &
                                                 UCase(tdocum.SUUID.ToString) & ".pdf"

                Dim sRut As String, sKey As String
                sKey = GlobalRFCEmpresa & "_" & Right(Year(tdocum.SFecha), 2) & UCase(Format(Month(tdocum.SFecha), "00"))
                sRut = Right(Year(tdocum.SFecha), 2) & "\" & UCase(Format(Month(tdocum.SFecha), "00"))
                If Not dCarpetas.ContainsKey(sKey) Then
                    dCarpetas.Add(sKey, sRut)
                End If

                'Else
                '    linkstring = ""
                'End If
                'If linkstring <> "" Then
                '    .Hyperlinks.Add(Anchor:= .Range("G" & CStr(f)),
                '                                        Address:=linkstring,
                '                                        TextToDisplay:=tdocum.SUUID.ToString)
                'End If
                .PageSetup.Zoom = False
                .PageSetup.FitToPagesTall = 1
                .PageSetup.FitToPagesWide = 1
                .PageSetup.RightMargin = 4
                .PageSetup.LeftMargin = 4
                .PageSetup.BottomMargin = 4
                .PageSetup.TopMargin = 9
                .PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape

                '.PageSetup.Zoom = False
                '.PageSetup.FitToPagesTall = 1
                '.PageSetup.FitToPagesWide = 1
                '.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperLetter
                '.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape

                .Cells(f, 8) = linkstring
                .Columns("B:H").AutoFit
                .Columns("H:I").Hidden = True
                .Cells(1, 1).value = IIf(tdocum.SRFCEmisor = GlobalRFCEmpresa, Trim(quitarSaltosLinea(tdocum.SNombreEmisor)), Trim(quitarSaltosLinea(tdocum.SNombreReceptor)))
                .Cells(2, 1).value = "Relación de Comprobantes"
            End With

            If f = 5 Then
                wbExcel.SaveAs(rArchivo & ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, False, False,
        0, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing)
            Else
                wbExcel.Save()
            End If

            'wbExcel.ExportAsFixedFormat(Type:=Excel.XlFixedFormatType.xlTypePDF,
            '                             Filename:=rArchivo,
            '                             Quality:=Excel.XlFixedFormatQuality.xlQualityStandard,
            '                             IncludeDocProperties:=True, IgnorePrintAreas:=False,
            '                             OpenAfterPublish:=False)

            wsExcel = Nothing
            wbExcel.Close()
            wbExcel = Nothing
            appExcel.Quit()
        Catch ex As Exception
            My.Computer.FileSystem.WriteAllText(FC_RutaModulos & "\ArchivosIncloud\" & tEmpresa & "\COMPROBANTES\errores.log", Format(Now, "dd/MM/yyy HH:mm") & " - " & ex.Message & vbCrLf, True)
        Finally
            releaseObject(wsExcel)
            releaseObject(wbExcel)
            releaseObject(appExcel)
        End Try

    End Sub

    Public Sub ImprimeExpediente(ByVal cUUID As String, ByVal uCon As SqlConnection, ByVal cEmpresa As String)
        Dim appXL As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wbXl As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim shXL As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim cQue As String = "", ref As String, Concep As String
        Dim f As Integer = 4
        Dim Fname As String = "", nomRur As String, nombrerutapdf As String
        'Fname = "C:\Users\Arturo Gallegos\Desktop\MODULOS\ARCHIVOXML\" & cEmpresa & "\" & cUUID & ".xlsx"
        nomRur = FC_RutaModulos & "\ArchivosIncloud\" & cEmpresa & "\COMPROBANTES"
        If Not System.IO.Directory.Exists(nomRur) Then
            Exit Sub
        End If

        'Fname = FC_RutaModulos & "\ArchivosIncloud\ARCHIVOSXML\" & cEmpresa & "\" & cUUID & ".xlsx"
        Fname = GetRutaFile(nomRur, UCase(cUUID) & ".xlsx")

        If System.IO.File.Exists(Fname) Then
            nombrerutapdf = Path.GetDirectoryName(Fname)
            If sConCarpetas = False Then
                DesMarcarLink(nombrerutapdf, UCase(cUUID), cEmpresa)
            End If
            Try
                appXL = New Microsoft.Office.Interop.Excel.Application
                appXL.Visible = False
                wbXl = appXL.Workbooks.Open(Fname)
                shXL = wbXl.ActiveSheet

                With shXL
                    .Range("K4:P100").ClearContents()
                    cQue = "SELECT m.Fecha,t.Nombre as Nompol,m.Folio,c.Codigo,m.Referencia,
                            c.Nombre as nomCuenta,m.TipoMovto,m.Importe,a.UUID,m.Concepto,m.IdPoliza FROM AsocCFDIs a 
                            INNER JOIN MovimientosPoliza AS m ON a.GuidRef = m.Guid 
                            INNER JOIN TiposPolizas t ON m.TipoPol=t.Codigo 
                            INNER JOIN Cuentas c ON m.IdCuenta=c.Id WHERE UUID=@uuid ORDER BY m.Fecha,m.IdPoliza,m.NumMovto"
                    Using mcom = New SqlCommand(cQue, uCon)
                        mcom.Parameters.AddWithValue("@uuid", cUUID)
                        Using mRs = mcom.ExecuteReader()
                            Do While mRs.Read()
                                .Cells(f, 11).value = Format(mRs("Fecha"), "dd/MM/yyyy")
                                .Cells(f, 12).value = mRs("Nompol")
                                .Cells(f, 13).value = mRs("Folio")
                                .Cells(f, 14).value = mRs("Codigo")
                                If mRs("TipoMovto") = 0 Then
                                    .Cells(f, 15).value = mRs("Importe")
                                Else
                                    .Cells(f, 16).value = mRs("Importe")
                                End If
                                f = f + 1
                                ref = IIf(Trim(mRs("Referencia")) <> "", " Ref: " & Trim(mRs("Referencia")), " ")
                                Concep = IIf(Trim(mRs("Concepto")) <> "", " Concepto: " & Trim(mRs("Concepto")), " ")
                                .Cells(f, 11) = mRs("nomCuenta") & ref & Concep
                                f = f + 1
                            Loop
                        End Using
                    End Using
                End With


                appXL.DisplayAlerts = False
                wbXl.Save()

                appXL.ScreenUpdating = False
                appXL.DisplayAlerts = False
                shXL.ExportAsFixedFormat(Type:=Excel.XlFixedFormatType.xlTypePDF,
                                     Filename:=nombrerutapdf & "\pdfs\" & UCase(cUUID),
                                     Quality:=Excel.XlFixedFormatQuality.xlQualityStandard,
                                     IncludeDocProperties:=True, IgnorePrintAreas:=False,
                                     OpenAfterPublish:=False)

                shXL = Nothing
                wbXl.Close()
                wbXl = Nothing
                appXL.Quit()

                'wbXl.Close()
                'wbXl = Nothing
                'appXL.Workbooks.Close()

            Catch ex As Exception
                My.Computer.FileSystem.WriteAllText(FC_RutaModulos & "\ArchivosIncloud\" & cEmpresa & "\COMPROBANTES\errores.log", Format(Now, "01/MM/yyy HH:mm") & " - " & ex.Message & vbCrLf, True)
            Finally
                releaseObject(shXL)
                releaseObject(wbXl)
                releaseObject(appXL)
            End Try
        End If
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    Public Sub KillAllExcels()
        Try

            Dim proc As System.Diagnostics.Process

            For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
                If proc.MainWindowTitle.Trim.Length = 0 Then
                    'proc.GetCurrentProcess.StartInfo
                    proc.Kill()
                End If
            Next
        Catch ex As Exception
            My.Computer.FileSystem.WriteAllText(FC_RutaModulos & "\ArchivosIncloud\errores.log", Format(Now, "01/MM/yyy HH:mm") & " - " & ex.Message & vbCrLf, True)
        End Try
    End Sub

    Private Sub CrearCodigoQR(ByVal codigoQ As String)
        Dim sRut As String
        Dim generador As BarcodeWriter = New BarcodeWriter

        generador.Format = BarcodeFormat.QR_CODE
        'sRut = "C:\Users\Arturo Gallegos\Desktop\MODULOS\ARCHIVOXML\codigoQr.bmp"

        sRut = FC_RutaModulos & "\ArchivosIncloud\codigoQr.bmp"

        Dim imagen As Bitmap = New Bitmap(generador.Write(codigoQ), 500, 500)
        imagen.Save(sRut)
    End Sub

    Private Function GetSumImpuesto(ByVal gCon As SqlConnection, ByVal gTabla As String, ByVal gFiltro As String, ByVal gGuiDDoc As String) As Double
        Dim gQu As String
        GetSumImpuesto = 0
        gQu = "SELECT SUM(Importe) as importe FROM " & gTabla & " WHERE ImpuestoDesc=@filt AND GuidDocument=@Guid"
        Using gCom = New SqlCommand(gQu, gCon)
            gCom.Parameters.AddWithValue("@filt", gFiltro)
            gCom.Parameters.AddWithValue("@Guid", gGuiDDoc)
            Using gRea = gCom.ExecuteReader()
                gRea.Read()
                If gRea.HasRows = True Then
                    If gRea("importe") IsNot DBNull.Value Then GetSumImpuesto = gRea("importe")
                End If
            End Using
        End Using
    End Function

    Public Function ImprimePoliza(ByVal cEmpresa As String, ByVal iIDPoliza As Integer,
                                  ByVal nomPlantilla As String, ByVal FechaI As Date,
                                  FechaF As Date, claEmpresa As CLEmpresa, ByVal cElimina As Boolean) As String
        Dim cQue As String, movQue As String, nomArchivo As String, f As Integer, filAnt As Integer
        Dim appXL As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wbXl As Microsoft.Office.Interop.Excel.Workbook = Nothing

        Dim appXLFac As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wbXlFac As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim shXLFac As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim nomfac As String, uFil As Long, n As Integer

        Dim shXL As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        ImprimePoliza = ""

        Dim tCargo As Double, tAbono As Double, Moneda As String, GuidPoliza As String
        Dim ImpIva, ImpBase, IvaRet, ISRRet, IEPS, OtrosImp, GranTotal, IVApagna As Double
        Dim sEntroRecord As Boolean, cfQue As String

        Dim tImporteTotal, tImporteBase, tImporteIVA, tImporteNoAcred As Double
        Dim complementoruta As String, compleRela As String
        Dim dProveedores As New Dictionary(Of String, String)
        Dim sArr(0 To 6) As String

        If Not System.IO.Directory.Exists(FC_RutaModulos & "\ArchivosIncloud\" & cEmpresa & "\POLIZAS") Then
            Exit Function
        End If

        Try
            cQue = "SELECT p.Id ,p.Fecha, p.TipoPol, tp.Nombre, p.Folio, p.Concepto, p.Guid 
                    FROM Polizas p INNER JOIN TiposPolizas tp ON p.TipoPol = tp.Codigo
                    WHERE p.id=@idpol"
            Using mcom = New SqlCommand(cQue, PConexionesPol(cEmpresa))
                mcom.Parameters.AddWithValue("@idpol", iIDPoliza)
                Using mCr = mcom.ExecuteReader()
                    mCr.Read()
                    If mCr.HasRows Then

                        GuidPoliza = mCr("Guid")

                        appXL = New Microsoft.Office.Interop.Excel.Application
                        appXL.Visible = False

                        wbXl = appXL.Workbooks.Open(FC_RutaModulos & "\ArchivosIncloud\" & cEmpresa & "\POLIZAS\" & nomPlantilla)
                        shXL = wbXl.ActiveSheet

                        With shXL
                            ''ENCABEZADO DOCUMENTO
                            .Cells(1, 1).value = UCase(claEmpresa.CNomEmpresa)
                            sArr(6) = UCase(claEmpresa.CNomEmpresa)
                            .Cells(2, 1).value = "Impreso de póliza del " & Mid(FechaI.ToString, 1, 10) & " al " & Mid(FechaF.ToString, 1, 10)
                            .Cells(4, 2).value = IIf(Trim(claEmpresa.CDireccion) <> "", Trim(claEmpresa.CDireccion), "0")
                            .Cells(5, 2).value = IIf(claEmpresa.CRFCEmpresa <> "", Trim(claEmpresa.CRFCEmpresa), "")
                            .Cells(5, 9).value = IIf(claEmpresa.CRegCamara <> "", Trim(claEmpresa.CRegCamara), "")
                            .Cells(5, 15).value = IIf(claEmpresa.CRegEstatal <> "", Trim(claEmpresa.CRegEstatal), "")
                            .Cells(4, 19).value = IIf(claEmpresa.CCodigoPostal <> "", Trim(claEmpresa.CCodigoPostal), "0")

                            ''ENCABEZADO POLIZA
                            complementoruta = Right(Year(mCr("Fecha")), 2) & "\" & UCase(Format(Month(mCr("Fecha")), "00"))
                            compleRela = Right(Year(mCr("Fecha")), 2) & UCase(Format(Month(mCr("Fecha")), "00"))
                            .Cells(8, 1).value = "'" & IIf(mCr("Fecha").ToString <> "", Mid(mCr("Fecha").ToString, 1, 10), "")
                            sArr(0) = .Cells(8, 1).value ''FECHA
                            .Cells(8, 3).value = IIf(mCr("Nombre") <> "", Trim(mCr("Nombre")), "")
                            sArr(1) = .Cells(8, 3).value ''TIPO
                            .Cells(8, 5).value = IIf(mCr("Folio").ToString <> "", Trim(mCr("Folio").ToString), "")
                            sArr(2) = .Cells(8, 5).value ''FOLIO
                            .Cells(8, 6).value = IIf(mCr("Concepto") <> "", Trim(quitarSaltosLinea(mCr("Concepto"))), "")
                            sArr(3) = .Cells(8, 6).value ''CONCEPTO

                            ''MOVIMIENTOS DE LA POLIZA
                            f = 9
                            movQue = "SELECT mp.TipoMovto, mp.Referencia, mp.Concepto, mp.Importe, c.Nombre, c.Codigo, c.IdMoneda, m.Nombre AS Moneda 
                                        FROM MovimientosPoliza mp 
                                        INNER JOIN Cuentas c On c.Id = mp.IdCuenta 
                                        INNER JOIN Monedas m On c.IdMoneda = m.Id 
                                        WHERE mp.IdPoliza =@idpol"
                            Using movCom = New SqlCommand(movQue, PConexionesPol(cEmpresa))
                                movCom.Parameters.AddWithValue("@idpol", iIDPoliza)
                                Using movCR = movCom.ExecuteReader()
                                    Do While movCR.Read()
                                        .Cells(f, 2).value = IIf(movCR("Referencia") <> "", "'" & Trim(movCR("Referencia")), "")
                                        .Cells(f, 4).value = IIf(movCR("Codigo") <> "", "'" & Trim(movCR("Codigo")), "")
                                        .Cells(f, 7).value = IIf(movCR("Nombre") <> "", Trim(movCR("Nombre")) & " - " & Trim(movCR("Concepto")), "")

                                        If movCR("TipoMovto") = 0 Then
                                            .Cells(f, 18).value = Format(movCR("Importe"), "#,##0.00")
                                            .Cells(f, 18).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                                            tCargo = tCargo + movCR("Importe")
                                        Else
                                            .Cells(f, 19).value = Format(movCR("Importe"), "#,##0.00")
                                            tAbono = tAbono + movCR("Importe")
                                        End If
                                        Moneda = movCR("Moneda")
                                        f = f + 1
                                        .Rows(f).Insert()
                                    Loop
                                End Using
                            End Using

                            f = f + 1
                            .Rows(f).Insert()
                            .Cells(3, 1).value = "Moneda: " & Moneda
                            .Cells(f, 17).value = "Total póliza"
                            .Cells(f, 18).value = Format(tCargo, "#,##0.00")
                            .Cells(f, 19).value = Format(tAbono, "#,##0.00")
                            sArr(4) = .Cells(f, 18).value ''TOTAL
                            f = f + 1
                            .Rows(f).Insert()
                            f = f + 3
                            ''INFORMACION PARA LA DIOT
                            sEntroRecord = False
                            movQue = "SELECT d.IdPoliza, IdProveedor, ImpTotal, PorIVA, ImpBase, ImpIVA, CausaIVA, OtrosImptos, p.nombre, p.codigo,
                                      IVARetenido, ISRRetenido, GranTotal, EjercicioAsignado, PeriodoAsignado, d.IdCuenta, IVAPagNoAcred, IEPS 
                                        FROM DevolucionesIVA d INNER JOIN Proveedores p ON d.IdProveedor=p.ID
                                            WHERE IdPoliza =@idpol"
                            Using movCom = New SqlCommand(movQue, PConexionesPol(cEmpresa))
                                movCom.Parameters.AddWithValue("@idpol", iIDPoliza)
                                Using movCR = movCom.ExecuteReader() ''CAUSACION IVA
                                    ImpIva = 0 : ImpBase = 0 : IvaRet = 0 : ISRRet = 0 : IEPS = 0 : OtrosImp = 0 : GranTotal = 0 : IVApagna = 0
                                    Do While movCR.Read()
                                        sEntroRecord = True
                                        ImpBase = ImpBase + movCR("ImpBase")
                                        ImpIva = ImpIva + movCR("ImpIVA")
                                        IvaRet = IvaRet + movCR("IVARetenido")
                                        ISRRet = ISRRet + movCR("ISRRetenido")
                                        IEPS = IEPS + movCR("IEPS")
                                        OtrosImp = OtrosImp + movCR("OtrosImptos")
                                        GranTotal = GranTotal + movCR("GranTotal")
                                        IVApagna = IVApagna + movCR("IVAPagNoAcred")
                                        .Cells(f, 1).Value = movCR("codigo")

                                        If Not dProveedores.ContainsKey(CStr(movCR("Codigo"))) Then
                                            dProveedores.Add(CStr(movCR("Codigo")), CStr(movCR("nombre")))
                                        End If

                                        .Cells(f, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft

                                        .Cells(f, 2).Value = IIf(movCR("PorIva") = 0, .Cells(f, 2).Value, movCR("PorIva") & "%")
                                        .Cells(f, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

                                        .Cells(f, 3).Value = Format(movCR("ImpBase"), "#,##0.00")
                                        .Cells(f, 3).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

                                        .Cells(f, 5).Value = Format(movCR("ImpIVA"), "#,##0.00")
                                        .Cells(f, 5).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

                                        .Cells(f, 7).Value = Format(movCR("IVARetenido"), "#,##0.00")
                                        .Cells(f, 7).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

                                        .Cells(f, 9).Value = Format(movCR("ISRRetenido"), "#,##0.00")
                                        .Cells(f, 9).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

                                        .Cells(f, 11).Value = Format(movCR("IEPS"), "#,##0.00")
                                        .Cells(f, 11).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

                                        .Cells(f, 13).Value = Format(movCR("OtrosImptos"), "#,##0.00")
                                        .Cells(f, 13).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

                                        .Cells(f, 15).Value = Format(movCR("GranTotal"), "#,##0.00")
                                        .Cells(f, 15).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

                                        .Cells(f, 17).Value = Format(movCR("IVAPagNoAcred"), "#,##0.00")
                                        .Cells(f, 17).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

                                        .Cells(f, 18).Value = IIf(movCR("CausaIVA") = True, "Si", "No")
                                        .Cells(f, 18).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

                                        .Cells(f, 19).Value = movCR("EjercicioAsignado") & "-" & movCR("PeriodoAsignado")
                                        .Cells(f, 19).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight
                                        f = f + 1
                                        .Rows(f).Insert()
                                    Loop
                                End Using
                            End Using
                            ''ELIMINA LA DIOT O LA PROVISION
                            If sEntroRecord = False Then
                                .Rows(f - 2).Delete()
                                .Rows(f - 2).Delete()
                                .Rows(f - 2).Delete()
                                .Rows(f - 2).Delete()
                                .Rows(f - 2).Delete()
                                movQue = "SELECT c.IdPoliza, c.Tipo, c.TotTasa16, c.BaseTasa16, c.IVATasa16, c.IVATasa16NoAcred, c.TotTasa11, c.BaseTasa11, c.IVATasa11, c.IVATasa11NoAcred, c.TotTasa0, 
                                            c.BaseTasa0, c.TotTasa15, c.TotTasaExento, c.BaseTasaExento, c.BaseTasa15, c.IVATasa15, c.IVATasa15NoAcred, c.TotTasa10, c.BaseTasa10, c.IVATasa10, c.IVATasa10NoAcred, c.TotOtraTasa,
                                            c.BaseOtraTasa, c.IVAOtraTasa, c.IVARetenido, c.ISRRetenido, c.IEPS, c.TotOtros, c.IETU, con.Nombre ,c.TotTasa8, c.BaseTasa8, c.IVATasa8, c.IVATasa8NoAcred
                                            FROM CausacionesIVA c 
                                            LEFT JOIN ConceptosIETU con ON c.IdConceptoIETU = con.Id 
                                            WHERE c.IdPoliza =@idpol"
                                Using movCom = New SqlCommand(movQue, PConexionesPol(cEmpresa))
                                    movCom.Parameters.AddWithValue("@idpol", iIDPoliza)
                                    Using movCR = movCom.ExecuteReader()
                                        movCR.Read()
                                        If movCR.HasRows Then
                                            .Cells(f, 1).Value = "IVA " & IIf(movCR("Tipo") = 1, "CAUSADO", "ACREDITABLE")
                                            .Cells(f, 9).Value = movCR("TotTasa16")
                                            .Cells(f, 11).Value = movCR("BaseTasa16")
                                            .Cells(f, 13).Value = movCR("IVATasa16")
                                            .Cells(f, 15).Value = movCR("IVATasa16NoAcred")

                                            f = f + 1
                                            .Cells(f, 9).Value = movCR("TotTasa8")
                                            .Cells(f, 11).Value = movCR("BaseTasa8")
                                            .Cells(f, 13).Value = movCR("IVATasa8")
                                            .Cells(f, 15).Value = movCR("IVATasa8NoAcred")

                                            f = f + 1
                                            .Cells(f, 9).Value = movCR("TotTasa11")
                                            .Cells(f, 11).Value = movCR("BaseTasa11")
                                            .Cells(f, 13).Value = movCR("IVATasa11")
                                            .Cells(f, 15).Value = movCR("IVATasa11NoAcred")

                                            f = f + 1
                                            .Cells(f, 9).Value = movCR("TotTasa0")
                                            .Cells(f, 11).Value = movCR("BaseTasa0")

                                            f = f + 1
                                            .Cells(f, 9).Value = movCR("TotTasaExento")
                                            .Cells(f, 11).Value = movCR("BaseTasaExento")

                                            f = f + 1
                                            .Cells(f, 9).Value = movCR("TotTasa15")
                                            .Cells(f, 11).Value = movCR("BaseTasa15")
                                            .Cells(f, 13).Value = movCR("IVATasa15")
                                            .Cells(f, 15).Value = movCR("IVATasa15NoAcred")

                                            f = f + 1
                                            .Cells(f, 9).Value = movCR("TotTasa10")
                                            .Cells(f, 11).Value = movCR("BaseTasa10")
                                            .Cells(f, 13).Value = movCR("IVATasa10")
                                            .Cells(f, 15).Value = movCR("IVATasa10NoAcred")

                                            f = f + 1
                                            .Cells(f, 9).Value = movCR("TotOtraTasa")
                                            .Cells(f, 11).Value = movCR("BaseOtraTasa")
                                            .Cells(f, 11).Value = movCR("IVAOtraTasa")

                                            f = f + 1
                                            .Cells(f, 9).Value = movCR("IVARetenido")
                                            f = f + 1
                                            .Cells(f, 9).Value = movCR("ISRRetenido")
                                            f = f + 1
                                            .Cells(f, 9).Value = movCR("IEPS")
                                            f = f + 1
                                            .Cells(f, 9).Value = movCR("TotOtros")

                                            tImporteTotal = movCR("TotTasa16") + movCR("TotTasa8") + movCR("TotTasa11") +
                                                            movCR("TotTasa0") + movCR("TotTasaExento") +
                                                            movCR("BaseTasa15") + movCR("TotTasa10") +
                                                            movCR("TotOtraTasa") + movCR("IVARetenido") +
                                                            movCR("ISRRetenido") + movCR("IEPS") + movCR("TotOtros")

                                            tImporteBase = movCR("BaseTasa16") + movCR("BaseTasa8") +
                                                            movCR("BaseTasa11") + movCR("BaseTasa0") +
                                                            movCR("BaseTasaExento") + movCR("BaseTasa15") +
                                                            movCR("BaseTasa10") + movCR("BaseOtraTasa")

                                            tImporteIVA = movCR("IVATasa16") + movCR("IVATasa8") +
                                                            movCR("IVATasa11") + movCR("IVATasa15") +
                                                            movCR("IVATasa10") + movCR("IVAOtraTasa")

                                            tImporteNoAcred = movCR("IVATasa16NoAcred") + movCR("IVATasa8NoAcred") +
                                                              movCR("IVATasa11NoAcred") + movCR("IVATasa15NoAcred") +
                                                              movCR("IVATasa10NoAcred")


                                            f = f + 1
                                            .Cells(f, 9).Value = IIf(tImporteTotal <> 0, tImporteTotal, 0)
                                            .Cells(f, 11).Value = IIf(tImporteBase <> 0, tImporteBase, 0)
                                            .Cells(f, 13).Value = IIf(tImporteIVA <> 0, tImporteIVA, 0)
                                            .Cells(f, 15).Value = IIf(tImporteNoAcred <> 0, tImporteNoAcred, 0)

                                            f = f + 1
                                            .Cells(f, 11).Value = movCR("IETU")
                                            .Cells(f, 14).Value = Trim(IIf(movCR("Nombre") IsNot DBNull.Value, movCR("Nombre"), "Ninguno"))
                                            f = f + 2
                                        Else
                                            f = f - 1
                                            .Rows(f - 1).Delete()
                                            .Rows(f - 1).Delete()
                                            .Rows(f - 1).Delete()
                                            .Rows(f - 1).Delete()
                                            .Rows(f - 1).Delete()
                                            .Rows(f - 1).Delete()
                                            .Rows(f - 1).Delete()
                                            .Rows(f - 1).Delete()
                                            .Rows(f - 1).Delete()
                                            .Rows(f - 1).Delete()
                                            .Rows(f - 1).Delete()
                                            .Rows(f - 1).Delete()
                                            .Rows(f - 1).Delete()
                                            .Rows(f - 1).Delete()
                                            .Rows(f - 1).Delete()
                                            .Rows(f - 1).Delete()

                                        End If
                                    End Using
                                End Using
                            Else
                                f = f + 2
                                .Rows(f - 1).Delete()
                                .Rows(f - 1).Delete()
                                .Rows(f - 1).Delete()
                                .Rows(f - 1).Delete()
                                .Rows(f - 1).Delete()
                                .Rows(f - 1).Delete()
                                .Rows(f - 1).Delete()
                                .Rows(f - 1).Delete()
                                .Rows(f - 1).Delete()
                                .Rows(f - 1).Delete()
                                .Rows(f - 1).Delete()
                                .Rows(f - 1).Delete()
                                .Rows(f - 1).Delete()
                                .Rows(f - 1).Delete()
                                .Rows(f - 1).Delete()
                                .Rows(f - 1).Delete()
                                .Rows(f - 1).Delete()
                                .Rows(f - 1).Delete()
                            End If


                            f = f + 2
                            n = 1
                            movQue = "SELECT UUID FROM AsocCFDIs WHERE GuidRef =@guidRef"
                            Using cCom = New SqlCommand(movQue, PConexionesPol(cEmpresa))
                                cCom.Parameters.AddWithValue("@guidRef", GuidPoliza)
                                Using cRs = cCom.ExecuteReader()
                                    Do While cRs.Read
                                        cfQue = "SELECT Fecha,TipoComprobante,Serie,Folio, 
                                                 RFCEmisor,NombreEmisor,RFCReceptor, NombreReceptor,Total
                                                FROM Comprobante WHERE UUID=@uuid"
                                        Using mComC = New SqlCommand(cfQue, FC_ConGuid)
                                            mComC.Parameters.AddWithValue("@uuid", cRs("UUID"))
                                            Using movCr = mComC.ExecuteReader()
                                                movCr.Read()
                                                If movCr.HasRows Then
                                                    .Cells(f, 1).Value = Mid(movCr("Fecha").ToString, 1, 10)
                                                    .Cells(f, 2).Value = IIf(movCr("TipoComprobante") = "I", "INGRESO", "EGRESO")
                                                    .Cells(f, 3).Value = movCr("Serie")
                                                    .Cells(f, 5).Value = movCr("Folio")
                                                    .Cells(f, 5).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                                                    .Cells(f, 7).Value = cRs("UUID")
                                                    .Cells(f, 13).Value = IIf(claEmpresa.CRFCEmpresa = movCr("RFCEmisor"), movCr("RFCReceptor"), movCr("RFCEmisor"))
                                                    .Cells(f, 19).Value = Format(movCr("Total"), "#,##0.00")
                                                    .Cells(f, 19).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                                                    f = f + 1
                                                    .Cells(f, 13).Value = IIf(claEmpresa.CRFCEmpresa = movCr("RFCEmisor"), Trim(quitarSaltosLinea(movCr("NombreReceptor"))), Trim(quitarSaltosLinea(movCr("NombreEmisor"))))

                                                    'nomfac = FC_RutaModulos & "\ArchivosIncloud\ARCHIVOSXML\" & cEmpresa & "\" & cRs("UUID") & ".xlsx"
                                                    nomfac = GetRutaFile(FC_RutaModulos & "\ArchivosIncloud\" & cEmpresa & "\COMPROBANTES", cRs("UUID") & ".xlsx")
                                                    If System.IO.File.Exists(nomfac) Then
                                                        filAnt = f - 1

                                                        appXLFac = New Microsoft.Office.Interop.Excel.Application
                                                        appXLFac.Visible = False
                                                        wbXlFac = appXLFac.Workbooks.Open(nomfac)
                                                        shXLFac = wbXlFac.ActiveSheet
                                                        uFil = getLastRow(shXLFac)


                                                        n = n + 1

                                                        appXLFac.ActiveWindow.DisplayGridlines = False
                                                        shXLFac.Range("A1:P" & CStr(uFil)).Copy()

                                                        Dim sheetTemp As Excel.Worksheet
                                                        sheetTemp = wbXl.Sheets(n)
                                                        sheetTemp.Activate()

                                                        With sheetTemp.PageSetup
                                                            '.PaperSize = Excel.XlPaperSize.xlPaperLetter
                                                            '.Orientation = Excel.XlPageOrientation.xlLandscape
                                                            .Zoom = False
                                                            .FitToPagesTall = 1
                                                            .FitToPagesWide = 1
                                                            .RightMargin = 4
                                                            .LeftMargin = 4
                                                            .BottomMargin = 4
                                                            .TopMargin = 9
                                                            .Orientation = Excel.XlPageOrientation.xlLandscape

                                                            wbXl.Sheets(n).Range("A2").PasteSpecial(
                                                                Paste:=Excel.XlPasteType.xlPasteAll, SkipBlanks:=False)
                                                            wbXl.Sheets(n).Range("A1") = cRs("UUID")
                                                        End With
                                                        appXLFac.ActiveWindow.DisplayGridlines = True

                                                        appXLFac.CutCopyMode = False

                                                        If n > 3 Then
                                                            wbXl.Worksheets.Add(After:=wbXl.Worksheets(wbXl.Worksheets.Count))
                                                        End If

                                                        .Hyperlinks.Add(Anchor:= .Range("G" & CStr(filAnt)),
                                                        Address:="", SubAddress:=wbXl.Sheets(n).Name & "!A1",
                                                        TextToDisplay:=cRs("UUID").ToString)

                                                        wbXl.Sheets(n).Hyperlinks.Add(Anchor:=wbXl.Sheets(n).Range("A1"),
                                                        Address:="", SubAddress:= .Name & "!G" & CStr(filAnt),
                                                        TextToDisplay:=cRs("UUID").ToString)

                                                        shXLFac = Nothing
                                                        wbXlFac.Close(False)
                                                        wbXlFac = Nothing
                                                        appXLFac.Quit()
                                                        appXLFac = Nothing

                                                    End If
                                                    f = f + 1
                                                End If
                                            End Using
                                        End Using
                                    Loop
                                End Using
                            End Using
                            f = f + 3
                            '''IMPRIME CODIGOS DE PROVEEDORES
                            If sEntroRecord Then
                                .Cells(f, 2) = "Relación de Proveedores DIOT"
                                .Cells(f, 1).Font.Bold = True
                                f = f + 1
                                .Cells(f, 1) = "CODIGO"
                                .Cells(f, 1).Font.Bold = True
                                .Cells(f, 3) = "PROVEEDOR"
                                .Cells(f, 3).Font.Bold = True
                                f = f + 1
                                For t = 0 To dProveedores.Count - 1
                                    .Cells(f, 1) = dProveedores.Keys(t)
                                    .Cells(f, 3) = dProveedores.Item(dProveedores.Keys(t))
                                    f = f + 1
                                Next
                            End If
                            appXL.DisplayAlerts = False
                            appXL.ScreenUpdating = False
                            .PageSetup.RightMargin = 4
                            .PageSetup.LeftMargin = 4
                            .PageSetup.BottomMargin = 4
                            .PageSetup.TopMargin = 9
                            '        wbXl.SaveAs(FC_RutaModulos & "\POLIZAS\" & cEmpresa & "\" & GuidPoliza & ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, False, False,
                            '0, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing)
                            nomArchivo = FC_RutaModulos & "\ArchivosIncloud\" & cEmpresa & "\POLIZAS\" & complementoruta
                            'compleRela
                            sArr(5) = GuidPoliza ''GUID
                            ImprimePoliza = GuidPoliza
                            If Not System.IO.Directory.Exists(nomArchivo) Then
                                My.Computer.FileSystem.CreateDirectory(nomArchivo)
                                My.Computer.FileSystem.CreateDirectory(nomArchivo & "\relacion")
                            End If
                            CrearPolizaPDF(appXL, nomArchivo & "\" & GuidPoliza & ".pdf")

                            compleRela = nomArchivo & "\relacion\" & GlobalRFCEmpresa & "_" & compleRela
                            AsociacionesPolizas(cEmpresa, compleRela, sArr, cElimina)

                            shXL = Nothing
                            wbXl.Close(False)
                            wbXl = Nothing
                            appXL.Quit()
                            appXL = Nothing


                        End With
                    End If
                End Using
            End Using
        Catch ex As Exception
            ImprimePoliza = ""
            My.Computer.FileSystem.WriteAllText(FC_RutaModulos & "\ArchivosIncloud\" & cEmpresa & "\POLIZAS\errores.log", Format(Date.Now, "dd/MM/yyyy HH:mm") & " - " & ex.Message & vbCrLf, True)
        Finally
            releaseObject(shXL)
            releaseObject(wbXl)
            releaseObject(appXL)
            releaseObject(shXLFac)
            releaseObject(wbXlFac)
            releaseObject(appXLFac)
            dProveedores.Clear()
        End Try


    End Function

    Public Sub AsociacionesPolizas(ByVal tEmpresa As String, ByVal rArchivo As String,
                                   ByVal dPol() As String, ByVal tEliminas As Boolean)
        Dim appExcel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wbExcel As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim wsExcel As Microsoft.Office.Interop.Excel.Worksheet = Nothing

        Dim rutalibro As String, f As Long, linkstring As String
        Dim x As Long
        appExcel = New Microsoft.Office.Interop.Excel.Application
        appExcel.Visible = False
        appExcel.DisplayAlerts = False
        f = 0
        Try
            If Not System.IO.File.Exists(rArchivo & ".xlsx") Then
                rutalibro = FC_RutaModulos & "\ArchivosIncloud\" & tEmpresa & "\POLIZAS\relacionpolizas.xlsx"
            Else
                rutalibro = rArchivo & ".xlsx"
            End If

            wbExcel = appExcel.Workbooks.Open(rutalibro)
            wsExcel = wbExcel.ActiveSheet

            f = getLastRow(wsExcel) + 1

            With wsExcel
                '.Cells(1, 1).value = dPol(6)
                '.Cells(2, 1).value = "Relación de Polizas"
                If f > 4 Then
                    x = 4
                    Do While .Cells(x, 1).value <> ""
                        If .Cells(x, 6).value = dPol(5) Then
                            f = x
                            If tEliminas = True Then
                                .Rows(f).Delete()
                            End If
                            Exit Do
                        End If
                        x += 1
                    Loop
                End If
                If tEliminas = False Then
                    .Cells(f, 1) = "'" & dPol(0)
                    .Cells(f, 2) = dPol(1)
                    .Cells(f, 3) = dPol(2)
                    .Cells(f, 4) = dPol(3)
                    .Cells(f, 5) = dPol(4)
                    .Cells(f, 6) = dPol(5)

                    'If f = 2 Or .Cells(f, 8).value = "" Then
                    linkstring = ""
                    'linkstring = getLinkCompartido(directContaXMLNUBE & "/" &
                    'Year(tdocum.SFecha) & "/" &
                    'UCase(MonthName(Month(tdocum.SFecha))) & "/pdfs/" & tdocum.SUUID.ToString & ".pdf")
                    .Cells(f, 8) = directContaXMLNUBE & "/POLIZAS/" &
                                                     Right(Year(dPol(0)), 2) & "/" &
                                                     UCase(Format(Month(dPol(0)), "00")) & "/" &
                                                     UCase(dPol(5)) & ".pdf"

                    Dim sRut As String, sKey As String
                    sKey = GlobalRFCEmpresa & "_" & Right(Year(dPol(0)), 2) & UCase(Format(Month(dPol(0)), "00"))
                    sRut = Right(Year(dPol(0)), 2) & "\" & UCase(Format(Month(dPol(0)), "00"))
                    If Not dCarpetasPol.ContainsKey(sKey) Then
                        dCarpetasPol.Add(sKey, sRut)
                    End If

                    'Else
                    '    linkstring = ""
                    'End If
                    'If linkstring <> "" Then
                    '    .Hyperlinks.Add(Anchor:= .Range("G" & CStr(f)),
                    '                                        Address:=linkstring,
                    '                                        TextToDisplay:=tdocum.SUUID.ToString)
                    'End If
                    .PageSetup.Zoom = False
                    .PageSetup.FitToPagesTall = 1
                    .PageSetup.FitToPagesWide = 1
                    .PageSetup.RightMargin = 4
                    .PageSetup.LeftMargin = 4
                    .PageSetup.BottomMargin = 4
                    .PageSetup.TopMargin = 9
                    .PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape

                    '.PageSetup.Zoom = False
                    '.PageSetup.FitToPagesTall = 1
                    '.PageSetup.FitToPagesWide = 1
                    '.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperLetter
                    '.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape

                    .Cells(f, 7) = linkstring
                    .Columns("G:H").AutoFit
                    .Cells(1, 1).value = dPol(6)
                    .Cells(2, 1).value = "Relación de Polizas"
                End If
            End With

            If f = 5 Then
                wbExcel.SaveAs(rArchivo & ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, False, False,
        0, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing)
            Else
                wbExcel.Save()
            End If

            'wbExcel.ExportAsFixedFormat(Type:=Excel.XlFixedFormatType.xlTypePDF,
            '                             Filename:=rArchivo,
            '                             Quality:=Excel.XlFixedFormatQuality.xlQualityStandard,
            '                             IncludeDocProperties:=True, IgnorePrintAreas:=False,
            '                             OpenAfterPublish:=False)

            wsExcel = Nothing
            wbExcel.Close()
            wbExcel = Nothing
            appExcel.Quit()
        Catch ex As Exception
            My.Computer.FileSystem.WriteAllText(FC_RutaModulos & "\ArchivosIncloud\" & tEmpresa & "\POLIZAS\errores.log", Format(Now, "dd/MM/yyy HH:mm") & " - " & ex.Message & vbCrLf, True)
        Finally
            releaseObject(wsExcel)
            releaseObject(wbExcel)
            releaseObject(appExcel)
        End Try

    End Sub

    Private Sub CrearPolizaPDF(ByVal libro As Excel.Application, ByVal nombreFinal As String)
        Dim pmkr2 As AdobePDFMakerForOffice.PDFMaker
        'Set pmkr2 = Application.COMAddIns.Item(6).Object ' Assign object reference.
        pmkr2 = Nothing
        'libro.Visible = False
        'libro.EnableEvents = False
        'libro.ScreenUpdating = False
        'libro.DisplayAlerts = False
        'Dim app As New Application
        For Each a In libro.COMAddIns
            If InStr(UCase(a.Description), "PDFMAKER") > 0 Then
                pmkr2 = a.Object
                Exit For
            End If
        Next

        If pmkr2 Is Nothing Then
            MsgBox("Cannot Find PDFMaker add-in", vbOKOnly, "")
            Exit Sub
        End If

        Dim pdfname As String
        pdfname = nombreFinal

        Dim stng As AdobePDFMakerForOffice.ISettings
        pmkr2.GetCurrentConversionSettings(stng)


        stng.AddBookmarks = True
        stng.AddLinks = True
        stng.AddTags = True
        stng.ConvertAllPages = True
        stng.CreateFootnoteLinks = True
        stng.CreateXrefLinks = True
        stng.OutputPDFFileName = pdfname
        stng.PromptForPDFFilename = False
        stng.ShouldShowProgressDialog = False
        stng.ViewPDFFile = False

        stng.LayoutBasedOnPrinterSettings = True
        stng.PrintActivesheetOnly = False '' para imprimir todas las hojas

        pmkr2.CreatePDFEx(stng, 0)

        stng = Nothing
        pmkr2 = Nothing ' Discontinue associa
        'libro.EnableEvents = True
        'libro.ScreenUpdating = True
        'libro.DisplayAlerts = True
        KillProcess("acrodist.exe")
    End Sub

    Public Sub KillProcess(ByVal processName As String)
        On Error GoTo ErrHandler
        Dim oWMI
        Dim ret
        Dim sService
        Dim oWMIServices
        Dim oWMIService
        Dim oServices
        Dim oService
        Dim servicename

        oWMI = GetObject("winmgmts:")
        oServices = oWMI.InstancesOf("win32_process")

        For Each oService In oServices
            servicename =
    LCase(Trim(CStr(oService.Name) & ""))

            If InStr(1, servicename,
    LCase(processName), vbTextCompare) > 0 Then
                ret = oService.Terminate
            End If
        Next

        oServices = Nothing
        oWMI = Nothing
        Exit Sub
ErrHandler:
        Err.Clear()
    End Sub

    Private Sub DesMarcarLink(ByVal mRuta As String, ByVal mUUId As String, ByVal mEmpresa As String)
        Dim appExcel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wbExcel As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim wsExcel As Microsoft.Office.Interop.Excel.Worksheet = Nothing

        Dim x As Long

        Try
            Dim rArchivo As String = mRuta
            Dim dDatos() As String = Split(rArchivo, "\")
            Dim sMes As String = dDatos(UBound(dDatos))
            Dim sYear As String = dDatos(UBound(dDatos) - 1)
            rArchivo = rArchivo & "\pdfs\relacion\" & GlobalRFCEmpresa & "_" & sMes & sYear & ".xlsx"
            If Not System.IO.File.Exists(rArchivo) Then
                Exit Sub
            End If
            appExcel = New Microsoft.Office.Interop.Excel.Application
            appExcel.Visible = False
            appExcel.DisplayAlerts = False

            wbExcel = appExcel.Workbooks.Open(rArchivo)
            wsExcel = wbExcel.ActiveSheet

            With wsExcel
                x = 2
                Do While .Cells(x, 1).value <> ""
                    If UCase(.Cells(x, 7).value) = UCase(mUUId) Then
                        .Cells(x, 8).value = ""
                        Exit Do
                    End If
                    x += 1
                Loop
            End With

            Dim sRut As String, sKey As String
            sKey = GlobalRFCEmpresa & "_" & UCase(sMes) & sYear
            sRut = UCase(sMes) & "\" & sYear
            If Not dCarpetas.ContainsKey(sKey) Then
                dCarpetas.Add(sKey, sRut)
            End If

            wbExcel.Save()
            wsExcel = Nothing
            wbExcel.Close()
            wbExcel = Nothing
            appExcel.Quit()
        Catch ex As Exception
            My.Computer.FileSystem.WriteAllText(FC_RutaModulos & "\ArchivosIncloud\" & mEmpresa & "\COMPROBANTES\errores.log", Format(Now, "dd/MM/yyy HH:mm") & " - " & ex.Message & vbCrLf, True)
        Finally
            releaseObject(wsExcel)
            releaseObject(wbExcel)
            releaseObject(appExcel)
        End Try
    End Sub
End Module
