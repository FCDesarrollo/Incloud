Imports System.Data.SqlClient
Imports ZXing
Module modImprimeXML
    Public cgXML As New Collection
    Public Const tFactura As String = "FACTURA"
    Public Const tPoliza As String = "POLIZA"
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
        rutaExiste = FC_RutaModulos & "\ARCHIVOSXML\" & sEmpresa & "\"

        If Not System.IO.Directory.Exists(rutaExiste) Then
            Exit Sub
        End If

        If xVersion = 0 Then
            cQue = "SELECT RFCEmisor,NombreEmisor,RegimenEmisor, RFCReceptor, NombreReceptor, GuidDocument,
                      Version, Serie, Folio, Fecha, FormaPago, CondicionesPago, Subtotal, Descuento, TipoCambio, Moneda, Total, TipoComprobante, MetodoPago, 
                      LugarExp, UUID, FechaTimbrado, NumeroCertificado, TipoDocumento, UsoCFDI FROM Comprobante
                    WHERE Cast(Fecha As Date)>=@fecha and Cast(Fecha As Date)<=@fechaF"
        Else
            cQue = "SELECT RFCEmisor,NombreEmisor,RegimenEmisor, RFCReceptor, NombreReceptor, GuidDocument,
                          Version, Serie, Folio, Fecha, FormaPago, CondicionesPago, Subtotal, Descuento, TipoCambio, Moneda, Total, TipoComprobante, MetodoPago, 
                          LugarExp, UUID, FechaTimbrado, NumeroCertificado, TipoDocumento, UsoCFDI FROM Comprobante
                           WHERE Cast(Fecha As Date)>=@fecha AND Cast(Fecha As Date)<=@fechaF AND UUID=@uuid"
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
        Dim nXml As CLXml
        For Each nXml In cgXML
            If Creafactura(nXml, cEmpresa, plantilla) = True Then

            End If

        Next nXml

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
        Dim Fname As String = ""
        Creafactura = True
        Dim i As Integer = 11
        Try
            'Fname = "C:\Users\Arturo Gallegos\Desktop\MODULOS\plantillafactura.xlsx"
            Fname = FC_RutaModulos & "\ARCHIVOSXML\" & cEmpresa & "\" & plantilla

            appXL = New Microsoft.Office.Interop.Excel.Application
            appXL.Visible = False
            wbXl = appXL.Workbooks.Open(Fname)
            shXL = wbXl.ActiveSheet
            With shXL
                ''ENCABEZADO
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
                rutaQr = FC_RutaModulos & "\ARCHIVOSXML\codigoQr.bmp"
                .Shapes.AddPicture(rutaQr, False, True, Izquierda, Arriba, 90, 80)

            End With

            appXL.DisplayAlerts = False
            '        wbXl.SaveAs("C:\Users\Arturo Gallegos\Desktop\MODULOS\ARCHIVOXML\" & cEmpresa & "\" & factu.SUUID.ToString & ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, False, False,
            '0, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing)

            wbXl.SaveAs(FC_RutaModulos & "\ARCHIVOSXML\" & cEmpresa & "\" & factu.SUUID.ToString & ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, False, False,
    0, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing)
            Celda = Nothing
            wbXl.Close()
            wbXl = Nothing
            appXL.Workbooks.Close()

        Catch ex As Exception
            Creafactura = False
            My.Computer.FileSystem.WriteAllText(FC_RutaModulos & "\ARCHIVOSXML\errores.log", Format(Now, "01/MM/yyy HH:mm") & " - " & ex.Message & vbCrLf, True)
        Finally
            releaseObject(Celda)
            releaseObject(shXL)
            releaseObject(wbXl)
            releaseObject(appXL)
        End Try
    End Function

    Public Sub ImprimeExpediente(ByVal cUUID As String, ByVal uCon As SqlConnection, ByVal cEmpresa As String)
        Dim appXL As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wbXl As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim shXL As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim cQue As String = "", ref As String, Concep As String
        Dim f As Integer = 4
        Dim Fname As String = ""
        'Fname = "C:\Users\Arturo Gallegos\Desktop\MODULOS\ARCHIVOXML\" & cEmpresa & "\" & cUUID & ".xlsx"
        Fname = FC_RutaModulos & "\ARCHIVOSXML\" & cEmpresa & "\" & cUUID & ".xlsx"
        If Not System.IO.Directory.Exists(FC_RutaModulos & "\ARCHIVOSXML\" & cEmpresa) Then
            Exit Sub
        End If

        If System.IO.File.Exists(Fname) Then
            Try

                appXL = New Microsoft.Office.Interop.Excel.Application
                appXL.Visible = False
                wbXl = appXL.Workbooks.Open(Fname)
                shXL = wbXl.ActiveSheet

                With shXL
                    .Range("K4:P100").ClearContents()
                    cQue = "SELECT m.Fecha,t.Nombre as Nompol,m.Folio,c.Codigo,m.Referencia,
                            c.Nombre as nomCuenta,m.TipoMovto,m.Importe,a.UUID,m.Concepto FROM AsocCFDIs a 
                            INNER JOIN MovimientosPoliza AS m ON a.GuidRef = m.Guid 
                            INNER JOIN TiposPolizas t ON m.TipoPol=t.Codigo 
                            INNER JOIN Cuentas c ON m.IdCuenta=c.Id WHERE UUID=@uuid ORDER BY m.NumMovto"
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
                wbXl.Close()
                wbXl = Nothing
                appXL.Workbooks.Close()

            Catch ex As Exception
                My.Computer.FileSystem.WriteAllText(FC_RutaModulos & "\ARCHIVOSXML\errores.log", Format(Now, "01/MM/yyy HH:mm") & " - " & ex.Message & vbCrLf, True)
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
            My.Computer.FileSystem.WriteAllText(FC_RutaModulos & "\ARCHIVOSXML\errores.log", Format(Now, "01/MM/yyy HH:mm") & " - " & ex.Message & vbCrLf, True)
        End Try
    End Sub

    Private Sub CrearCodigoQR(ByVal codigoQ As String)
        Dim sRut As String
        Dim generador As BarcodeWriter = New BarcodeWriter

        generador.Format = BarcodeFormat.QR_CODE
        'sRut = "C:\Users\Arturo Gallegos\Desktop\MODULOS\ARCHIVOXML\codigoQr.bmp"

        sRut = FC_RutaModulos & "\ARCHIVOSXML\codigoQr.bmp"

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

End Module
