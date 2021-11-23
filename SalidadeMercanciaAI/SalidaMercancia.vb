Public Class SalidaMercancia

    Private SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private SBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION

    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        SBOApplication = oCatchingEvents.SBOApplication
        SBOCompany = oCatchingEvents.SBOCompany
    End Sub


    Public Function AddSalidaEntrada(ByVal DocNum As String, ByVal DocEntry As String)

        Dim stQueryH1, stQueryH2 As String
        Dim oRecSetH1, oRecSetH2 As SAPbobsCOM.Recordset
        Dim DocNumOIGE, DocNumOIGN, Comment As String

        oRecSetH1 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            stQueryH1 = "Select ""DocEntry"" from OIGE where ""Ref2""='" & DocNum & "'"
            oRecSetH1.DoQuery(stQueryH1)

            If oRecSetH1.RecordCount = 0 Then


                stQueryH2 = "Select T1.""ItemCode"",T1.""WhsCode"",T1.""Quantity""-T1.""U_CanReal"" as ""Quantity"",T0.""ObjType"",T1.""LineNum"",T2.""ManBtchNum"" From OPDN T0 Inner Join PDN1 T1 on T1.""DocEntry""=T0.""DocEntry"" Inner Join OITM T2 on T2.""ItemCode""=T1.""ItemCode"" Where T1.""Quantity"">T1.""U_CanReal"" AND T0.""DocEntry""=" & DocEntry
                oRecSetH2.DoQuery(stQueryH2)

                If oRecSetH2.RecordCount > 0 Then

                    DocNumOIGE = AddSalidaMercancia(oRecSetH2, DocNum, DocEntry)

                End If

                stQueryH2 = "Select T1.""ItemCode"",T1.""WhsCode"",T1.""U_CanReal""-T1.""Quantity"" as ""Quantity"",T0.""ObjType"",T1.""LineNum"",T2.""ManBtchNum"" From OPDN T0 Inner Join PDN1 T1 on T1.""DocEntry""=T0.""DocEntry"" Inner Join OITM T2 on T2.""ItemCode""=T1.""ItemCode"" Where T1.""Quantity""<T1.""U_CanReal"" AND T0.""DocEntry""=" & DocEntry
                oRecSetH2.DoQuery(stQueryH2)

                If oRecSetH2.RecordCount > 0 Then

                    DocNumOIGN = AddEntradaMercancia(oRecSetH2, DocNum, DocEntry)

                End If


                '//// arma mensaje para actualizar comment de entrada de mercacias en compras
                If DocNumOIGE <> Nothing Or DocNumOIGE <> "" Then

                    Comment = "Salida de mercancia: " & DocNumOIGE

                End If

                If DocNumOIGN <> Nothing Or DocNumOIGN <> "" Then

                    If Comment <> Nothing Or Comment <> "" Then

                        Comment = Comment & ", Entrada de mercancia: " & DocNumOIGN

                    Else

                        Comment = "Entrada de mercancia: " & DocNumOIGN

                    End If

                End If


                '//// actualiza la entrada de mercancia de compras
                If Comment <> Nothing Or Comment <> "" Then

                    UpdateEntradaMercanciaOC(DocEntry, Comment)

                End If


            End If

        Catch ex As Exception

            SBOApplication.MessageBox("Error en AddSalidaEntrada. " & ex.Message)

        End Try

    End Function


    Public Function AddSalidaMercancia(ByVal oRecSetH2 As SAPbobsCOM.Recordset, ByVal DocNum As String, ByVal DocEntry As String)

        Dim stQueryH3, stQueryH4 As String
        Dim oRecSetH3, oRecSetH4 As SAPbobsCOM.Recordset
        Dim oOIGE As SAPbobsCOM.Documents
        Dim ItemCode, WhsCode, Quantity, ObjType, LineNum, Lote, DocNumOIGE As String
        Dim llError As Long
        Dim lsError As String
        Dim AOIGE As Integer

        oRecSetH3 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH4 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oOIGE = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)

        Try

            oOIGE.DocDate = DateTime.Now
            oOIGE.Reference2 = DocNum

            oRecSetH2.MoveFirst()

            For i = 0 To oRecSetH2.RecordCount - 1

                ItemCode = oRecSetH2.Fields.Item("ItemCode").Value.ToString
                WhsCode = oRecSetH2.Fields.Item("WhsCode").Value.ToString
                Quantity = oRecSetH2.Fields.Item("Quantity").Value.ToString
                ObjType = oRecSetH2.Fields.Item("ObjType").Value.ToString
                LineNum = oRecSetH2.Fields.Item("LineNum").Value.ToString
                Lote = oRecSetH2.Fields.Item("ManBtchNum").Value.ToString

                oOIGE.Lines.ItemCode = ItemCode
                oOIGE.Lines.WarehouseCode = WhsCode
                oOIGE.Lines.Quantity = Quantity

                If Lote = "Y" Then

                    stQueryH4 = "Select T1.""BatchNum"" from IBT1 T1 where T1.""BaseType""=" & ObjType & " And T1.""BaseEntry""=" & DocEntry & " And T1.""BaseLinNum""=" & LineNum & " And T1.""ItemCode""='" & ItemCode & "'"
                    oRecSetH4.DoQuery(stQueryH4)

                    If oRecSetH4.RecordCount > 0 Then

                        oRecSetH4.MoveFirst()

                        oOIGE.Lines.BatchNumbers.BatchNumber = oRecSetH4.Fields.Item("BatchNum").Value.ToString
                        oOIGE.Lines.BatchNumbers.Quantity = Quantity

                        oOIGE.Lines.BatchNumbers.Add()

                    End If

                End If

                oOIGE.Lines.Add()
                oRecSetH2.MoveNext()

            Next

            If oOIGE.Add() <> 0 Then

                SBOCompany.GetLastError(llError, lsError)
                Err.Raise(-1, 1, lsError)

            Else

                AOIGE = SBOCompany.GetNewObjectKey().ToString()
                stQueryH3 = "Select ""DocNum"" from OIGE where ""DocEntry""=" & AOIGE
                oRecSetH3.DoQuery(stQueryH3)

                If oRecSetH3.RecordCount = 1 Then

                    DocNumOIGE = oRecSetH3.Fields.Item("DocNum").Value

                End If

            End If

            Return DocNumOIGE

        Catch ex As Exception

            SBOApplication.MessageBox("Error al crear Salida de Mercancia. " & ex.Message)

        End Try

    End Function


    Public Function AddEntradaMercancia(ByVal oRecSetH2 As SAPbobsCOM.Recordset, ByVal DocNum As String, ByVal DocEntry As String)

        Dim stQueryH3, stQueryH4 As String
        Dim oRecSetH3, oRecSetH4 As SAPbobsCOM.Recordset
        Dim oOIGN As SAPbobsCOM.Documents
        Dim ItemCode, WhsCode, Quantity, ObjType, LineNum, Lote, DocNumOIGN As String
        Dim llError As Long
        Dim lsError As String
        Dim AOIGN As Integer

        oRecSetH3 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH4 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oOIGN = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)

        Try

            oOIGN.DocDate = DateTime.Now
            oOIGN.Reference2 = DocNum

            oRecSetH2.MoveFirst()

            For i = 0 To oRecSetH2.RecordCount - 1

                ItemCode = oRecSetH2.Fields.Item("ItemCode").Value.ToString
                WhsCode = oRecSetH2.Fields.Item("WhsCode").Value.ToString
                Quantity = oRecSetH2.Fields.Item("Quantity").Value.ToString
                ObjType = oRecSetH2.Fields.Item("ObjType").Value.ToString
                LineNum = oRecSetH2.Fields.Item("LineNum").Value.ToString
                Lote = oRecSetH2.Fields.Item("ManBtchNum").Value.ToString

                oOIGN.Lines.ItemCode = ItemCode
                oOIGN.Lines.WarehouseCode = WhsCode
                oOIGN.Lines.Quantity = Quantity

                If Lote = "Y" Then

                    stQueryH4 = "Select T1.""BatchNum"" from IBT1 T1 where T1.""BaseType""=" & ObjType & " And T1.""BaseEntry""=" & DocEntry & " And T1.""BaseLinNum""=" & LineNum & " And T1.""ItemCode""='" & ItemCode & "'"
                    oRecSetH4.DoQuery(stQueryH4)

                    If oRecSetH4.RecordCount > 0 Then

                        oRecSetH4.MoveFirst()

                        oOIGN.Lines.BatchNumbers.BatchNumber = oRecSetH4.Fields.Item("BatchNum").Value.ToString
                        oOIGN.Lines.BatchNumbers.Quantity = Quantity

                        oOIGN.Lines.BatchNumbers.Add()

                    End If

                End If

                oOIGN.Lines.Add()
                oRecSetH2.MoveNext()

            Next

            If oOIGN.Add() <> 0 Then

                SBOCompany.GetLastError(llError, lsError)
                Err.Raise(-1, 1, lsError)

            Else

                AOIGN = SBOCompany.GetNewObjectKey().ToString()
                stQueryH3 = "Select ""DocNum"" from OIGN where ""DocEntry""=" & AOIGN
                oRecSetH3.DoQuery(stQueryH3)

                If oRecSetH3.RecordCount = 1 Then

                    DocNumOIGN = oRecSetH3.Fields.Item("DocNum").Value

                End If

            End If

            Return DocNumOIGN

        Catch ex As Exception

            SBOApplication.MessageBox("Error al crear Entrada de Mercancia. " & ex.Message)

        End Try

    End Function


    Public Function UpdateEntradaMercanciaOC(ByVal DocEntry As String, ByVal Comment As String)

        Dim oOPDN As SAPbobsCOM.Documents
        Dim llError As Long
        Dim lsError As String

        oOPDN = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)

        Try

            oOPDN.GetByKey(DocEntry)
            oOPDN.Comments = Comment

            If oOPDN.Update() <> 0 Then

                SBOCompany.GetLastError(llError, lsError)
                Err.Raise(-1, 1, lsError)

            Else

                SBOApplication.MessageBox("Se creo en automatico " & Comment)

            End If

        Catch ex As Exception

            SBOApplication.MessageBox("Error al actualizar la Entrada de Mercancia Compras. " & ex.Message)

        End Try

    End Function

End Class
