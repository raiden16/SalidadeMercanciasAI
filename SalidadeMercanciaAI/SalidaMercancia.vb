Public Class SalidaMercancia

    Private SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private SBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION

    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        SBOApplication = oCatchingEvents.SBOApplication
        SBOCompany = oCatchingEvents.SBOCompany
    End Sub


    Public Function AddSalidaMercancia(ByVal DocNum As String, ByVal DocEntry As String)

        Dim stQueryH1, stQueryH2, stQueryH3, stQueryH4 As String
        Dim oRecSetH1, oRecSetH2, oRecSetH3, oRecSetH4 As SAPbobsCOM.Recordset
        Dim oOIGE As SAPbobsCOM.Documents
        Dim ItemCode, WhsCode, Quantity, ObjType, LineNum, Lote, DocNumOIGE As String
        Dim llError As Long
        Dim lsError As String
        Dim AOIGE As Integer

        oRecSetH1 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH3 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH4 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oOIGE = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)

        Try

            stQueryH1 = "Select ""DocEntry"" from OIGE where ""Ref2""='" & DocNum & "'"
            oRecSetH1.DoQuery(stQueryH1)

            If oRecSetH1.RecordCount = 0 Then

                stQueryH2 = "Select T1.""ItemCode"",T1.""WhsCode"",T1.""Quantity""-T1.""U_CanReal"" as ""Quantity"",T0.""ObjType"",T1.""LineNum"",T2.""ManBtchNum"" From OPDN T0 Inner Join PDN1 T1 on T1.""DocEntry""=T0.""DocEntry"" Inner Join OITM T2 on T2.""ItemCode""=T1.""ItemCode"" Where T1.""Quantity"">T1.""U_CanReal"" AND T0.""DocEntry""=" & DocEntry
                oRecSetH2.DoQuery(stQueryH2)

                If oRecSetH2.RecordCount > 0 Then

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
                            UpdateEntradaMercancia(DocEntry, DocNumOIGE)

                        End If

                    End If

                End If

            End If

        Catch ex As Exception

            SBOApplication.MessageBox("Error al crear Salida de Mercancia. " & ex.Message)

        End Try

    End Function


    Public Function UpdateEntradaMercancia(ByVal DocEntry As String, ByVal DocNumOIGE As String)

        Dim oOPDN As SAPbobsCOM.Documents
        Dim llError As Long
        Dim lsError As String

        oOPDN = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)

        Try

            oOPDN.GetByKey(DocEntry)
            oOPDN.Comments = DocNumOIGE

            If oOPDN.Update() <> 0 Then

                SBOCompany.GetLastError(llError, lsError)
                Err.Raise(-1, 1, lsError)

            Else

                SBOApplication.MessageBox("Se creo en automatico una salida de mercancia con el numero " & DocNumOIGE)

            End If

        Catch ex As Exception

            SBOApplication.MessageBox("Error al actualizar la Entrada de Mercancia. " & ex.Message)

        End Try

    End Function

End Class
