﻿Public Class CatchingEvents

    Friend WithEvents SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Friend SBOCompany As SAPbobsCOM.Company '//OBJETO COMPAÑIA
    Friend csDirectory As String '//DIRECTORIO DONDE SE ENCUENTRAN LOS .SRF
    Dim DocNum As String

    Public Sub New()

        MyBase.New()
        SetAplication()
        SetConnectionContext()
        ConnectSBOCompany()

        setFilters()

    End Sub

    '//----- ESTABLECE LA COMUNICACION CON SBO
    Private Sub SetAplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String
        Try
            SboGuiApi = New SAPbouiCOM.SboGuiApi
            sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
            SboGuiApi.Connect(sConnectionString)
            SBOApplication = SboGuiApi.GetApplication()
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con la aplicación SBO " & ex.Message)
            System.Windows.Forms.Application.Exit()
            End '//termina aplicación
        End Try
    End Sub

    '//----- ESTABLECE EL CONTEXTO DE LA APLICACION
    Private Sub SetConnectionContext()
        Try
            SBOCompany = SBOApplication.Company.GetDICompany
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con el DI")
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End '//termina aplicación
            'Finally
        End Try
    End Sub

    '//----- CONEXION CON LA BASE DE DATOS
    Private Sub ConnectSBOCompany()
        Dim loRecSet As SAPbobsCOM.Recordset
        Try
            '//ESTABLECE LA CONEXION A LA COMPAÑIA
            csDirectory = My.Application.Info.DirectoryPath
            If (csDirectory = "") Then
                System.Windows.Forms.Application.Exit()
                End
            End If
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con la BD. " & ex.Message)
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End '//termina aplicación
        Finally
            loRecSet = Nothing
        End Try
    End Sub

    '//----- ESTABLECE FILTROS DE EVENTOS DE LA APLICACION
    Private Sub setFilters()
        Dim lofilter As SAPbouiCOM.EventFilter
        Dim lofilters As SAPbouiCOM.EventFilters

        Try
            lofilters = New SAPbouiCOM.EventFilters
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            lofilter.AddEx(143) '// FORMA Entrada de Mercancías
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            lofilter.AddEx(143) '// FORMA Entrada de Mercancías

            SBOApplication.SetFilter(lofilters)

        Catch ex As Exception
            SBOApplication.MessageBox("SetFilter: " & ex.Message)
        End Try

    End Sub

    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ''// METODOS PARA EVENTOS DE LA APLICACION
    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBOApplication.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                System.Windows.Forms.Application.Exit()
                End
        End Select

    End Sub

    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// METODOS PARA MANEJO DE EVENTOS ITEM
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub SBOApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBOApplication.ItemEvent

        If pVal.Before_Action = True And pVal.FormTypeEx <> "" Then
        Else
            If pVal.Before_Action = False And pVal.FormTypeEx <> "" Then
                Select Case pVal.FormTypeEx

                    Case 143                           '////// FORMA Entrada de Mercancías
                        frmOINVControllerAfter(FormUID, pVal)

                End Select
            End If
        End If

    End Sub


    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// CONTROLADOR DE EVENTOS FORMA Entrada de Mercancías
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub frmOINVControllerAfter(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent)

        'Dim oSNE As FrmtekEcommerce
        'Dim obotton As OINV
        Dim coForm As SAPbouiCOM.Form
        Dim stTabla As String
        Dim oDatatable As SAPbouiCOM.DBDataSource
        Dim Respuesta As Integer
        Dim ClientType, CardCode As String
        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset

        oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            Select Case pVal.EventType

                '///// Carga de formas
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    'obotton = New OINV
                    'obotton.addFormItems(FormUID)

                '//////Evento Presionar Item
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID

                        Case "1"

                            'stTabla = "OINV"
                            'coForm = SBOApplication.Forms.Item(FormUID)

                            'oDatatable = coForm.DataSources.DBDataSources.Item(stTabla)
                            'DocNum = oDatatable.GetValue("DocNum", 0)
                            'CardCode = oDatatable.GetValue("CardCode", 0)

                            'stQueryH = "Select ""GroupCode"" from OCRD where ""CardCode""='" & CardCode & "'"
                            'oRecSetH.DoQuery(stQueryH)

                            'If oRecSetH.RecordCount = 1 Then

                            '    ClientType = oRecSetH.Fields.Item("GroupCode").Value

                            '    If ClientType = 121 Then

                            '        Respuesta = SBOApplication.MessageBox("¿Deseas ReFacturar este documento?", Btn1Caption:="Si", Btn2Caption:="No")

                            '        If Respuesta = 1 Then

                            '            'oSNE = New FrmtekEcommerce
                            '            'oSNE.openForm(csDirectory)

                            '        End If

                            '    Else

                            '        SBOApplication.MessageBox("No se puede Refacturar para este cliente")

                            '    End If

                            'End If

                    End Select

            End Select

        Catch ex As Exception
            SBOApplication.MessageBox("Error en el evento sobre Forma Facturacion Clientes. " & ex.Message)
        Finally
            'oPO = Nothing
        End Try
    End Sub

End Class
