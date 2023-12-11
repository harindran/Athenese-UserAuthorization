
Namespace UserAuthorization
    Public Class ClsMIPLQC
        Public Const Formtype = "MIPLQC"
        Dim objform As SAPbouiCOM.Form
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strSQL As String
        Dim objRs As SAPbobsCOM.Recordset

        Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                objmatrix = objform.Items.Item("20").Specific
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If pVal.BeforeAction Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                    End Select
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                            If objform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
                            If pVal.ItemUID = "23" Or pVal.ItemUID = "13" Or pVal.ItemUID = "51" Then
                                strSQL = "select 1 from ""@AT_USRCFG"" where ""Code""='" & objaddon.objcompany.UserName & "'"
                                Dim status As String = objaddon.objglobalmethods.getSingleValue(strSQL)
                                If status = "" Then
                                    objaddon.objapplication.StatusBar.SetText("User Authorization Not Mapped for the User: " & objaddon.objcompany.UserName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    GoTo disable
                                End If
                                Dim tranEntry, LineTable As String
                                    If objform.Items.Item("8").Specific.selected.value = "G" Then
                                        tranEntry = objform.Items.Item("13B").Specific.string
                                        LineTable = "PDN1"
                                    ElseIf objform.Items.Item("8").Specific.selected.value = "P" Or objform.Items.Item("8").Specific.selected.value = "R" Then
                                        tranEntry = objform.Items.Item("51B").Specific.string
                                        LineTable = "IGN1"
                                    ElseIf objform.Items.Item("8").Specific.selected.value = "T" Then
                                        tranEntry = objform.Items.Item("23B").Specific.string
                                        LineTable = "WTR1"
                                    End If
                                    If tranEntry = "" Then Exit Sub
                                    strSQL = "Select count(*) as ""Status"" from"
                                    strSQL += vbCrLf + "(select distinct 1 from ""@AT_USRCFG"" T0 join ""@AT_USRCFG2"" T1 on T0.""Code""=T1.""Code"" where T0.""Code""='" & objaddon.objcompany.UserName & "' and ifnull(T1.""U_Assign"",'N')='N'"
                                    strSQL += vbCrLf + "and T1.""U_WhsCode"" in (Select distinct ""WhsCode"" from " & LineTable & " where ""ItemCode""<>'' and ""DocEntry""='" & tranEntry & "')"
                                    strSQL += vbCrLf + "Union all"
                                    strSQL += vbCrLf + "select distinct 1 from ""@AT_USRCFG"" T0 join ""@AT_USRCFG1"" T1 on T0.""Code""=T1.""Code"" where T0.""Code""='" & objaddon.objcompany.UserName & "' and ifnull(T1.""U_Assign"",'N')='N'"
                                    strSQL += vbCrLf + "and T1.""U_GrpCode"" in (Select A.""ItmsGrpCod"" from OITM A where A.""ItemCode"" in (Select distinct ""ItemCode"" from " & LineTable & " where ""ItemCode""<>'' and ""DocEntry""='" & tranEntry & "'))) A"
                                    status = objaddon.objglobalmethods.getSingleValue(strSQL)
                                    If CInt(status) <> 0 Then
disable:                            Try
                                        objaddon.objapplication.ActivateMenuItem("1281")
                                        objform.Items.Item("5").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                                    Catch ex As Exception
                                        objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        End Try
                                    End If
                                End If

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(BusinessObjectInfo.FormUID)
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim status As String
                If BusinessObjectInfo.BeforeAction = True Then
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            strSQL = "select 1 from ""@AT_USRCFG"" where ""Code""='" & objaddon.objcompany.UserName & "'"
                            status = objaddon.objglobalmethods.getSingleValue(strSQL)
                            If status = "" Then
                                objaddon.objapplication.StatusBar.SetText("User Authorization Not Mapped for the User: " & objaddon.objcompany.UserName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False
                                GoTo disable
                            End If
                    End Select
                Else
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            If BusinessObjectInfo.ActionSuccess = False Then Exit Sub

                            Dim odbdsHeader As SAPbouiCOM.DBDataSource
                            odbdsHeader = objform.DataSources.DBDataSources.Item("@MIPLQC")
                            Dim DocEntry As String = odbdsHeader.GetValue("DocEntry", 0)
                            strSQL = "Select count(*) as ""Status"" from"
                            strSQL += vbCrLf + "(select distinct 1 from ""@AT_USRCFG"" T0 join ""@AT_USRCFG2"" T1 on T0.""Code""=T1.""Code"" where T0.""Code""='" & objaddon.objcompany.UserName & "' and ifnull(T1.""U_Assign"",'N')='N'"
                            strSQL += vbCrLf + "and T1.""U_WhsCode"" in (Select distinct Case when ""U_AccQty"">0 then ""U_AccWhse"" when ""U_RejQty"">0 then ""U_RejWhse"" when ""U_RewQty"">0 then ""U_RewWhse"" End ""Whse"" from ""@MIPLQC1"" "
                            strSQL += vbCrLf + "where ""U_ItemCode""<>'' and (""U_AccWhse""<>'' or ""U_RejWhse""<>'' or ""U_RewWhse""<>'') and ""DocEntry""='" & DocEntry & "')"
                            strSQL += vbCrLf + "Union all"
                            strSQL += vbCrLf + "select distinct 1 from ""@AT_USRCFG"" T0 join ""@AT_USRCFG1"" T1 on T0.""Code""=T1.""Code"" where T0.""Code""='" & objaddon.objcompany.UserName & "' and ifnull(T1.""U_Assign"",'N')='N'"
                            strSQL += vbCrLf + "and T1.""U_GrpCode"" in (Select A.""ItmsGrpCod"" from OITM A where A.""ItemCode"" = (Select distinct ""U_ItemCode"" from ""@MIPLQC1"" "
                            strSQL += vbCrLf + "where ""U_ItemCode""<>'' and (""U_AccWhse""<>'' or ""U_RejWhse""<>'' or ""U_RewWhse""<>'') and ""U_ItemCode""=A.""ItemCode"" and ""DocEntry""='" & DocEntry & "'))) A"
                            status = objaddon.objglobalmethods.getSingleValue(strSQL)
                            If CInt(status) <> 0 Then
disable:                        Try
                                    objaddon.objapplication.ActivateMenuItem("1281")
                                    objform.Items.Item("5").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                                Catch ex As Exception
                                    objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                End Try
                            End If
                    End Select
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

    End Class

End Namespace
