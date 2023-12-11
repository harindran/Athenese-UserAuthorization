Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports SAPbobsCOM
Imports SAPbouiCOM.Framework
Namespace UserAuthorization
    Public Class ClsManfactOrder_ProcessForce
        Public Const Formtype = "CT_PF_ManufacOrd"
        Dim objform As SAPbouiCOM.Form
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strSQL As String
        Dim objRs As SAPbobsCOM.Recordset

        Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                objmatrix = objform.Items.Item("Items").Specific

                If pVal.BeforeAction Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                            If pVal.ItemUID = "11" And objform.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then 'Or (pVal.ItemUID = "Items" And pVal.ColUID = "col_122")
                                Try
                                    Dim oCFL As SAPbouiCOM.ChooseFromList
                                    oCFL = objform.ChooseFromLists.Item("CFL_OMOR") 'ParentMorCfl
                                    Dim oConds As SAPbouiCOM.Conditions
                                    Dim oCond As SAPbouiCOM.Condition
                                    Dim oEmptyConds As New SAPbouiCOM.Conditions
                                    oCFL.SetConditions(oEmptyConds)
                                    oConds = oCFL.GetConditions()
                                    objRs = Nothing
                                    objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    If objaddon.HANA Then
                                        strSQL = "Select distinct ""ItemCode"",""ItmsGrpCod"" from OITM where ""ItmsGrpCod"" in (select T1.""U_GrpCode"""
                                        strSQL += vbCrLf + "from ""@AT_USRCFG"" T0 join ""@AT_USRCFG1"" T1 on T0.""Code""=T1.""Code"" where T0.""Code""='" & objaddon.objcompany.UserName & "' and T1.""U_Assign""='Y'"
                                        strSQL += vbCrLf + "and T1.""U_GrpCode"" in (Select ""ItmsGrpCod"" from OITM where ""ItemCode"" in (select ""U_ItemCode"" from ""@CT_PF_OBOM""))) "
                                        strSQL += vbCrLf + "and ""ItemCode"" in  (select ""U_ItemCode"" from ""@CT_PF_OBOM"") "
                                    Else
                                        strSQL = "Select distinct ItemCode,ItmsGrpCod from OITM where ItmsGrpCod in (select T1.U_GrpCode"
                                        strSQL += vbCrLf + "from [@AT_USRCFG] T0 join [@AT_USRCFG1] T1 on T0.Code=T1.Code where T0.Code='" & objaddon.objcompany.UserName & "' and T1.U_Assign='Y'"
                                        strSQL += vbCrLf + "and T1.U_GrpCode in (Select ItmsGrpCod from OITM where ItemCode in (select U_ItemCode from [@CT_PF_OBOM]))) "
                                        strSQL += vbCrLf + "and ItemCode in  (select U_ItemCode from [@CT_PF_OBOM]) "
                                    End If
                                    objRs.DoQuery(strSQL)
                                    If objRs.RecordCount > 0 Then
                                        For Val As Integer = 0 To objRs.RecordCount - 1
                                            If Val = 0 Then
                                                oCond = oConds.Add()
                                                oCond.Alias = "U_ItemCode"
                                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                                oCond.CondVal = objRs.Fields.Item("ItemCode").Value.ToString
                                            Else
                                                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                                                oCond = oConds.Add()
                                                oCond.Alias = "U_ItemCode"
                                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                                oCond.CondVal = objRs.Fields.Item("ItemCode").Value.ToString
                                            End If
                                            objRs.MoveNext()
                                        Next
                                    Else
                                        oCond = oConds.Add()
                                        oCond.Alias = "U_ItemCode"
                                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                        oCond.CondVal = "-1"
                                    End If
                                    oCFL.SetConditions(oConds)
                                Catch ex As Exception
                                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                End Try
                            ElseIf pVal.ItemUID = "31" And objform.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                Try
                                    Dim oCFL As SAPbouiCOM.ChooseFromList
                                    oCFL = objform.ChooseFromLists.Item("CFL_WHS")
                                    Dim oConds As SAPbouiCOM.Conditions
                                    Dim oCond As SAPbouiCOM.Condition
                                    Dim oEmptyConds As New SAPbouiCOM.Conditions
                                    oCFL.SetConditions(oEmptyConds)
                                    oConds = oCFL.GetConditions()
                                    If objaddon.HANA Then
                                        strSQL = "select T1.""U_WhsCode"" from ""@AT_USRCFG"" T0 join ""@AT_USRCFG2"" T1 on T0.""Code""=T1.""Code"" where T0.""Code""='" & objaddon.objcompany.UserName & "' and T1.""U_Assign""='Y'"
                                    Else
                                        strSQL = "select T1.U_WhsCode from @AT_USRCFG T0 join @AT_USRCFG2 T1 on T0.Code=T1.Code where T0.Code='" & objaddon.objcompany.UserName & "' and T1.U_Assign='Y'"
                                    End If
                                    Dim objRs1 As SAPbobsCOM.Recordset
                                    objRs1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    objRs1.DoQuery(strSQL)
                                    If objRs1.RecordCount > 0 Then
                                        objRs1.MoveFirst()
                                        For Val As Integer = 0 To objRs1.RecordCount - 1
                                            If Val = 0 Then
                                                oCond = oConds.Add()
                                                oCond.Alias = "WhsCode"
                                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                                oCond.CondVal = objRs1.Fields.Item("U_WhsCode").Value
                                            Else
                                                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                                                oCond = oConds.Add()
                                                oCond.Alias = "WhsCode"
                                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                                oCond.CondVal = objRs1.Fields.Item("U_WhsCode").Value
                                            End If
                                            objRs1.MoveNext()
                                        Next
                                    Else
                                        oCond = oConds.Add()
                                        oCond.Alias = "WhsCode"
                                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                        oCond.CondVal = "-1"
                                    End If
                                    oCFL.SetConditions(oConds)
                                Catch ex As Exception
                                    'SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                End Try
                            End If


                        Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                            If pVal.ItemUID = "11" Then
                                objRs = Nothing
                                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                strSQL = "select 1 from ""@AT_USRCFG"" where ""Code""='" & objaddon.objcompany.UserName & "'"
                                Dim status As String = objaddon.objglobalmethods.getSingleValue(strSQL)
                                If status = "" Then
                                    objaddon.objapplication.StatusBar.SetText("User Authorization Not Mapped for the User: " & objaddon.objcompany.UserName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    GoTo disable
                                End If
                                Dim ItemCode As String = objform.Items.Item("11").Specific.String
                                If objaddon.HANA Then
                                    strSQL = "Select distinct 1 as ""Status"" from OITM where ""ItmsGrpCod"" in (select T1.""U_GrpCode"""
                                    strSQL += vbCrLf + "from ""@AT_USRCFG"" T0 join ""@AT_USRCFG1"" T1 on T0.""Code""=T1.""Code"" where T0.""Code""='" & objaddon.objcompany.UserName & "' and T1.""U_Assign""='Y'"
                                    strSQL += vbCrLf + "and T1.""U_GrpCode"" in (Select ""ItmsGrpCod"" from OITM where ""ItemCode"" in (select ""U_ItemCode"" from ""@CT_PF_OBOM"" where ""ItemCode""='" & ItemCode & "'))) "
                                    strSQL += vbCrLf + "and ""ItemCode"" in  (select ""U_ItemCode"" from ""@CT_PF_OBOM"" where ""ItemCode""='" & ItemCode & "') "
                                Else
                                    strSQL = "Select distinct 1 as Status from OITM where ItmsGrpCod in (select T1.U_GrpCode"
                                    strSQL += vbCrLf + "from [@5AT_USRCFG] T0 join [@AT_USRCFG1] T1 on T0.Code=T1.Code where T0.Code='" & objaddon.objcompany.UserName & "' and T1.U_Assign='Y'"
                                    strSQL += vbCrLf + "and T1.U_GrpCode in (Select ItmsGrpCod from OITM where ItemCode in (select U_ItemCode from [@CT_PF_OBOM] where ItemCode='" & ItemCode & "'))) "
                                    strSQL += vbCrLf + "and ItemCode in (select U_ItemCode from [@CT_PF_OBOM] where ItemCode='" & ItemCode & "') "
                                End If
                                objRs.DoQuery(strSQL)
                                If objRs.RecordCount = 0 Then
disable:                            objmatrix.Clear()
                                End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_CLICK

                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                    End Select
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                            Try

                            Catch ex As Exception
                            End Try
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            If pVal.ActionSuccess Then
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                        Case SAPbouiCOM.BoEventTypes.et_CLICK

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
                If BusinessObjectInfo.BeforeAction = True Then
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    End Select
                Else
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            If BusinessObjectInfo.ActionSuccess = False Then Exit Sub
                            strSQL = "select 1 from ""@AT_USRCFG"" where ""Code""='" & objaddon.objcompany.UserName & "'"
                            Dim status As String = objaddon.objglobalmethods.getSingleValue(strSQL)
                            If status = "" Then
                                objaddon.objapplication.StatusBar.SetText("User Authorization Not Mapped for the User: " & objaddon.objcompany.UserName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                GoTo disable
                            End If
                            Dim odbdsHeader As SAPbouiCOM.DBDataSource
                            odbdsHeader = objform.DataSources.DBDataSources.Item("@CT_PF_OMOR")
                            Dim DocEntry As String = odbdsHeader.GetValue("DocEntry", 0)
                            strSQL = "Select count(*) as ""Status"" from"
                            strSQL += vbCrLf + "(select distinct 1 from ""@AT_USRCFG"" T0 join ""@AT_USRCFG2"" T1 on T0.""Code""=T1.""Code"" where T0.""Code""='" & objaddon.objcompany.UserName & "' and T1.""U_Assign""='Y'"
                            strSQL += vbCrLf + "and T1.""U_WhsCode"" in (select ""U_Warehouse"" from ""@CT_PF_OMOR"" where ""DocEntry""='" & odbdsHeader.GetValue("DocEntry", 0) & "')"
                            strSQL += vbCrLf + "Union all"
                            strSQL += vbCrLf + "select distinct 1 from ""@AT_USRCFG"" T0 join ""@AT_USRCFG1"" T1 on T0.""Code""=T1.""Code"" where T0.""Code""='" & objaddon.objcompany.UserName & "' and T1.""U_Assign""='Y'"
                            strSQL += vbCrLf + "and T1.""U_GrpCode"" in (Select ""ItmsGrpCod"" from OITM where ""ItemCode""=(select ""U_ItemCode"" from ""@CT_PF_OMOR"" where ""DocEntry""='" & odbdsHeader.GetValue("DocEntry", 0) & "'))) A"
                            status = objaddon.objglobalmethods.getSingleValue(strSQL)
                            If CInt(status) < 2 Then
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
