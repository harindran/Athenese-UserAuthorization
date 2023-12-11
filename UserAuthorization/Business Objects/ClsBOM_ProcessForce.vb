Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports SAPbobsCOM
Imports SAPbouiCOM.Framework
Namespace UserAuthorization
    Public Class ClsBOM_ProcessForce
        Public Const Formtype = "CT_PF_OBOMCode"
        Dim objform As SAPbouiCOM.Form
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strSQL As String
        Dim objRs As SAPbobsCOM.Recordset

        Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                objmatrix = objform.Items.Item("Items").Specific
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If pVal.BeforeAction Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                            If pVal.ItemUID = "90" Then 'Or (pVal.ItemUID = "Items" And pVal.ColUID = "col_122")
                                Try
                                    Dim oCFL As SAPbouiCOM.ChooseFromList
                                    oCFL = objform.ChooseFromLists.Item("CFL_10") '90 -CFL_10, Matrix col_122- CFL_13 

                                    Dim oConds As SAPbouiCOM.Conditions
                                    Dim oCond As SAPbouiCOM.Condition
                                    Dim oEmptyConds As New SAPbouiCOM.Conditions
                                    oCFL.SetConditions(oEmptyConds)
                                    oConds = oCFL.GetConditions()
                                    'If ToWhseInInput = "Y" Then
                                    If objaddon.HANA Then
                                        strSQL = "select T1.""U_WhsCode"" from ""@AT_USRCFG"" T0 join ""@AT_USRCFG2"" T1 on T0.""Code""=T1.""Code"" where T0.""Code""='" & objaddon.objcompany.UserName & "' and T1.""U_Assign""='Y'"
                                    Else
                                        strSQL = "select T1.U_WhsCode from @AT_USRCFG T0 join @AT_USRCFG2 T1 on T0.Code=T1.Code where T0.Code='" & objaddon.objcompany.UserName & "' and T1.U_Assign='Y'"
                                    End If
                                    objRs.DoQuery(strSQL)
                                    If objRs.RecordCount > 0 Then
                                        For Val As Integer = 0 To objRs.RecordCount - 1
                                            If Val = 0 Then
                                                oCond = oConds.Add()
                                                oCond.Alias = "WhsCode"
                                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                                oCond.CondVal = objRs.Fields.Item("U_WhsCode").Value.ToString
                                            Else
                                                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                                                oCond = oConds.Add()
                                                oCond.Alias = "WhsCode"
                                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                                oCond.CondVal = objRs.Fields.Item("U_WhsCode").Value.ToString
                                            End If
                                            objRs.MoveNext()
                                        Next
                                    Else
                                        oCond = oConds.Add()
                                        oCond.Alias = "WhsCode"
                                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                        oCond.CondVal = -1
                                    End If
                                    oCFL.SetConditions(oConds)
                                Catch ex As Exception
                                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                End Try
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                        Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                        Case SAPbouiCOM.BoEventTypes.et_CLICK

                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    End Select
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                            Try

                            Catch ex As Exception
                            End Try
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

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
                            strSQL = "select 1 from ""@AT_USRCFG"" where ""Code""='" & objaddon.objcompany.UserName & "'"
                            Dim status As String = objaddon.objglobalmethods.getSingleValue(strSQL)
                            If status = "" Then
                                objaddon.objapplication.StatusBar.SetText("User Authorization Not Mapped for the User: " & objaddon.objcompany.UserName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                GoTo disable
                            End If
                            Dim odbdsHeader As SAPbouiCOM.DBDataSource
                            odbdsHeader = objform.DataSources.DBDataSources.Item("@CT_PF_OBOM")
                            Dim code As String = odbdsHeader.GetValue("Code", 0)
                            strSQL = "Select count(*) as ""Status"" from"
                            strSQL += vbCrLf + "(select distinct 1 from ""@AT_USRCFG"" T0 join ""@AT_USRCFG2"" T1 on T0.""Code""=T1.""Code"" where T0.""Code""='" & objaddon.objcompany.UserName & "' and T1.""U_Assign""='Y'"
                            strSQL += vbCrLf + "and T1.""U_WhsCode"" in (select ""U_WhsCode"" from ""@CT_PF_OBOM"" where ""Code""='" & odbdsHeader.GetValue("Code", 0) & "')"
                            strSQL += vbCrLf + "Union all"
                            strSQL += vbCrLf + "select distinct 1 from ""@AT_USRCFG"" T0 join ""@AT_USRCFG1"" T1 on T0.""Code""=T1.""Code"" where T0.""Code""='" & objaddon.objcompany.UserName & "' and T1.""U_Assign""='Y'"
                            strSQL += vbCrLf + "and T1.""U_GrpCode"" in (Select ""ItmsGrpCod"" from OITM where ""ItemCode""=(select ""U_ItemCode"" from ""@CT_PF_OBOM"" where ""Code""='" & odbdsHeader.GetValue("Code", 0) & "'))) A"
                            status = objaddon.objglobalmethods.getSingleValue(strSQL)
                            If CInt(status) < 2 Then
                                'objaddon.objapplication.ActivateMenuItem("1281")
disable:                        objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            End If
                            'If objform.Items.Item("5").Specific.String = "130" Then
                            'End If
                    End Select
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

    End Class
End Namespace
