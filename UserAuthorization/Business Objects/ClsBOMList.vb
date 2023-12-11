Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports SAPbobsCOM
Imports SAPbouiCOM.Framework
Namespace UserAuthorization
    Public Class ClsBOMList
        Public Const Formtype = "9999"
        Dim objform As SAPbouiCOM.Form
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strSQL As String
        Dim objRs As SAPbobsCOM.Recordset

        Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)

                If pVal.BeforeAction Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                        Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                        Case SAPbouiCOM.BoEventTypes.et_CLICK

                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    End Select
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            BOMListFormCount += 1
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If BOMListFormCount = 1 And pVal.ItemUID <> "" Then
                                strSQL = "select 1 from ""@AT_USRCFG"" where ""Code""='" & objaddon.objcompany.UserName & "'"
                                Dim status As String = objaddon.objglobalmethods.getSingleValue(strSQL)
                                If status = "" Then
                                    objaddon.objapplication.StatusBar.SetText("User Authorization Not Mapped for the User: " & objaddon.objcompany.UserName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Exit Sub
                                End If
                                objmatrix = objform.Items.Item("7").Specific
                                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                For i As Integer = 1 To objmatrix.VisualRowCount
                                    If objaddon.HANA Then
                                        strSQL = "Select distinct 1 as ""Status"" from OITM where ""ItmsGrpCod"" in (select T1.""U_GrpCode"""
                                        strSQL += vbCrLf + "from ""@AT_USRCFG"" T0 join ""@AT_USRCFG1"" T1 on T0.""Code""=T1.""Code"" where T0.""Code""='" & objaddon.objcompany.UserName & "' and ifnull(T1.""U_Assign"",'N')='Y'"
                                        strSQL += vbCrLf + "and T1.""U_GrpCode"" in (Select ""ItmsGrpCod"" from OITM where ""ItemCode"" in (select ""U_ItemCode"" from ""@CT_PF_OBOM"" where ""ItemCode""='" & objmatrix.Columns.Item("U_ItemCode").Cells.Item(i).Specific.String & "'))) "
                                        strSQL += vbCrLf + "and ""ItemCode"" in  (select ""U_ItemCode"" from ""@CT_PF_OBOM"" where ""ItemCode""='" & objmatrix.Columns.Item("U_ItemCode").Cells.Item(i).Specific.String & "') "
                                    Else
                                        strSQL = "Select distinct 1 as Status from OITM where ItmsGrpCod in (select T1.U_GrpCode"
                                        strSQL += vbCrLf + "from [@AT_USRCFG[ T0 join [@AT_USRCFG1] T1 on T0.Code=T1.Code where T0.Code='" & objaddon.objcompany.UserName & "' and isnull(T1.U_Assign,'N')='Y'"
                                        strSQL += vbCrLf + "and T1.U_GrpCode in (Select ItmsGrpCod from OITM where ItemCode in (select U_ItemCode from [@CT_PF_OBOM] where ItemCode='" & objmatrix.Columns.Item("U_ItemCode").Cells.Item(i).Specific.String & "'))) "
                                        strSQL += vbCrLf + "and ItemCode in (select U_ItemCode from [@CT_PF_OBOM] where ItemCode='" & objmatrix.Columns.Item("U_ItemCode").Cells.Item(i).Specific.String & "') "
                                    End If
                                    If i = 1 Then objaddon.objapplication.StatusBar.SetText("Details sorting based on User Authorization.Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    objRs.DoQuery(strSQL)
                                    If objRs.RecordCount = 0 Then
                                        objmatrix.ClearRowData(i)
                                    End If
                                Next
                                BOMListFormCount = 0
                                objaddon.objapplication.StatusBar.SetText("Details sorted...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                objmatrix.Columns.Item("U_ItemCode").TitleObject.Click(SAPbouiCOM.BoCellClickType.ct_Double)
                                objmatrix.Columns.Item("U_ItemCode").TitleObject.Click(SAPbouiCOM.BoCellClickType.ct_Double)
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
                If BusinessObjectInfo.BeforeAction = True Then
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    End Select
                Else
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    End Select
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

    End Class
End Namespace
