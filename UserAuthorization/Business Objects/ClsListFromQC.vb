Namespace UserAuthorization
    Public Class ClsListFromQC

        Public Const Formtype = "OCFL"
        Dim objform As SAPbouiCOM.Form
        Dim objMatrix As SAPbouiCOM.Matrix
        Dim strSQL As String
        Dim objRs As SAPbobsCOM.Recordset

        Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                If pVal.BeforeAction Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN


                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    End Select
                Else
                    Select Case pVal.EventType

                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            QCListFormCount += 1
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If QCListFormCount = 1 And pVal.ItemUID <> "" And pVal.ItemUID <> "2" Then
                                objMatrix = objform.Items.Item("3").Specific
                                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                For i As Integer = 1 To objMatrix.VisualRowCount
                                    If objaddon.HANA Then
                                        strSQL = "select T0.""Code"",(Select Case when ifnull(T1.""U_Assign"",'N')='N' then 'N' Else 'Y' End from ""@AT_USRCFG1"" T1 where T1.""Code""=T0.""Code"""
                                        strSQL += vbCrLf + "and T1.""U_GrpCode"" = (Select ""ItmsGrpCod"" from OITM where ""ItemCode"" ='" & objMatrix.Columns.Item("3A").Cells.Item(i).Specific.String & "')) ""ItemGrpStatus"","
                                        strSQL += vbCrLf + "(Select Case when ifnull(T1.""U_Assign"",'N')='N' then 'N' Else 'Y' End from ""@AT_USRCFG2"" T1 where T1.""Code""=T0.""Code"" and T1.""U_WhsCode"" ='" & objMatrix.Columns.Item("6A").Cells.Item(i).Specific.String & "') ""WhseStatus"""
                                        strSQL += vbCrLf + "from ""@AT_USRCFG"" T0 where T0.""Code""='" & objaddon.objcompany.UserName & "'"
                                    Else
                                        strSQL = "select T0.Code,(Select Case when isnull(T1.U_Assign,'N')='N' then 'N' Else 'Y' End from [@AT_USRCFG1] T1 where T1.Code=T0.Code"
                                        strSQL += vbCrLf + "and T1.U_GrpCode = (Select ItmsGrpCod from OITM where ItemCode ='" & objMatrix.Columns.Item("3A").Cells.Item(i).Specific.String & "')) ItemGrpStatus,"
                                        strSQL += vbCrLf + "(Select Case when isnull(T1.U_Assign,'N')='N' then 'N' Else 'Y' End from [@AT_USRCFG2] T1 where T1.Code=T0.Code and T1.U_WhsCode ='" & objMatrix.Columns.Item("6A").Cells.Item(i).Specific.String & "') WhseStatus"
                                        strSQL += vbCrLf + "from [@AT_USRCFG] T0 where T0.Code='" & objaddon.objcompany.UserName & "'"
                                    End If
                                    objRs.DoQuery(strSQL)
                                    If i = 1 Then objaddon.objapplication.StatusBar.SetText("Details sorting based on User Authorization.Please wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    If objRs.RecordCount > 0 Then
                                        If objRs.Fields.Item("WhseStatus").Value.ToString = "N" Or objRs.Fields.Item("ItemGrpStatus").Value.ToString = "N" Then
                                            objMatrix.ClearRowData(i)
                                        End If
                                    End If
                                Next
                                QCListFormCount = 0
                                objaddon.objapplication.StatusBar.SetText("Details sorted...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                objMatrix.Columns.Item("4").TitleObject.Click(SAPbouiCOM.BoCellClickType.ct_Double)
                                objMatrix.Columns.Item("4").TitleObject.Click(SAPbouiCOM.BoCellClickType.ct_Double)
                            End If




                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub


    End Class
End Namespace

