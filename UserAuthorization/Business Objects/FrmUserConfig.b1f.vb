Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace UserAuthorization
    <FormAttribute("USRCONFIG", "Business Objects/FrmUserConfig.b1f")>
    Friend Class FrmUserConfig
        Inherits UserFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg
        Dim objRs As SAPbobsCOM.Recordset
        Dim FormCount As Integer = 0
        Dim strsql As String
        Private WithEvents objDBHeader, odbdsDetails, odbdsDetails1 As SAPbouiCOM.DBDataSource
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.StaticText0 = CType(Me.GetItem("lblcode").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("txtcode").Specific, SAPbouiCOM.EditText)
            Me.EditText1 = CType(Me.GetItem("txtname").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("lblcrteon").Specific, SAPbouiCOM.StaticText)
            Me.EditText2 = CType(Me.GetItem("txtcrteon").Specific, SAPbouiCOM.EditText)
            Me.StaticText3 = CType(Me.GetItem("lbldept").Specific, SAPbouiCOM.StaticText)
            Me.EditText3 = CType(Me.GetItem("txtdept").Specific, SAPbouiCOM.EditText)
            Me.StaticText4 = CType(Me.GetItem("lblcrteby").Specific, SAPbouiCOM.StaticText)
            Me.EditText4 = CType(Me.GetItem("txtcrteby").Specific, SAPbouiCOM.EditText)
            Me.StaticText5 = CType(Me.GetItem("lblupdton").Specific, SAPbouiCOM.StaticText)
            Me.EditText5 = CType(Me.GetItem("txtupdton").Specific, SAPbouiCOM.EditText)
            Me.StaticText6 = CType(Me.GetItem("lblupdtby").Specific, SAPbouiCOM.StaticText)
            Me.EditText6 = CType(Me.GetItem("txtupdtby").Specific, SAPbouiCOM.EditText)
            Me.StaticText7 = CType(Me.GetItem("lblrem").Specific, SAPbouiCOM.StaticText)
            Me.EditText7 = CType(Me.GetItem("txtrem").Specific, SAPbouiCOM.EditText)
            Me.Folder0 = CType(Me.GetItem("fldritmgrp").Specific, SAPbouiCOM.Folder)
            Me.Folder1 = CType(Me.GetItem("fldrwhse").Specific, SAPbouiCOM.Folder)
            Me.Matrix0 = CType(Me.GetItem("mtxitmgrp").Specific, SAPbouiCOM.Matrix)
            Me.Matrix1 = CType(Me.GetItem("mtxwhse").Specific, SAPbouiCOM.Matrix)
            Me.LinkedButton0 = CType(Me.GetItem("lnkcode").Specific, SAPbouiCOM.LinkedButton)
            Me.EditText8 = CType(Me.GetItem("txtuserid").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler LoadAfter, AddressOf Me.Form_LoadAfter
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter
            AddHandler ResizeAfter, AddressOf Me.Form_ResizeAfter

        End Sub

#Region "Fields"
        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText
        Private WithEvents EditText4 As SAPbouiCOM.EditText
        Private WithEvents StaticText5 As SAPbouiCOM.StaticText
        Private WithEvents EditText5 As SAPbouiCOM.EditText
        Private WithEvents StaticText6 As SAPbouiCOM.StaticText
        Private WithEvents EditText6 As SAPbouiCOM.EditText
        Private WithEvents StaticText7 As SAPbouiCOM.StaticText
        Private WithEvents EditText7 As SAPbouiCOM.EditText
        Private WithEvents Folder0 As SAPbouiCOM.Folder
        Private WithEvents Folder1 As SAPbouiCOM.Folder
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents Matrix1 As SAPbouiCOM.Matrix
#End Region

        Private Sub OnCustomInitialize()
            Try
                'objform.Freeze(True)
                'objform.Items.Item("mtxitmgrp").Enabled = False
                'objform.Items.Item("mtxwhse").Enabled = False
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtcode", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtname", False, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtdept", False, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtcrteon", False, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtcrteby", False, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtupdton", False, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtupdtby", False, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "mtxitmgrp", True, False, True)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "mtxwhse", True, False, True)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtuserid", False, True, False)
                objform.EnableMenu("1283", False)
                objDBHeader = objform.DataSources.DBDataSources.Item("@AT_USRCFG")
                objDBHeader.SetValue("Code", 0, USERNAME)
                If objform.Items.Item("txtcode").Specific.String = "" Then objform.Items.Item("txtcode").Specific.String = USERNAME 'objaddon.objcompany.UserName
                objform.Items.Item("txtname").Specific.String = objaddon.objglobalmethods.getSingleValue("Select ""U_NAME"" from OUSR where ""USER_CODE""='" & objaddon.objcompany.UserName & "'")
                objDBHeader.SetValue("U_UserId", 0, objaddon.objglobalmethods.getSingleValue("Select ""USERID"" from OUSR where ""USER_CODE""='" & objaddon.objcompany.UserName & "'"))
                objform.Items.Item("txtdept").Specific.String = objaddon.objglobalmethods.getSingleValue("Select (Select ""Name"" from OUDP where ""Code""=""Department"") ""Dept"" from OUSR where ""USER_CODE""='" & objaddon.objcompany.UserName & "'")

                If USERNAME <> "" Then
                    If objaddon.HANA Then
                        strsql = objaddon.objglobalmethods.getSingleValue("select 1 ""Status"" from ""@AT_USRCFG"" where ""Code""='" & USERNAME & "' ")
                    Else
                        strsql = objaddon.objglobalmethods.getSingleValue("select 1 Status from [@AT_USRCFG] where Code='" & USERNAME & "' ")
                    End If
                    If strsql = "1" Then
                        objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        EditText0.Value = USERNAME
                        objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Folder0.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Exit Sub
                    End If
                End If
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then objDBHeader.SetValue("U_CreateBy", 0, objaddon.objcompany.UserName)
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then objDBHeader.SetValue("U_CreateOn", 0, objaddon.objglobalmethods.getSingleValue("Select TO_VARCHAR(Current_TIMESTAMP,'DD/MM/YYYY HH:MM:SS AM') ""Created On"" from Dummy"))
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then objDBHeader.SetValue("U_UpdateBy", 0, objaddon.objcompany.UserName)
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then objDBHeader.SetValue("U_UpdateOn", 0, objaddon.objglobalmethods.getSingleValue("Select TO_VARCHAR(Current_TIMESTAMP,'DD/MM/YYYY HH:MM:SS AM') ""Updated On"" from Dummy"))
                'EditText8.Item.Visible = False
                LoadDetails()

                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Private Sub LoadDetails()
            Try
                If EditText0.Value = "" Then Exit Sub
                objaddon.objapplication.StatusBar.SetText("Loading Item Group & Warehouse Details. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objDBHeader = objform.DataSources.DBDataSources.Item("@AT_USRCFG")
                odbdsDetails = objform.DataSources.DBDataSources.Item("@AT_USRCFG1")
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objform.Freeze(True)
                If objaddon.HANA Then
                    objRs.DoQuery("select ROW_NUMBER() OVER () AS ""LineId"" ,""ItmsGrpCod"",""ItmsGrpNam"" from OITB")
                Else
                    objRs.DoQuery("select ROW_NUMBER() OVER () AS LineId ,ItmsGrpCod,ItmsGrpNam from OITB")
                End If
                Matrix0.Clear()
                odbdsDetails.Clear()
                Folder0.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                While Not objRs.EoF
                    Matrix0.AddRow()
                    Matrix0.GetLineData(Matrix0.VisualRowCount)
                    odbdsDetails.SetValue("LineId", 0, objRs.Fields.Item("LineId").Value.ToString)
                    odbdsDetails.SetValue("U_GrpCode", 0, objRs.Fields.Item("ItmsGrpCod").Value.ToString)
                    odbdsDetails.SetValue("U_GrpName", 0, objRs.Fields.Item("ItmsGrpNam").Value.ToString)
                    Matrix0.SetLineData(Matrix0.VisualRowCount)
                    objRs.MoveNext()
                End While
                Matrix0.AutoResizeColumns()

                If objaddon.HANA Then
                    objRs.DoQuery("select ROW_NUMBER() OVER () AS ""LineId"" ,""WhsCode"",""WhsName"" from OWHS")
                Else
                    objRs.DoQuery("select ROW_NUMBER() OVER () AS LineId ,WhsCode,WhsName from OWHS")
                End If
                odbdsDetails1 = objform.DataSources.DBDataSources.Item("@AT_USRCFG2")
                Matrix1.Clear()
                odbdsDetails1.Clear()
                Folder1.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                While Not objRs.EoF
                    Matrix1.AddRow()
                    Matrix1.GetLineData(Matrix1.VisualRowCount)
                    odbdsDetails1.SetValue("LineId", 0, objRs.Fields.Item("LineId").Value.ToString)
                    odbdsDetails1.SetValue("U_WhsCode", 0, objRs.Fields.Item("WhsCode").Value.ToString)
                    odbdsDetails1.SetValue("U_WhsName", 0, objRs.Fields.Item("WhsName").Value.ToString)
                    Matrix1.SetLineData(Matrix1.VisualRowCount)
                    objRs.MoveNext()
                End While
                Matrix1.AutoResizeColumns()
                objaddon.objapplication.StatusBar.SetText("Item Group & Warehouse Details loaded Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                objform.Freeze(False)
                objRs = Nothing
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                Matrix0.AutoResizeColumns()
                Matrix1.AutoResizeColumns()
                'EditText8.Item.Visible = False
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_ResizeAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                Matrix0.AutoResizeColumns()
                Matrix1.AutoResizeColumns()
            Catch ex As Exception

            End Try
        End Sub

        Private Sub EditText0_ChooseFromListBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles EditText0.ChooseFromListBefore
            Try

            Catch ex As Exception

            End Try

        End Sub

        Private Sub EditText0_KeyDownAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText0.KeyDownAfter
            Try
                If objform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
                If pVal.CharPressed = 9 Then
                    If EditText0.Value <> "" Then
                        If objaddon.HANA Then
                            strsql = objaddon.objglobalmethods.getSingleValue("select 1 ""Status"" from ""@AT_USRCFG"" where ""Code""='" & EditText0.Value & "' ")
                        Else
                            strsql = objaddon.objglobalmethods.getSingleValue("select 1 Status from [@AT_USRCFG] where Code='" & EditText0.Value & "' ")
                        End If
                        If strsql = "1" Then
                            objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            EditText0.Value = EditText0.Value
                            objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Folder0.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Exit Sub
                        End If
                    End If
                    LoadDetails()
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub EditText0_ChooseFromListAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText0.ChooseFromListAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                'If pVal.InnerEvent = True Then Exit Sub
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    objDBHeader = objform.DataSources.DBDataSources.Item("@AT_USRCFG")
                    Try
                        objDBHeader.SetValue("Code", 0, pCFL.SelectedObjects.Columns.Item("USER_CODE").Cells.Item(0).Value)
                    Catch ex As Exception
                    End Try
                    Try
                        objDBHeader.SetValue("Name", 0, pCFL.SelectedObjects.Columns.Item("U_NAME").Cells.Item(0).Value)
                    Catch ex As Exception
                    End Try
                    Try
                        Dim dept As String = objaddon.objglobalmethods.getSingleValue("Select (Select ""Name"" from OUDP where ""Code""=""Department"") ""Dept"" from OUSR where ""USER_CODE""='" & pCFL.SelectedObjects.Columns.Item("USER_CODE").Cells.Item(0).Value & "'")
                        objDBHeader.SetValue("U_Dept", 0, dept)
                    Catch ex As Exception
                    End Try
                    Try
                        objDBHeader.SetValue("U_UserId", 0, pCFL.SelectedObjects.Columns.Item("USERID").Cells.Item(0).Value)
                    Catch ex As Exception
                    End Try
                    Try
                        If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then objDBHeader.SetValue("U_CreateBy", 0, objaddon.objcompany.UserName)
                        If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then objDBHeader.SetValue("U_CreateOn", 0, objaddon.objglobalmethods.getSingleValue("Select TO_VARCHAR(Current_TIMESTAMP,'DD/MM/YYYY HH:MM:SS AM') ""Created On"" from Dummy"))
                        If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then objDBHeader.SetValue("U_UpdateBy", 0, objaddon.objcompany.UserName)
                        If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then objDBHeader.SetValue("U_UpdateOn", 0, objaddon.objglobalmethods.getSingleValue("Select TO_VARCHAR(Current_TIMESTAMP,'DD/MM/YYYY HH:MM:SS AM') ""Updated On"" from Dummy"))

                    Catch ex As Exception
                    End Try
                End If
            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try

        End Sub

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then objDBHeader.SetValue("U_CreateBy", 0, objaddon.objcompany.UserName)
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then objDBHeader.SetValue("U_CreateOn", 0, objaddon.objglobalmethods.getSingleValue("Select TO_VARCHAR(Current_TIMESTAMP,'DD/MM/YYYY HH:MM:SS AM') ""Created On"" from Dummy"))
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then objDBHeader.SetValue("U_UpdateBy", 0, objaddon.objcompany.UserName)
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then objDBHeader.SetValue("U_UpdateOn", 0, objaddon.objglobalmethods.getSingleValue("Select TO_VARCHAR(Current_TIMESTAMP,'DD/MM/YYYY HH:MM:SS AM') ""Updated On"" from Dummy"))
                objDBHeader.SetValue("U_UserId", 0, objaddon.objglobalmethods.getSingleValue("Select ""USERID"" from OUSR where ""USER_CODE""='" & objaddon.objcompany.UserName & "'"))
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_LoadAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                objform = objaddon.objapplication.Forms.GetForm("USRCONFIG", pVal.FormTypeCount)
            Catch ex As Exception
            End Try

        End Sub

        Private WithEvents LinkedButton0 As SAPbouiCOM.LinkedButton
        Private WithEvents EditText8 As SAPbouiCOM.EditText
    End Class
End Namespace
