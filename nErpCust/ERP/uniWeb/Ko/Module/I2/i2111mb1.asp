<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : Create physical inventory document 
'*  3. Program ID           : I2111mb1.asp
'*  4. Program Name         : 실사선별batch등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             PI2G010.cIMaintPhyInvBch
'*  7. Modified date(First) : 2000/04/06
'*  8. Modified date(Last)  : 2002/06/28
'*  9. Modifier (First)     : Mr  Kim Nam Hoon
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment              : VB Conversion
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
Call LoadBasisGlobalInf()
							
On Error Resume Next
Err.Clear
Call HideStatusWnd

Dim pPI2G010														
Dim strMode															
Dim iCommandSent
Dim I1_cnt_period
Dim I2_i_physical_inventory_detail_abc_flag 

Dim I3_b_item_from
	Const I201_I3_item_cd = 0
    Const I201_I3_item_group_cd = 1
ReDim I3_b_item_from(I201_I3_item_group_cd)

Dim I4_b_item_to
    Const I201_I4_item_cd = 0
    Const I201_I4_item_group_cd = 1
ReDim I4_b_item_to(I201_I4_item_group_cd)

Dim I5_b_plant_plant_cd
Dim I6_b_storage_location_sl_cd
Dim I7_i_onhand_stock_tracking_no

Dim I8_i_physical_inventory_header
    Const I201_I8_phy_inv_no = 0
    Const I201_I8_real_insp_dt = 1
ReDim I8_i_physical_inventory_header(I201_I8_real_insp_dt)

Dim I9_b_minor_minor_cd

Dim E1_i_physical_inventory_header_phy_inv_no

'============================================================================================================
' Name : SubBizBatchMulti
' Desc : Batch Multi Data
'============================================================================================================
Sub SubBizBatchMulti()

    Dim IntRetCD
    Dim strMsg_cd, strMsg_text
    Dim strPhyInvNo,strInspODt, strCurrDt, strInspDt

	strPhyInvNo	=	E1_i_physical_inventory_header_phy_inv_no
	strInspODt	=	UNIConvDate(Request("txtInspDt"))			
	
	strInspDt	=	UNIConvDateAToB(strInspODt, gAPDateFormat, gServerDateFormat)	
	strCurrDt   =   GetSvrDate		
	
	If strInspDt < strCurrDt Then
	 Call SubOpenDB(lgObjConn)                    
	 Call SubCreateCommandObject(lgObjComm)
	
		With lgObjComm
			.CommandText = "usp_i_update_phy_inv_qty"
			.CommandType = adCmdStoredProc

			lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@phy_inv_no"     ,advarxchar,adParamInput,Len(Trim(strPhyInvNo)), strPhyInvNo)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@real_count_dt",adDate,adParamInput,Len(Trim(strInspODt)), strInspODt)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@updt_user_id"     ,advarxchar,adParamInput,Len(Trim(gUsrID)), gUsrID)
		    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"   ,advarxchar ,adParamOutput,6)

			lgObjComm.Execute ,, adExecuteNoRecords

		End With

		If  Err.number = 0 Then
			IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
			
			If  IntRetCD <> 1 then
            
				strMsg_cd   = lgObjComm.Parameters("@msg_cd").Value
				strSpId     = FilterVar(lgObjComm.Parameters("@updt_user_id").Value, "''", "S")
            
				If strMsg_cd <> MSG_OK_STR Then
					Call DisplayMsgBox(strMsg_cd, vbInformation, "", "", I_MKSCRIPT)
				End If
				IntRetCD = -1
				Exit Sub
			Else
				IntRetCD = 1
			End if
		Else           
			Call SvrMsgBox(Err.Description, vbinformation, i_mkscript)
			Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
			IntRetCD = -1
		End if

     Call SubCloseCommandObject(lgObjComm)
     Call SubCloseDB(lgObjConn)       

	Else 
		Exit Sub
	End if

End Sub		

strMode = Request("txtMode")										


Select Case strMode

Case CStr(UID_M0002)													

    iCommandSent					="CREATE"
    I5_b_plant_plant_cd             = UCase(Request("txtPlantCd"))
    I6_b_storage_location_sl_cd     = UCase(Request("txtSLCd"))
    I3_b_item_from(I201_I3_item_cd) = Request("txtItemOriginCd")
    I4_b_item_to(I201_I4_item_cd)   = Request("txtItemDestCd")
    I3_b_item_from(I201_I3_item_group_cd) = Request("txtItemGroupOriginCd")
    I4_b_item_to(I201_I4_item_group_cd)   = Request("txtItemGroupDestCd")
    I7_i_onhand_stock_tracking_no   = Request("txtTrackingNo") 
    I1_cnt_period					= Request("cboCntPerd")   
    I2_i_physical_inventory_detail_abc_flag				= Request("cboABCFlag")
    I8_i_physical_inventory_header(I201_I8_phy_inv_no)  = Trim(Request("txtPhyinvNo"))
    I8_i_physical_inventory_header(I201_I8_real_insp_dt)= UNIConvDate(Request("txtInspDt"))     
	I9_b_minor_minor_cd				= Request("cboInvMgr") 


     Set pPI2G010 = Server.CreateObject("PI2G010.cIMaintPhyInvBch")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If

    '-----------------------
    'Com action area
    '-----------------------                                                   
	E1_i_physical_inventory_header_phy_inv_no = pPI2G010.I_MAINT_PHY_INV_BATCH(gStrGlobalCollection, iCommandSent, _
																				I1_cnt_period, _
																				I2_i_physical_inventory_detail_abc_flag, _
																				I3_b_item_from, _
																				I4_b_item_to, _
																				I5_b_plant_plant_cd, _
																				I6_b_storage_location_sl_cd, _
																				I7_i_onhand_stock_tracking_no, _
																				I8_i_physical_inventory_header, _
																				I9_b_minor_minor_cd)

	If CheckSYSTEMError(Err, True) = True Then
		Set pPI2G010 = Nothing															
		Response.End
	End If

    Set pPI2G010 = Nothing	
    													
	 Call SubBizBatchMulti()
End select
%>
<SCRIPT Language="VBScript">
   With parent
   	.frm1.txtPhyInvNo.value = "<%=ConvSPChars(E1_i_physical_inventory_header_phy_inv_no)%>"
	.DbSaveOk
   End With
</SCRIPT>
