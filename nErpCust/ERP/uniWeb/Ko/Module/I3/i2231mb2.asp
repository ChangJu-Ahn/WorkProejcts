<%@  LANGUAGE = VBSCript%>
<% Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : i2231mb2.asp
'*  4. Program Name         : 재고이월확정 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002-01-03
'*  8. Modified date(Last)  : 2005-12-29
'*  9. Modifier (First)     : Han, SungGyu
'* 10. Modifier (Last)      : Lee SeungWook
'* 11. Comment              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%																			
Call LoadBasisGlobalInf()
Call HideStatusWnd															
	
Dim strRetMsg										
Dim IntRetCd

Dim txtPlantCd
Dim txtInsrtUserId

	On Error Resume Next														
 	Err.Clear																

	Call SubOpenDB(lgObjConn)
	Call SubCreateCommandObject(lgObjComm)
	
	txtPlantCd = UCase(Trim(Request("txtPlantCd")))
	txtInsrtUserId = UCase(Trim(Request("txtInsrtUserId")))
	
	With lgObjComm
		.CommandText = "usp_i_CREATE_INVENTORY_CARRYOVER"
		.CommandType = adCmdStoredProc
		.CommandTimeout = 1800	
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE", adInteger,	adParamReturnValue)
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_cd", adVarChar, adParamInput, 4, txtPlantCd)
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@insrt_user_id", adVarChar, adParamInput, 13,	txtInsrtUserId)	   
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_nm", adVarChar, adParamOutput, 40)
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@from_inv_cls_dt", adVarChar, adParamOutput, 10)
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd", adVarChar, adParamOutput, 8)
	
		lgObjComm.Execute ,, adExecuteNoRecords
	End With
	
	If Err.number = 0 Then
		intRetCd = lgObjComm.Parameters("RETURN_VALUE").Value
		
		If intRetCd <> 0 Then
			strRetMsg = lgObjComm.Parameters("@msg_cd").Value
			If strRetMsg <> "" Then
				Call DisplayMsgBox(strRetMsg, vbInformation, "", "", I_MKSCRIPT)
				
			End If	
		End If
	Else
		Call DisplayMsgBox(Err.Description, vbInformation, "", "", I_MKSCRIPT)	
	End If
	
	Response.Write "<Script Language=vbscript>	"	& vbcr
	Response.Write "With parent.frm1			"	& vbcr

	Response.Write "	.txtinvclsdt.text =	"""		& UniMonthClientFormat(lgObjComm.Parameters("@from_inv_cls_dt").Value) & """ " & vbcr
	Response.Write "	Call parent.Btnabled()	"	& vbcr
	If strRetMsg = "" Then
		Response.Write "	parent.DbSaveOk2	"		& vbcr
	Else
		Response.Write "parent.frm1.txtPlantCd.focus	"	& vbcr	
	End If
	Response.Write "End With					"	& vbcr
	Response.Write "</Script>					"	& vbcr	 

    Call SubCloseCommandObject(lgObjComm)
	Call SubCloseDB(lgObjConn) 
		
%>

