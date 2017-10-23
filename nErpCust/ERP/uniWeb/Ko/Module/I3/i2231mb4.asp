<%@  LANGUAGE = VBSCript%>
<% Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : Cancel Inventory Carry-Over
'*  3. Program ID           : I2231mb4.asp
'*  4. Program Name         : 재고이월 취소작업 
'*  5. Program Desc         : 이월된 재고를 전월로 복귀시킨다.
'*  6. Comproxy List        : usp_I_CANCEL_INVENTORY_CARRYOVER.sql                           
'                             
'*  7. Modified date(First) : 2002/07/11
'*  8. Modified date(Last)  : 2005/12/29
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2002/07/11 : ..........
'**********************************************************************************************-->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
Call LoadBasisGlobalInf()
Call HideStatusWnd 

On Error Resume Next
Err.Clear


Dim strRetMsg										
Dim intRetCd

Dim txtPlantCd
Dim txtInsrtUserId

	Call SubOpenDB(lgObjConn)
 	Call SubCreateCommandObject(lgObjComm)

	txtPlantCd = UCase(Trim(Request("txtPlantCd")))
	txtInsrtUserId = UCase(Trim(Request("txtInsrtUserId")))
	
	With lgObjComm
		.CommandText = "usp_I_CANCEL_INVENTORY_CARRYOVER"
		.CommandType = adCmdStoredProc
		.CommandTimeout = 1800
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adinteger,adParamReturnValue)
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_cd",adVarChar,adParamInput,4,txtPlantCd)
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@insrt_user_id", adVarChar, adParamInput, 13,	txtInsrtUserId)
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_nm", adVarChar, adParamOutput, 40)
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ex_inv_cls_dt", adVarChar, adParamOutput, 10)
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd", adVarChar, adParamOutput, 8)
		
		lgObjComm.Execute ,, adExecuteNoRecords
		
	End With
	
	If Err.number = 0 Then
		intRetCd = lgObjComm.Parameters("RETURN_VALUE").Value
		If intRetCd <> 0 then
			strRetMsg = lgObjComm.Parameters("@msg_cd").Value
			If strRetMsg <> "" Then
				Call DisplayMsgBox(strRetMsg, vbInformation, "", "", I_MKSCRIPT)
				
				Response.Write "<Script Language=vbscript>	"		& vbcr
				Response.Write "parent.frm1.txtPlantCd.focus	"	& vbcr
				Response.Write "	Call parent.Btnabled()	"		& vbcr
				Response.Write "</Script>					"		& vbcr
				Response.End
			End If
		End If
	Else
		Call DisplayMsgBox(Err.Description, vbInformation, "", "", I_MKSCRIPT)	
	End If
	
	Response.Write "<Script Language=vbscript>	"	& vbcr
	Response.Write "With parent.frm1			"	& vbcr

	Response.Write "	.txtinvclsdt.text =	"""		& UniMonthClientFormat(lgObjComm.Parameters("@from_inv_cls_dt").Value) & """ " & vbcr
	Response.Write "	Call parent.Btnabled()	"	& vbcr
	Response.Write "	parent.DbSaveOk3	"		& vbcr
	Response.Write "End With					"	& vbcr
	Response.Write "</Script>					"	& vbcr	 

    Call SubCloseCommandObject(lgObjComm)
	Call SubCloseDB(lgObjConn)

%>
