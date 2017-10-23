<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../B81/B81COMM.ASP" -->

<%'Call loadInfTB19029B("I", "*", "NOCOOKIE","MB") %>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

On Error Resume Next
Call HideStatusWnd 

Const Y107_I1_INTERNAL_CD    = 0
Const Y107_I1_ITEM_CD        = 1
Const Y107_I1_OLD_ITEM_NM    = 2
Const Y107_I1_OLD_ITEM_NM2   = 3
Const Y107_I1_OLD_ITEM_SPEC  = 4
Const Y107_I1_OLD_ITEM_SPEC2 = 5
Const Y107_I1_NEW_ITEM_NM    = 6
Const Y107_I1_NEW_ITEM_NM2   = 7
Const Y107_I1_NEW_ITEM_SPEC  = 8
Const Y107_I1_NEW_ITEM_SPEC2 = 9
Const Y107_I1_REQ_ID         = 10
Const Y107_I1_REQ_DT         = 11
Const Y107_I1_REQ_REASON     = 12

'Dim lgIntFlgMode
Dim iStrSelectChar
Dim iCisChangeItemNmReq
Dim iStrReqNo
Dim iStrRefReqNo

Dim PY2G107

lgIntFlgMode = CInt(Request("txtFlgMode"))			'☜: 저장시 Create/Update 판별 
	    
If lgIntFlgMode = OPMD_CMODE Then
	iStrSelectChar = "CREATE"
	iStrReqNo = ""
ElseIf lgIntFlgMode = OPMD_UMODE Then
	iStrSelectChar = "UPDATE"
	iStrReqNo = Request("txtReqNo")
End If

Redim iCisChangeItemNmReq(12)

iCisChangeItemNmReq(Y107_I1_INTERNAL_CD)    = Request("htxtInternalCd")
iCisChangeItemNmReq(Y107_I1_ITEM_CD)        = Request("txtItemCd")
iCisChangeItemNmReq(Y107_I1_OLD_ITEM_NM)    = Request("txtItemNm")
iCisChangeItemNmReq(Y107_I1_OLD_ITEM_NM2)   = Request("txtItemNm2")
iCisChangeItemNmReq(Y107_I1_OLD_ITEM_SPEC)  = Request("txtSpec")
iCisChangeItemNmReq(Y107_I1_OLD_ITEM_SPEC2) = Request("txtSpec2")
iCisChangeItemNmReq(Y107_I1_NEW_ITEM_NM)    = Request("txtNewItemNm")
iCisChangeItemNmReq(Y107_I1_NEW_ITEM_NM2)   = Request("txtNewItemNm2")
iCisChangeItemNmReq(Y107_I1_NEW_ITEM_SPEC)  = Request("txtNewSpec")
iCisChangeItemNmReq(Y107_I1_NEW_ITEM_SPEC2) = Request("txtNewSpec2")
iCisChangeItemNmReq(Y107_I1_REQ_ID)         = Request("txtreq_user")
iCisChangeItemNmReq(Y107_I1_REQ_DT)         = uniconvDate(Request("txtReqDt"))
iCisChangeItemNmReq(Y107_I1_REQ_REASON)     = Request("htxtReqReason")
Set PY2G107 = Server.CreateObject("PY2G107.cCisChangeItemNmReq")

If CheckSYSTEMErrorY(Err,True,"") = True Then
	Response.End
End If


dim iErrorPosition,sErrTitle
iStrRefReqNo = PY2G107.Y_MAINT_CHANGE_ITEM_NM_REQ_SVR(gStrGlobalCollection, iStrSelectChar, iStrReqNo, iCisChangeItemNmReq,iErrorPosition )

if iErrorPosition<>"" then  
	
	select case iErrorPosition
		case "txtItemUnit"			: sErrTitle="재고단위"
		case "txtPurVendor"			: sErrTitle="공급처"
		case "txtPurGroup"			: sErrTitle="구매그룹"
		case "txtReqId"			: sErrTitle="의뢰자"
		case "txtNetWeightUnit"		: sErrTitle="Net중량"
		case "txtGrossWeightUnit"	: sErrTitle="Gross중량"
		case "txtHSCd"				: sErrTitle="HS Code"
		
	end select

end if
if iErrorPosition = "txtReqId" then  iErrorPosition="txtReq_user"

 
If CheckSYSTEMErrorY(Err,True,sErrTitle) = True Then
	if iErrorPosition<>"" then
		goFocus(iErrorPosition)
	end if
	Response.End
else 

End If


		              
Set PY2G107 = Nothing

Response.Write "<Script Language=vbscript>" & vbCr
Response.Write "With parent"				& vbCr
	
If iStrRefReqNo <> "" then																		
   Response.Write ".frm1.txtReqNo.value	= """ & ConvSPChars(iStrRefReqNo) & """" & vbcr
End If


	
Response.Write ".DbSaveOk"                  & vbCr
Response.Write "End With"                   & vbCr
Response.Write "</Script>"                  & vbCr
Response.End




	
%>