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

Const Y101_I1_ITEM_ACCT       = 0
Const Y101_I1_ITEM_KIND       = 1
Const Y101_I1_ITEM_LVL1       = 2
Const Y101_I1_ITEM_LVL2       = 3
Const Y101_I1_ITEM_LVL3       = 4
Const Y101_I1_ITEM_DERIVE     = 5
Const Y101_I1_ITEM_VER        = 6
Const Y101_I1_ITEM_NM         = 7
Const Y101_I1_ITEM_NM2        = 8
Const Y101_I1_ITEM_SPEC       = 9
Const Y101_I1_ITEM_SPEC2      = 10
Const Y101_I1_ITEM_UNIT       = 11
Const Y101_I1_PUR_TYPE        = 12
Const Y101_I1_BASIC_CODE      = 13
Const Y101_I1_PUR_GROUP       = 14
Const Y101_I1_PUR_VENDOR      = 15
Const Y101_I1_UNIFY_PUR_FLAG  = 16
Const Y101_I1_UNIT_WEIGHT     = 17
Const Y101_I1_UNIT_OF_WEIGHT  = 18
Const Y101_I1_GROSS_WEIGHT    = 19
Const Y101_I1_GROSS_UNIT      = 20
Const Y101_I1_CBM             = 21
Const Y101_I1_CBM_DESCRIPTION = 22
Const Y101_I1_HS_CODE         = 23
Const Y101_I1_VALID_FROM_DT   = 24
Const Y101_I1_VALID_TO_DT     = 25
Const Y101_I1_DOC_NO          = 26
Const Y101_I1_REQ_ID          = 27
Const Y101_I1_REQ_DT          = 28
Const Y101_I1_REQ_REASON      = 29
Const Y101_I1_REQ_END_DT      = 30
Const Y101_I1_REMARK          = 31
Const Y101_I1_STATUS          = 32

'Dim lgIntFlgMode
Dim iStrSelectChar
Dim iCisNewItemReq
Dim iStrReqNo
Dim iStrRefReqNo

Dim PY2G101

lgIntFlgMode = CInt(Request("txtFlgMode"))			'☜: 저장시 Create/Update 판별 
call chk()


	    
If lgIntFlgMode = OPMD_CMODE Then
	iStrSelectChar = "CREATE"
	iStrReqNo = ""
ElseIf lgIntFlgMode = OPMD_UMODE Then
	iStrSelectChar = "UPDATE"
	iStrReqNo = Request("txtReqNo")
End If
 
  Redim iCisNewItemReq(33)

	
        
iCisNewItemReq(Y101_I1_ITEM_ACCT)      = Request("cboItemAcct")
iCisNewItemReq(Y101_I1_ITEM_KIND)      = Request("txtItemKind")
iCisNewItemReq(Y101_I1_ITEM_LVL1)      = Request("txtItemLvl1")
iCisNewItemReq(Y101_I1_ITEM_LVL2)      = Request("txtItemLvl2")
iCisNewItemReq(Y101_I1_ITEM_LVL3)      = Request("txtItemLvl3")
iCisNewItemReq(Y101_I1_ITEM_DERIVE)    = Request("hrdoDerive")
iCisNewItemReq(Y101_I1_ITEM_VER)       = Request("cboItemVer")
iCisNewItemReq(Y101_I1_ITEM_NM)        = Request("txtItemNm")
iCisNewItemReq(Y101_I1_ITEM_NM2)       = Request("txtItemNm2")
iCisNewItemReq(Y101_I1_ITEM_SPEC)      = Request("txtSpec")
iCisNewItemReq(Y101_I1_ITEM_SPEC2)     = Request("txtSpec2")
iCisNewItemReq(Y101_I1_ITEM_UNIT)      = Request("txtItemUnit")
iCisNewItemReq(Y101_I1_PUR_TYPE)       = Request("cboPurType")
iCisNewItemReq(Y101_I1_BASIC_CODE)     = Request("txtBasicItem")
iCisNewItemReq(Y101_I1_PUR_GROUP)      = Request("txtPurGroup")
iCisNewItemReq(Y101_I1_PUR_VENDOR)     = Request("txtPurVendor")
iCisNewItemReq(Y101_I1_UNIFY_PUR_FLAG) = Request("hrdoUnifyPurFlg")
iCisNewItemReq(Y101_I1_UNIT_WEIGHT)    = UNIConvNum(Request("txtNetWeight"),0)
iCisNewItemReq(Y101_I1_UNIT_OF_WEIGHT) = Request("txtNetWeightUnit")
iCisNewItemReq(Y101_I1_GROSS_WEIGHT)   = UNIConvNum(Request("txtGrossWeight"),0)
iCisNewItemReq(Y101_I1_GROSS_UNIT)     = Request("txtGrossWeightUnit")
iCisNewItemReq(Y101_I1_CBM)            = UNIConvNum(Request("txtCBM"),0) 
iCisNewItemReq(Y101_I1_CBM_DESCRIPTION)= Request("txtCBMInfo")
iCisNewItemReq(Y101_I1_HS_CODE)        = Request("txtHSCd")
iCisNewItemReq(Y101_I1_VALID_FROM_DT)  = uniconvDate(Request("txtValidFromDt"))
iCisNewItemReq(Y101_I1_VALID_TO_DT)    = uniconvDate(Request("txtValidToDt"))
iCisNewItemReq(Y101_I1_DOC_NO)         = Request("txtDocNo")
iCisNewItemReq(Y101_I1_REQ_ID)         = Request("txtreq_user")
iCisNewItemReq(Y101_I1_REQ_DT)         = uniconvDate(Request("txtEndReqDt"))
iCisNewItemReq(Y101_I1_REQ_REASON)     = Request("htxtReqReason")
iCisNewItemReq(Y101_I1_REQ_END_DT)     = uniconvDate(Request("txtReqDt"))
iCisNewItemReq(Y101_I1_REMARK)         = Request("htxtRemark")
iCisNewItemReq(Y101_I1_STATUS)         = Request("htxtStatus")

'call chkItemCd()
Set PY2G101 = Server.CreateObject("PY2G101.cCisNewItemReq")

If CheckSYSTEMErrorY(Err,True,"") = True Then
	Response.End
End If

dim iErrorPosition,sErrTitle
iStrRefReqNo = PY2G101.Y_MAINT_NEW_ITEM_REQ_SVR(gStrGlobalCollection, iStrSelectChar, iStrReqNo, iCisNewItemReq ,iErrorPosition)
if iErrorPosition<>"" then  
	
	select case iErrorPosition
		case "txtItemUnit"			: sErrTitle="재고단위"
		case "txtPurVendor"			: sErrTitle="공급처"
		case "txtPurGroup"			: sErrTitle="구매그룹"
		case "txtReqId"				: sErrTitle="의뢰자"
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

On Error Goto 0	    
		              
Set PY2G101 = Nothing




Response.Write "<Script Language=vbscript>" & vbCr
Response.Write "With parent"				& vbCr
	
If iStrRefReqNo <> "" then																		
   Response.Write ".frm1.txtarReqNo.value	= """ & ConvSPChars(iStrRefReqNo) & """" & vbcr
End If
	
Response.Write ".DbSaveOk"                  & vbCr
Response.Write "End With"                   & vbCr
Response.Write "</Script>"                  & vbCr
Response.End



sub chk()

	if Request("txtItemKindNm")="" then
		Call DisplayMsgBox("970000", vbInformation, "품목구분", "", I_MKSCRIPT)	
		goFocus("txtItemKind")
		Response.End 
	end if
	if Request("txtItemLvl1nm")="" then
		Call DisplayMsgBox("970000", vbInformation, "대분류", "", I_MKSCRIPT)	
		goFocus("txtItemLvl1")
		Response.End 
	end if
	if Request("txtItemLvl2nm")="" then
		Call DisplayMsgBox("970000", vbInformation, "중분류", "", I_MKSCRIPT)	
		goFocus("txtItemLvl2")
		Response.End 
	end if
	if Request("txtItemLvl3nm")="" then
		Call DisplayMsgBox("970000", vbInformation, "소분류", "", I_MKSCRIPT)	
		goFocus("txtItemLvl3")
		Response.End 
	end if
	
end sub



%>