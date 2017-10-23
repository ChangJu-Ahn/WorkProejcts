<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
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

Const Y104_I1_INTERNAL_CD     = 0
Const Y104_I1_ITEM_CD         = 1
Const Y104_I1_ITEM_NM         = 2
Const Y104_I1_ITEM_NM2        = 3
Const Y104_I1_ITEM_SPEC       = 4
Const Y104_I1_ITEM_SPEC2      = 5
Const Y104_I1_VERSION_CHG     = 6
Const Y104_I1_END_ITEM_FLAG   = 7
Const Y104_I1_ITEM_UNIT       = 8
Const Y104_I1_PUR_TYPE        = 9
Const Y104_I1_PUR_GROUP       = 10
Const Y104_I1_PUR_VENDOR      = 11
Const Y104_I1_UNIFY_PUR_FLAG  = 12
Const Y104_I1_UNIT_WEIGHT     = 13
Const Y104_I1_UNIT_OF_WEIGHT  = 14
Const Y104_I1_GROSS_WEIGHT    = 15
Const Y104_I1_GROSS_UNIT      = 16
Const Y104_I1_CBM             = 17
Const Y104_I1_CBM_DESCRIPTION = 18
Const Y104_I1_HS_CODE         = 19
Const Y104_I1_VALID_FROM_DT   = 20
Const Y104_I1_VALID_TO_DT     = 21
Const Y104_I1_DOC_NO          = 22
Const Y104_I1_REQ_ID          = 23
Const Y104_I1_REQ_DT          = 24
Const Y104_I1_REQ_REASON      = 25
Const Y104_I1_REMARK          = 26


'Dim lgIntFlgMode
Dim iStrSelectChar
Dim iCisChangeItemReq
Dim iStrReqNo
Dim iStrRefReqNo

Dim PY2G104

lgIntFlgMode = CInt(Request("txtFlgMode"))			'☜: 저장시 Create/Update 판별 
	    
If lgIntFlgMode = OPMD_CMODE Then
	iStrSelectChar = "CREATE"
	iStrReqNo = ""
ElseIf lgIntFlgMode = OPMD_UMODE Then
	iStrSelectChar = "UPDATE"
	iStrReqNo = Request("txtReqNo")
End If

Redim iCisChangeItemReq(27)

iCisChangeItemReq(Y104_I1_INTERNAL_CD)    = Request("htxtInternalCd")
iCisChangeItemReq(Y104_I1_ITEM_CD)        = Request("txtItemCd")
iCisChangeItemReq(Y104_I1_ITEM_NM)        = Request("txtItemNm")
iCisChangeItemReq(Y104_I1_ITEM_NM2)       = Request("txtItemNm2")
iCisChangeItemReq(Y104_I1_ITEM_SPEC)      = Request("txtSpec")
iCisChangeItemReq(Y104_I1_ITEM_SPEC2)     = Request("txtSpec2")
iCisChangeItemReq(Y104_I1_VERSION_CHG)    = Request("hrdoChgVer")
iCisChangeItemReq(Y104_I1_END_ITEM_FLAG)  = Request("hrdoEndItem")
iCisChangeItemReq(Y104_I1_ITEM_UNIT)      = Request("txtItemUnit")
iCisChangeItemReq(Y104_I1_PUR_TYPE)       = Request("cboPurType")
iCisChangeItemReq(Y104_I1_PUR_GROUP)      = Request("txtPurGroup")
iCisChangeItemReq(Y104_I1_PUR_VENDOR)     = Request("txtPurVendor")
iCisChangeItemReq(Y104_I1_UNIFY_PUR_FLAG) = Request("hrdoUnifyPurFlg")
iCisChangeItemReq(Y104_I1_UNIT_WEIGHT)    = UNIConvNum(Request("txtNetWeight"),0)
iCisChangeItemReq(Y104_I1_UNIT_OF_WEIGHT) = Request("txtNetWeightUnit")
iCisChangeItemReq(Y104_I1_GROSS_WEIGHT)   = UNIConvNum(Request("txtGrossWeight"),0)
iCisChangeItemReq(Y104_I1_GROSS_UNIT)     = Request("txtGrossWeightUnit")
iCisChangeItemReq(Y104_I1_CBM)            = UNIConvNum(Request("txtCBM"),0) 
iCisChangeItemReq(Y104_I1_CBM_DESCRIPTION)= Request("txtCBMInfo")
iCisChangeItemReq(Y104_I1_HS_CODE)        = Request("txtHSCd")
iCisChangeItemReq(Y104_I1_VALID_FROM_DT)  = uniconvDate(Request("txtValidFromDt"))
iCisChangeItemReq(Y104_I1_VALID_TO_DT)    = uniconvDate(Request("txtValidToDt"))
iCisChangeItemReq(Y104_I1_DOC_NO)         = Request("txtDocNo")
iCisChangeItemReq(Y104_I1_REQ_ID)         = Request("txtreq_user")
iCisChangeItemReq(Y104_I1_REQ_DT)         = uniconvDate(Request("txtReqDt"))
iCisChangeItemReq(Y104_I1_REQ_REASON)     = Request("htxtReqReason")
iCisChangeItemReq(Y104_I1_REMARK)         = Request("htxtRemark")

Set PY2G104 = Server.CreateObject("PY2G104.cCisChangeItemReq")

    
If CheckSYSTEMErrorY(Err,True,"") = True Then
	Set PY2G104 = Nothing
	Response.End
End If		 


dim iErrorPosition,sErrTitle
iStrRefReqNo = PY2G104.Y_MAINT_CHANGE_ITEM_REQ_SVR(gStrGlobalCollection, iStrSelectChar, iStrReqNo, iCisChangeItemReq ,iErrorPosition)
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
		              
Set PY2G104 = Nothing

Response.Write "<Script Language=vbscript>" & vbCr
Response.Write "With parent"				& vbCr
	
If iStrRefReqNo <> "" then																		
   Response.Write ".frm1.txtReqNo.value	= """ & ConvSPChars(iStrRefReqNo) & """" & vbcr
End If
	
Response.Write ".DbSaveOk"                  & vbCr
Response.Write "End With"                   & vbCr
Response.Write "</Script>"                  & vbCr
Response.End







'----------------------------------------------------------------------------------------------------------
' chkItemCd
'  Code Value check.
'----------------------------------------------------------------------------------------------------------
sub chkItemCd()
  
    Call SubOpenDB(lgObjConn)
    
    call GetNameChk("UNIT_NM","B_UNIT_OF_MEASURE","DIMENSION <> 'TM' and Unit="&filtervar(Request("txtItemUnit"),"''","S") &"",Request("txtItemUnit") ,"txtItemUnit","재고단위","Y") '
    
    call GetNameChk("BP_NM","B_BIZ_PARTNER","BP_TYPE In ('S','CS') And usage_flag='Y' AND IN_OUT_FLAG = 'O' And BP_CD="&filterVar(request("txtPurVendor"),"''","S")&"",request("txtPurVendor") ,"txtPurVendor","공급처","Y") '
    call GetNameChk("PUR_GRP_NM","B_Pur_Grp","USAGE_FLG='Y' And PUR_GRP="&filterVar(Request("txtPurGroup"),"''","S") ,Request("txtPurGroup") ,"txtPurGroup","구매그룹","Y") 
    call GetNameChk("MINOR_NM","B_MINOR"," MAJOR_CD = 'Y1006' And MINOR_CD="&filtervar(Request("txtreq_user"),"''","S") &"",request("txtreq_user"),"txtreq_user","의뢰자","Y") '
    call GetNameChk("UNIT_NM","B_UNIT_OF_MEASURE","DIMENSION <> 'TM' And UNIT="&filterVar(Request("txtNetWeightUnit"),"''","S") ,Request("txtNetWeightUnit") ,"txtNetWeightUnit","Net중량","Y") '
    call GetNameChk("UNIT_NM","B_UNIT_OF_MEASURE","DIMENSION <> 'TM' And UNIT="&filterVar(Request("txtGrossWeightUnit"),"''","S") ,Request("txtGrossWeightUnit") ,"txtGrossWeightUnit","Gross중량","Y") '
    call GetNameChk("HS_NM","B_HS_CODE","HS_CD="&filterVar(Request("txtHSCd"),"''","S") ,Request("txtHSCd") ,"txtHSCd","HS Code","Y") '
	
	Call SubCloseDB(lgObjConn) 
   
	
End Sub	




	
%>