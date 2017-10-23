<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%Call loadInfTB19029B("I", "*", "NOCOOKIE","MB") %>
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

Dim PY2G107													'☆ : 조회용 ComProxy Dll 사용 변수 

Const Y107_E1_REQ_NO         = 0
Const Y107_E1_STATUS         = 1
Const Y107_E1_ITEM_CD        = 2
Const Y107_E1_OLD_ITEM_NM    = 3
Const Y107_E1_OLD_ITEM_NM2   = 4
Const Y107_E1_OLD_ITEM_SPEC  = 5
Const Y107_E1_OLD_ITEM_SPEC2 = 6
Const Y107_E1_ITEM_ACCT      = 7
Const Y107_E1_ITEM_ACCT_NM   = 8
Const Y107_E1_ITEM_KIND      = 9
Const Y107_E1_ITEM_KIND_NM   = 10
Const Y107_E1_ITEM_LVL1      = 11
Const Y107_E1_ITEM_LVL1_NM   = 12
Const Y107_E1_ITEM_LVL2      = 13
Const Y107_E1_ITEM_LVL2_NM   = 14
Const Y107_E1_ITEM_LVL3      = 15
Const Y107_E1_ITEM_LVL3_NM   = 16
Const Y107_E1_ITEM_SEQNO     = 17
Const Y107_E1_ITEM_DERIVE    = 18
Const Y107_E1_ITEM_LVL_D     = 19
Const Y107_E1_ITEM_VER       = 20
Const Y107_E1_BASIC_CODE     = 21
Const Y107_E1_BASIC_CODE_NM  = 22
Const Y107_E1_NEW_ITEM_NM    = 23
Const Y107_E1_NEW_ITEM_NM2   = 24
Const Y107_E1_NEW_ITEM_SPEC  = 25
Const Y107_E1_NEW_ITEM_SPEC2 = 26
Const Y107_E1_REQ_ID         = 27
Const Y107_E1_REQ_ID_NM      = 28
Const Y107_E1_REQ_DT         = 29
Const Y107_E1_REQ_REASON     = 30
Const Y107_E1_REQ_END_DT     = 31
Const Y107_E1_END_DT         = 32
Const Y107_E1_TRANS_DT       = 33
Const Y107_E1_R_GRADE        = 34
Const Y107_E1_T_GRADE        = 35
Const Y107_E1_P_GRADE        = 36
Const Y107_E1_Q_GRADE        = 37
Const Y107_E1_STATUS_NM      = 38
Const Y107_E1_INTERNAL_CD    = 39

Dim iExportChangItemNmReq
Dim iStrPreNextError

Set PY2G107 = Server.CreateObject("PY2G107.cCisChangeItemNmReqQuery")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PY2G107.Y_CIS_CHANGE_ITEM_NM_REQ_SVR(gStrGlobalCollection, Request("PrevNextFlg"), Request("txtReqNo"), iExportChangItemNmReq, iStrPreNextError)

If CheckSYSTEMError(Err,True) = True Then
	Set PY2G107 = Nothing
	Response.End
End If

If iStrPreNextError = "900011" Or iStrPreNextError = "900012" Then
	Call DisplayMsgBox(iStrPreNextError, vbOKOnly, "", "", I_MKSCRIPT)
End If
	
Set PY2G107 = Nothing

%>
<Script Language=vbscript>
With Parent.frm1
    .txtarReqNo.Value		   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_REQ_NO)))%>"
			
	.txtReqNo.Value		       = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_REQ_NO)))%>"
	.txtStatus.Value		   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_STATUS_NM)))%>"
			
	.txtItemCd.Value	       = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_ITEM_CD)))%>"			
	.txtItemNm.Value	       = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_OLD_ITEM_NM)))%>"
	.txtItemNm2.Value		   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_OLD_ITEM_NM2)))%>"
	.txtSpec.Value	           = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_OLD_ITEM_SPEC)))%>"
	.txtSpec2.Value	           = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_OLD_ITEM_SPEC2)))%>"
			
	.txtItemAcct.Value		   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_ITEM_ACCT)))%>"
	.txtItemAcctNm.Value	   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_ITEM_ACCT_NM)))%>"
	.txtItemKind.Value		   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_ITEM_KIND)))%>"
	.txtItemKindNm.Value	   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_ITEM_KIND_NM)))%>"
	.txtItemLvl1.Value		   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_ITEM_LVL1)))%>"	
	.txtItemLvl1NM.Value	   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_ITEM_LVL1_NM)))%>"
	.txtItemLvl2.Value		   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_ITEM_LVL2)))%>"
	.txtItemLvl2Nm.Value	   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_ITEM_LVL2_NM)))%>"			
	.txtItemLvl3.Value		   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_ITEM_LVL3)))%>"
	.txtItemLvl3Nm.Value	   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_ITEM_LVL3_NM)))%>"
			
	If Trim("<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_ITEM_DERIVE)))%>") = "Y" Then
		.rdoDerive1.CHECKED	= TRUE
		.rdoDerive2.CHECKED	= FALSE
		.hrdoDerive.value = "Y"
	Else
	    .rdoDerive1.CHECKED	= FALSE
		.rdoDerive2.CHECKED	= TRUE
		.hrdoDerive.value     = "N"
	End If
			
	.txtSerialNo.Value		   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_ITEM_SEQNO)))%>"
	.txtBasicItem.Value	       = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_BASIC_CODE)))%>"
	.txtBasicItemNm.Value	   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_BASIC_CODE_NM)))%>"
	.tXtItemVer.Value	       = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_ITEM_VER)))%>"
			
	.txtNewItemNm.Value	       = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_NEW_ITEM_NM)))%>"
	.txtNewItemNm2.Value	   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_NEW_ITEM_NM2)))%>"
	.txtNewSpec.Value	       = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_NEW_ITEM_SPEC)))%>"
	.txtNewSpec2.Value	       = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_NEW_ITEM_SPEC2)))%>"
			
	.txtreq_user.Value	           = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_REQ_ID)))%>"
	.txtreq_user_Nm.Value		   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_REQ_ID_NM)))%>"
	.txtReqDt.Text             = "<%=UniConvDateDbToCompany(Trim(iExportChangItemNmReq(Y107_E1_REQ_DT)),"")%>"
			
	If "<%=UniConvDateDbToCompany(Trim(iExportChangItemNmReq(Y107_E1_END_DT)),"")%>" = "1900-01-01" Then
		.txtEndDt.Text = ""
	Else
		.txtEndDt.Text = "<%=UniConvDateDbToCompany(Trim(iExportChangItemNmReq(Y107_E1_END_DT)),"")%>"
	End If
	If "<%=UniConvDateDbToCompany(Trim(iExportChangItemNmReq(Y107_E1_TRANS_DT)),"")%>" = "1900-01-01" Then
		.txtTransDt.Text = ""
	Else
		.txtTransDt.Text = "<%=UniConvDateDbToCompany(Trim(iExportChangItemNmReq(Y107_E1_TRANS_DT)),"")%>"
	End If
			
	.txtRgrade.Value	       = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_R_GRADE)))%>"
	.txtTgrade.Value	       = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_T_GRADE)))%>"
	.txtPgrade.Value	       = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_P_GRADE)))%>"
	.txtQgrade.Value	       = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_Q_GRADE)))%>" 
			
	.htxtReqReason.Value	   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_REQ_REASON)))%>"
			
	.htxtStatus.Value		   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_STATUS)))%>"
	.htxtInternalCd.Value	   = "<%=ConvSPChars(Trim(iExportChangItemNmReq(Y107_E1_INTERNAL_CD)))%>"	
			
End with
	
Call Parent.DbQueryOk()

</Script>