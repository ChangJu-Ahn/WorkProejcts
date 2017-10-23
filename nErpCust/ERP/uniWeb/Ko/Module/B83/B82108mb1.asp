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

Dim PY2G108													'☆ : 조회용 ComProxy Dll 사용 변수 

Const Y108_E1_REQ_NO         = 0
Const Y108_E1_STATUS         = 1
Const Y108_E1_ITEM_CD        = 2
Const Y108_E1_OLD_ITEM_NM    = 3
Const Y108_E1_OLD_ITEM_NM2   = 4
Const Y108_E1_OLD_ITEM_SPEC  = 5
Const Y108_E1_OLD_ITEM_SPEC2 = 6
Const Y108_E1_ITEM_ACCT      = 7
Const Y108_E1_ITEM_ACCT_NM   = 8
Const Y108_E1_ITEM_KIND      = 9
Const Y108_E1_ITEM_KIND_NM   = 10
Const Y108_E1_ITEM_LVL1      = 11
Const Y108_E1_ITEM_LVL1_NM   = 12
Const Y108_E1_ITEM_LVL2      = 13
Const Y108_E1_ITEM_LVL2_NM   = 14
Const Y108_E1_ITEM_LVL3      = 15
Const Y108_E1_ITEM_DERIVE    = 16
Const Y108_E1_ITEM_LVL_D     = 17
Const Y108_E1_ITEM_SEQNO     = 18
Const Y108_E1_ITEM_VER       = 19
Const Y108_E1_BASIC_CODE     = 20
Const Y108_E1_BASIC_CODE_NM  = 21
Const Y108_E1_NEW_ITEM_NM    = 22
Const Y108_E1_NEW_ITEM_NM2   = 23
Const Y108_E1_NEW_ITEM_SPEC  = 24
Const Y108_E1_NEW_ITEM_SPEC2 = 25
Const Y108_E1_REQ_ID         = 26
Const Y108_E1_REQ_ID_NM      = 27
Const Y108_E1_REQ_DT         = 28
Const Y108_E1_REQ_REASON     = 29
Const Y108_E1_REQ_END_DT     = 30
Const Y108_E1_END_DT         = 31
Const Y108_E1_TRANS_DT       = 32
Const Y108_E1_R_DT           = 33
Const Y108_E1_R_GRADE        = 34
Const Y108_E1_R_DESC         = 35
Const Y108_E1_R_PERSON       = 36
Const Y108_E1_R_PERSON_NM    = 37
Const Y108_E1_T_DT           = 38
Const Y108_E1_T_GRADE        = 39
Const Y108_E1_T_DESC         = 40
Const Y108_E1_T_PERSON       = 41
Const Y108_E1_T_PERSON_NM    = 42
Const Y108_E1_P_DT           = 43
Const Y108_E1_P_GRADE        = 44
Const Y108_E1_P_DESC         = 45
Const Y108_E1_P_PERSON       = 46
Const Y108_E1_P_PERSON_NM    = 47
Const Y108_E1_Q_DT           = 48
Const Y108_E1_Q_GRADE        = 49
Const Y108_E1_Q_DESC         = 50
Const Y108_E1_Q_PERSON       = 51
Const Y108_E1_STATUS_NM      = 52
Const Y108_E1_INTERNAL_CD    = 53
Const Y108_E1_ITEM_LVL3_NM    = 54

Dim iExportChangeItemNmReq
Dim iStrPreNextError

Set PY2G108 = Server.CreateObject("PY2G108.cCisChangeItemNmReqAppQuery")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PY2G108.Y_CIS_CHANGE_ITEM_NM_REQ_APP_SVR(gStrGlobalCollection, Request("PrevNextFlg"), Request("txtReqNo"), iExportChangeItemNmReq, iStrPreNextError)

If CheckSYSTEMError(Err,True) = True Then
	Set PY2G108 = Nothing
	Response.End
End If

If iStrPreNextError = "900011" Or iStrPreNextError = "900012" Then
	Call DisplayMsgBox(iStrPreNextError, vbOKOnly, "", "", I_MKSCRIPT)
End If
	
Set PY2G108 = Nothing

%>
<Script Language=vbscript>

With Parent.frm1

    .txtarReqNo.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_REQ_NO)))%>"
			
	.txtReqNo.Value		       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_REQ_NO)))%>"
	.txtStatus.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_STATUS_NM)))%>"
			
	.txtItemCd.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_ITEM_CD)))%>"			
	.txtItemNm.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_OLD_ITEM_NM)))%>"
	.txtItemNm2.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_OLD_ITEM_NM2)))%>"
	.txtSpec.Value	           = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_OLD_ITEM_SPEC)))%>"
	.txtSpec2.Value	           = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_OLD_ITEM_SPEC2)))%>"
			
	.txtItemAcct.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_ITEM_ACCT)))%>"
	.txtItemAcctNm.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_ITEM_ACCT_NM)))%>"
	.txtItemKind.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_ITEM_KIND)))%>"
	.txtItemKindNm.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_ITEM_KIND_NM)))%>"
	.txtItemLvl1.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_ITEM_LVL1)))%>"	
	.txtItemLvl1NM.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_ITEM_LVL1_NM)))%>"
	.txtItemLvl2.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_ITEM_LVL2)))%>"
	.txtItemLvl2Nm.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_ITEM_LVL2_NM)))%>"			
	.txtItemLvl3.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_ITEM_LVL3)))%>"
	.txtItemLvl3Nm.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_ITEM_LVL3_NM)))%>"
			
	If Trim("<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_ITEM_DERIVE)))%>") = "Y" Then
		.rdoDerive1.CHECKED	= TRUE
		.rdoDerive2.CHECKED	= FALSE
		.hrdoDerive.value = "Y"
	Else
	    .rdoDerive1.CHECKED	= FALSE
		.rdoDerive2.CHECKED	= TRUE
		.hrdoDerive.value = "N"
	End If	
					
	.txtSerialNo.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_ITEM_SEQNO)))%>"
	.txtBasicItem.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_BASIC_CODE)))%>"
	.txtBasicItemNm.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_BASIC_CODE_NM)))%>"
	.tXtItemVer.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_ITEM_VER)))%>"
			
	.txtNewItemNm.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_NEW_ITEM_NM)))%>"
	.txtNewItemNm2.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_NEW_ITEM_NM2)))%>"
	.txtNewSpec.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_NEW_ITEM_SPEC)))%>"
	.txtNewSpec2.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_NEW_ITEM_SPEC2)))%>"
			
	.txtReqId.Value	           = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_REQ_ID)))%>"
	.txtReqIdNm.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_REQ_ID_NM)))%>"
	.txtReqDt.Text             = "<%=UniConvDateDbToCompany(Trim(iExportChangeItemNmReq(Y108_E1_REQ_DT)),"")%>"
			
	If "<%=UniConvDateDbToCompany(Trim(iExportChangeItemNmReq(Y108_E1_END_DT)),"")%>" = "1900-01-01" Then
		.txtEndDt.Text = ""
	Else
		.txtEndDt.Text = "<%=UniConvDateDbToCompany(Trim(iExportChangeItemNmReq(Y108_E1_END_DT)),"")%>"
	End If
	If "<%=UniConvDateDbToCompany(Trim(iExportChangeItemNmReq(Y108_E1_TRANS_DT)),"")%>" = "1900-01-01" Then
		.txtTransDt.Text = ""
	Else
		.txtTransDt.Text = "<%=UniConvDateDbToCompany(Trim(iExportChangeItemNmReq(Y108_E1_TRANS_DT)),"")%>"
	End If
			
	.htxtRgrade.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_R_GRADE)))%>"
	.htxtTgrade.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_T_GRADE)))%>"
	.htxtPgrade.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_P_GRADE)))%>"
	.htxtQgrade.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_Q_GRADE)))%>"
			
	.htxtRdt.Value	           = "<%=UniConvDateDbToCompany(Trim(iExportChangeItemNmReq(Y108_E1_R_DT)),"")%>"
	.htxtTdt.Value	           = "<%=UniConvDateDbToCompany(Trim(iExportChangeItemNmReq(Y108_E1_T_DT)),"")%>"
	.htxtPdt.Value	           = "<%=UniConvDateDbToCompany(Trim(iExportChangeItemNmReq(Y108_E1_P_DT)),"")%>"
	.htxtQdt.Value	           = "<%=UniConvDateDbToCompany(Trim(iExportChangeItemNmReq(Y108_E1_Q_DT)),"")%>"
			
	.htxtRdesc.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_R_DESC)))%>"
	.htxtTdesc.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_T_DESC)))%>"
	.htxtPdesc.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_P_DESC)))%>"
	.htxtQdesc.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_Q_DESC)))%>"
			
	.htxtRperson.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_R_PERSON)))%>"
	.htxtTperson.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_T_PERSON)))%>"
	.htxtPperson.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_P_PERSON)))%>"
	.htxtQperson.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_Q_PERSON)))%>"
			
	.htxtReqReason.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_REQ_REASON)))%>"
			
	.htxtStatus.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_STATUS)))%>"
	.htxtInternalCd.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemNmReq(Y108_E1_INTERNAL_CD)))%>"
			
	.txtRgrade.Value	       = .htxtRgrade.Value
	.txtTgrade.Value	       = .htxtTgrade.Value
	.txtPgrade.Value	       = .htxtPgrade.Value
	.txtQgrade.Value	       = .htxtQgrade.Value
					
End with
	
Call Parent.DbQueryOk()

</Script>