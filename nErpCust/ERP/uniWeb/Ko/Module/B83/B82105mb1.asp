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

Dim PY2G105													'☆ : 조회용 ComProxy Dll 사용 변수 

Const Y105_E1_REQ_NO          = 0
Const Y105_E1_ITEM_CD         = 1
Const Y105_E1_ITEM_NM         = 2
Const Y105_E1_ITEM_NM2        = 3
Const Y105_E1_ITEM_SPEC       = 4
Const Y105_E1_ITEM_SPEC2      = 5
Const Y105_E1_VERSION_CHG     = 6
Const Y105_E1_END_ITEM_FLAG   = 7
Const Y105_E1_ITEM_UNIT       = 8
Const Y105_E1_BASIC_CODE      = 9
Const Y105_E1_BASIC_CODE_NM   = 10
Const Y105_E1_PUR_TYPE        = 11
Const Y105_E1_PUR_GROUP       = 12
Const Y105_E1_PUR_GROUP_NM    = 13
Const Y105_E1_PUR_VENDOR      = 14
Const Y105_E1_PUR_VENDOR_NM   = 15
Const Y105_E1_UNIFY_PUR_FLAG  = 16 
Const Y105_E1_UNIT_WEIGHT     = 17
Const Y105_E1_UNIT_OF_WEIGHT  = 18
Const Y105_E1_GROSS_WEIGHT    = 19
Const Y105_E1_GROSS_UNIT      = 20
Const Y105_E1_CBM             = 21
Const Y105_E1_CBM_DESCRIPTION = 22
Const Y105_E1_HS_CODE         = 23
Const Y105_E1_HS_CODE_NM      = 24
Const Y105_E1_VALID_FROM_DT   = 25
Const Y105_E1_VALID_TO_DT     = 26
Const Y105_E1_DOC_NO          = 27
Const Y105_E1_FILE_NM         = 28
Const Y105_E1_REQ_ID          = 29
Const Y105_E1_REQ_ID_NM       = 30
Const Y105_E1_REQ_DT          = 31
Const Y105_E1_REQ_REASON      = 32
Const Y105_E1_REMARK          = 33
Const Y105_E1_REQ_END_DT      = 34
Const Y105_E1_STATUS          = 35
Const Y105_E1_STATUS_NM       = 36
Const Y105_E1_INTERNAL_CD     = 37
Const Y105_E1_END_DT          = 38
Const Y105_E1_TRANS_DT        = 39
Const Y105_E1_R_DT            = 40 
Const Y105_E1_R_GRADE         = 41
Const Y105_E1_R_DESC          = 42
Const Y105_E1_R_PERSON        = 43
Const Y105_E1_R_PERSON_NM     = 44
Const Y105_E1_T_DT            = 45
Const Y105_E1_T_GRADE         = 46
Const Y105_E1_T_DESC          = 47
Const Y105_E1_T_PERSON        = 48
Const Y105_E1_T_PERSON_NM     = 49
Const Y105_E1_P_DT            = 50
Const Y105_E1_P_GRADE         = 51
Const Y105_E1_P_DESC          = 52
Const Y105_E1_P_PERSON        = 53
Const Y105_E1_P_PERSON_NM     = 54
Const Y105_E1_Q_DT            = 55
Const Y105_E1_Q_GRADE         = 56
Const Y105_E1_Q_DESC          = 57
Const Y105_E1_Q_PERSON        = 58
Const Y105_E1_ITEM_ACCT       = 59
Const Y105_E1_ITEM_KIND       = 60

Dim iExportChangeItemReq
Dim iStrPreNextError

Set PY2G105 = Server.CreateObject("PY2G105.cCisChangeItemReqAppQuery")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PY2G105.Y_CIS_CHANGE_ITEM_REQ_APP_SVR(gStrGlobalCollection, Request("PrevNextFlg"), Request("txtReqNo"), iExportChangeItemReq, iStrPreNextError)

If CheckSYSTEMError(Err,True) = True Then
	Set PY2G105 = Nothing
	Response.End
End If

If iStrPreNextError = "900011" Or iStrPreNextError = "900012" Then
	Call DisplayMsgBox(iStrPreNextError, vbOKOnly, "", "", I_MKSCRIPT)
End If
	
Set PY2G105 = Nothing

%>
<Script Language=vbscript>
With Parent.frm1
    .txtarReqNo.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_REQ_NO)))%>"				
	.txtReqNo.Value		       = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_REQ_NO)))%>"
	.txtItemCd.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_ITEM_CD)))%>"
	.txtItemNm.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_ITEM_NM)))%>"
	.txtItemNm2.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_ITEM_NM2)))%>"
	.txtSpec.Value	           = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_ITEM_SPEC)))%>"
	.txtSpec2.Value	           = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_ITEM_SPEC2)))%>"
			
	If Trim("<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_VERSION_CHG)))%>") = "Y" Then
		.rdoChgVer1.Checked	= True
		.rdoChgVer2.Checked	= False
		.hrdoChgVer.value = "Y" 
	Else
	    .rdoChgVer1.Checked	= False
		.rdoChgVer2.Checked	= True
		.hrdoChgVer.value      = "N" 
	End If 	
			
	If Trim("<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_END_ITEM_FLAG)))%>") = "Y" Then
		.rdoEndItem1.Checked = True
		.rdoEndItem2.Checked = False
		.hrdoEndItem.value = "Y" 
	Else
	    .rdoEndItem1.Checked = False
		.rdoEndItem2.Checked = True
		.hrdoEndItem.value = "N" 
	End If
			
	.txtBasicItem.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_BASIC_CODE)))%>"			
	.txtBasicItemNm.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_BASIC_CODE_NM)))%>"
	.txtItemUnit.Value   	   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_ITEM_UNIT)))%>"
	.cboPurType.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_PUR_TYPE)))%>"
	.txtPurGroup.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_PUR_GROUP)))%>"
	.txtPurGroupNm.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_PUR_GROUP_NM)))%>"
	.txtPurVendor.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_PUR_VENDOR)))%>"
	.txtPurVendorNm.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_PUR_VENDOR_NM)))%>"
			
	If Trim("<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_UNIFY_PUR_FLAG)))%>") = "Y" Then
		.rdoUnifyPurFlg1.Checked	= True
		.rdoUnifyPurFlg2.Checked	= False
		.hrdoUnifyPurFlg.value = "Y" 
	Else
	    .rdoUnifyPurFlg1.Checked	= False
		.rdoUnifyPurFlg2.Checked	= True
		.hrdoUnifyPurFlg.value = "N" 
	End If 	
			
	.txtNetWeight.Value	      = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_UNIT_WEIGHT)))%>"
	.txtNetWeightUnit.Value   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_UNIT_OF_WEIGHT)))%>"			
	.txtGrossWeight.Value	  = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_GROSS_WEIGHT)))%>"
	.txtGrossWeightUnit.Value = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_GROSS_UNIT)))%>"
	.txtCBM.Value             = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_CBM)))%>"	
	.txtCBMInfo.Value		  = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_CBM_DESCRIPTION)))%>"
	.txtHSCd.Value            = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_HS_CODE)))%>"
	.txtHSNm.Value            = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_HS_CODE_NM)))%>"		
	.txtValidFromDt.Text      = "<%=UniConvDateDbToCompany(Trim(iExportChangeItemReq(Y105_E1_VALID_FROM_DT)),"")%>"
	.txtValidToDt.Text	      = "<%=UniConvDateDbToCompany(Trim(iExportChangeItemReq(Y105_E1_VALID_TO_DT)),"")%>"
	.txtDocNo.Value	          = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_DOC_NO)))%>"
	.txtFileNm.Value	      = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_FILE_NM)))%>"
	.txtReqId.Value	          = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_REQ_ID)))%>"
	.txtReqIdNm.Value		  = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_REQ_ID_NM)))%>"
	.txtReqDt.Text            = "<%=UniConvDateDbToCompany(Trim(iExportChangeItemReq(Y105_E1_REQ_DT)),"")%>"
			
	
	.txtEndDt.Text				= "<%=UniConvDateDbToCompany(Trim(iExportChangeItemReq(Y105_E1_END_DT)),"")%>"
	.txtTransDt.Text			= "<%=UniConvDateDbToCompany(Trim(iExportChangeItemReq(Y105_E1_TRANS_DT)),"")%>"	
	.htxtStatus.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_STATUS)))%>"
	.txtStatus.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_STATUS_NM)))%>"
	.htxtInternalCd.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_INTERNAL_CD)))%>"
			
	.htxtReqReason.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_REQ_REASON)))%>"
	.htxtRemark.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_REMARK)))%>"
			
	.htxtRDt.Value		       = "<%=UniConvDateDbToCompany(Trim(iExportChangeItemReq(Y105_E1_R_DT)),"")%>"
	.htxtRGrade.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_R_GRADE)))%>"
	.htxtRDesc.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_R_DESC)))%>"
	.htxtRPerson.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_R_PERSON)))%>"
	.htxtRPersonNm.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_R_PERSON_NM)))%>"
			
	.htxtTDt.Value		       = "<%=UniConvDateDbToCompany(Trim(iExportChangeItemReq(Y105_E1_T_DT)),"")%>"
	.htxtTGrade.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_T_GRADE)))%>"
	.htxtTDesc.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_T_DESC)))%>"
	.htxtTPerson.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_T_PERSON)))%>"
	.htxtTPersonNm.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_T_PERSON_NM)))%>"
			
	.htxtPDt.Value		       = "<%=UniConvDateDbToCompany(Trim(iExportChangeItemReq(Y105_E1_P_DT)),"")%>"
	.htxtPGrade.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_P_GRADE)))%>"
	.htxtPDesc.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_P_DESC)))%>"
	.htxtPPerson.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_P_PERSON)))%>"
	.htxtPPersonNm.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_P_PERSON_NM)))%>"
			
	.htxtQDt.Value		       = "<%=UniConvDateDbToCompany(Trim(iExportChangeItemReq(Y105_E1_Q_DT)),"")%>"
	.htxtQGrade.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_Q_GRADE)))%>"
	.htxtQDesc.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_Q_DESC)))%>"
	.htxtQPerson.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_Q_PERSON)))%>"
	.htxtQPersonNm.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_Q_PERSON_NM)))%>"
			
	.htxtItemAcct.value        = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_ITEM_ACCT)))%>"
	.htxtItemKind.value        = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y105_E1_ITEM_KIND)))%>"
			
	.txtRgrade.Value	       = .htxtRgrade.Value
	.txtTgrade.Value	       = .htxtTgrade.Value
	.txtPgrade.Value	       = .htxtPgrade.Value
	.txtQgrade.Value	       = .htxtQgrade.Value
					
End with
	
Call Parent.DbQueryOk()

</Script>