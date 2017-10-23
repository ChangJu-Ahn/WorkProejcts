<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<% Call loadInfTB19029B("I", "*", "NOCOOKIE","MB") %>
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

Dim PY2G104													'☆ : 조회용 ComProxy Dll 사용 변수 

Const Y104_E1_REQ_NO         = 0
Const Y104_E1_ITEM_CD        = 1
Const Y104_E1_ITEM_NM        = 2
Const Y104_E1_ITEM_NM2       = 3
Const Y104_E1_ITEM_SPEC      = 4
Const Y104_E1_ITEM_SPEC2     = 5
Const Y104_E1_VERSION_CHG    = 6
Const Y104_E1_END_ITEM_FLAG  = 7
Const Y104_E1_ITEM_UNIT      = 8
Const Y104_E1_BASIC_CODE     = 9
Const Y104_E1_BASIC_CODE_NM  = 10
Const Y104_E1_PUR_TYPE       = 11
Const Y104_E1_PUR_GROUP      = 12
Const Y104_E1_PUR_GROUP_NM   = 13
Const Y104_E1_PUR_VENDOR     = 14
Const Y104_E1_PUR_VENDOR_NM  = 15
Const Y104_E1_UNIFY_PUR_FLAG = 16
Const Y104_E1_UNIT_WEIGHT    = 17
Const Y104_E1_UNIT_OF_WEIGHT = 18
Const Y104_E1_GROSS_WEIGHT   = 19
Const Y104_E1_GROSS_UNIT     = 20
Const Y104_E1_CBM            = 21
Const Y104_E1_CBM_DESCRIPTION= 22
Const Y104_E1_HS_CODE        = 23
Const Y104_E1_HS_CODE_NM     = 24
Const Y104_E1_VALID_FROM_DT  = 25
Const Y104_E1_VALID_TO_DT    = 26
Const Y104_E1_DOC_NO         = 27
Const Y104_E1_FILE_NM        = 28
Const Y104_E1_REQ_ID         = 29
Const Y104_E1_REQ_ID_NM      = 30
Const Y104_E1_REQ_DT         = 31
Const Y104_E1_REQ_REASON     = 32
Const Y104_E1_REMARK         = 33
Const Y104_E1_REQ_END_DT     = 34
Const Y104_E1_STATUS         = 35
Const Y104_E1_STATUS_NM      = 36
Const Y104_E1_INTERNAL_CD    = 37
Const Y104_E1_END_DT         = 38
Const Y104_E1_TRANS_DT       = 39
Const Y104_E1_R_GRADE        = 40
Const Y104_E1_T_GRADE        = 41
Const Y104_E1_P_GRADE        = 42
Const Y104_E1_Q_GRADE        = 43

Dim iExportChangeItemReq
Dim iStrPreNextError

Set PY2G104 = Server.CreateObject("PY2G104.cCisChangeItemReqQuery")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PY2G104.Y_CIS_CHANGE_ITEM_REQ_SVR(gStrGlobalCollection, Request("PrevNextFlg"), Request("txtReqNo"), iExportChangeItemReq, iStrPreNextError)

If CheckSYSTEMError(Err,True) = True Then
	Set PY2G104 = Nothing
	Response.End
End If

If iStrPreNextError = "900011" Or iStrPreNextError = "900012" Then
	Call DisplayMsgBox(iStrPreNextError, vbOKOnly, "", "", I_MKSCRIPT)
End If
	
Set PY2G104 = Nothing

%>
<Script Language=vbscript>
With Parent.frm1
    .txtarReqNo.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_REQ_NO)))%>"				
	.txtReqNo.Value		       = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_REQ_NO)))%>"
	.txtItemCd.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_ITEM_CD)))%>"
	.txtItemNm.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_ITEM_NM)))%>"
	.txtItemNm2.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_ITEM_NM2)))%>"
	.txtSpec.Value	           = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_ITEM_SPEC)))%>"
	.txtSpec2.Value	           = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_ITEM_SPEC2)))%>"
			
	If Trim("<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_VERSION_CHG)))%>") = "Y" Then
		.rdoChgVer1.Checked	= True
		.rdoChgVer2.Checked	= False
		.hrdoChgVer.value      = "Y" 
	Else
	    .rdoChgVer1.Checked	= False
		.rdoChgVer2.Checked	= True
		.hrdoChgVer.value      = "N" 
	End If 	
			
	If Trim("<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_END_ITEM_FLAG)))%>") = "Y" Then
		.rdoEndItem1.Checked	= True
		.rdoEndItem2.Checked	= False
		.hrdoEndItem.value = "Y" 
	Else
	    .rdoEndItem1.Checked	= False
		.rdoEndItem2.Checked	= True
		.hrdoEndItem.value = "N" 
	End If   
			
	.txtBasicItem.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_BASIC_CODE)))%>"			
	.txtBasicItemNm.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_BASIC_CODE_NM)))%>"
	.txtItemUnit.Value   	   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_ITEM_UNIT)))%>"
	.cboPurType.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_PUR_TYPE)))%>"
	.txtPurGroup.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_PUR_GROUP)))%>"
	.txtPurGroupNm.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_PUR_GROUP_NM)))%>"
	.txtPurVendor.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_PUR_VENDOR)))%>"
	.txtPurVendorNm.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_PUR_VENDOR_NM)))%>"
			
	If Trim("<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_UNIFY_PUR_FLAG)))%>") = "Y" Then
		.rdoUnifyPurFlg1.Checked = True
		.rdoUnifyPurFlg2.Checked = False
		.hrdoUnifyPurFlg.value   = "Y" 
	Else
	    .rdoUnifyPurFlg1.Checked = False
		.rdoUnifyPurFlg2.Checked = True
		.hrdoUnifyPurFlg.value   = "N" 
	End If 	
			
	.txtNetWeight.Value	      = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_UNIT_WEIGHT)))%>"
	.txtNetWeightUnit.Value   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_UNIT_OF_WEIGHT)))%>"			
	.txtGrossWeight.Value	  = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_GROSS_WEIGHT)))%>"
	.txtGrossWeightUnit.Value = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_GROSS_UNIT)))%>"
	.txtCBM.Value             = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_CBM)))%>"	
	.txtCBMInfo.Value		  = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_CBM_DESCRIPTION)))%>"
	.txtHSCd.Value            = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_HS_CODE)))%>"
	.txtHSNm.Value            = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_HS_CODE_NM)))%>"		
	.txtValidFromDt.Text      = "<%=UniConvDateDbToCompany(Trim(iExportChangeItemReq(Y104_E1_VALID_FROM_DT)),"")%>"
	.txtValidToDt.Text	      = "<%=UniConvDateDbToCompany(Trim(iExportChangeItemReq(Y104_E1_VALID_TO_DT)),"")%>"
	.txtDocNo.Value	          = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_DOC_NO)))%>"
	.txtFileNm.Value	      = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_FILE_NM)))%>"
	.txtreq_user.Value	          = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_REQ_ID)))%>"
	.txtreq_user_nm.Value		  = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_REQ_ID_NM)))%>"
	.txtReqDt.Text            = "<%=UniConvDateDbToCompany(Trim(iExportChangeItemReq(Y104_E1_REQ_DT)),"")%>"
			
	
	.txtEndDt.Text = "<%=UniConvDateDbToCompany(Trim(iExportChangeItemReq(Y104_E1_END_DT)),"")%>"
	.txtTransDt.Text = "<%=UniConvDateDbToCompany(Trim(iExportChangeItemReq(Y104_E1_TRANS_DT)),"")%>"
	
			
	.htxtStatus.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_STATUS)))%>"
	.txtStatus.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_STATUS_NM)))%>"
	.htxtInternalCd.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_INTERNAL_CD)))%>"
			
	.txtRgrade.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_R_GRADE)))%>"
	.txtTgrade.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_T_GRADE)))%>"
	.txtPgrade.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_P_GRADE)))%>"
	.txtQgrade.Value	       = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_Q_GRADE)))%>" 
			
	.htxtReqReason.Value	   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_REQ_REASON)))%>"			
	.htxtRemark.Value		   = "<%=ConvSPChars(Trim(iExportChangeItemReq(Y104_E1_REMARK)))%>"	
			
End with
	
Call Parent.DbQueryOk()

</Script>