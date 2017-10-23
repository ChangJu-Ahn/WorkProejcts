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

Dim PY2G101													'☆ : 조회용 ComProxy Dll 사용 변수 

Const Y101_E1_REQ_NO          = 0
Const Y101_E1_STATUS_NM       = 1
Const Y101_E1_ITEM_ACCT       = 2
Const Y101_E1_ITEM_ACCT_NM    = 3
Const Y101_E1_ITEM_KIND       = 4
Const Y101_E1_ITEM_KIND_NM    = 5
Const Y101_E1_ITEM_LVL1       = 6
Const Y101_E1_ITEM_LVL1_NM    = 7
Const Y101_E1_ITEM_LVL2       = 8
Const Y101_E1_ITEM_LVL2_NM    = 9
Const Y101_E1_ITEM_LVL3       = 10
Const Y101_E1_ITEM_LVL3_NM    = 11
Const Y101_E1_ITEM_SEQNO      = 12
Const Y101_E1_ITEM_DERIVE     = 13
Const Y101_E1_ITEM_LVL_D      = 14
Const Y101_E1_ITEM_VER        = 15
Const Y101_E1_BASIC_CODE      = 16
Const Y101_E1_BASIC_CODE_NM   = 17
Const Y101_E1_ITEM_CD         = 18
Const Y101_E1_ITEM_NM         = 19
Const Y101_E1_ITEM_NM2        = 20
Const Y101_E1_ITEM_SPEC       = 21
Const Y101_E1_ITEM_SPEC2      = 22
Const Y101_E1_ITEM_UNIT       = 23
Const Y101_E1_PUR_TYPE        = 24
Const Y101_E1_PUR_GROUP       = 25
Const Y101_E1_PUR_GROUP_NM    = 26
Const Y101_E1_PUR_VENDOR      = 27
Const Y101_E1_PUR_VENDOR_NM   = 28
Const Y101_E1_UNIFY_PUR_FLAG  = 29
Const Y101_E1_UNIT_WEIGHT     = 30
Const Y101_E1_UNIT_OF_WEIGHT  = 31
Const Y101_E1_GROSS_WEIGHT    = 32
Const Y101_E1_GROSS_UNIT      = 33
Const Y101_E1_CBM             = 34
Const Y101_E1_CBM_DESCRIPTION = 35
Const Y101_E1_HS_CODE         = 36
Const Y101_E1_HS_CODE_NM      = 37
Const Y101_E1_VALID_FROM_DT   = 38
Const Y101_E1_VALID_TO_DT     = 39
Const Y101_E1_DOC_NO          = 40
Const Y101_E1_FILE_NM         = 41
Const Y101_E1_REQ_ID          = 42
Const Y101_E1_REQ_ID_NM       = 43
Const Y101_E1_REQ_DT          = 44
Const Y101_E1_REQ_REASON      = 45
Const Y101_E1_REMARK          = 46
Const Y101_E1_REQ_END_DT      = 47
Const Y101_E1_STATUS          = 48
Const Y101_E1_INTERNAL_CD     = 49

Dim iExportNewItemReq
Dim iStrPreNextError

Set PY2G101 = Server.CreateObject("PY2G101.cCisNewItemReqQuery")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call PY2G101.Y_CIS_NEW_ITEM_REQ_SVR(gStrGlobalCollection, Request("PrevNextFlg"), Request("txtReqNo"), iExportNewItemReq, iStrPreNextError)

If CheckSYSTEMError(Err,True) = True Then
	Set PY2G112 = Nothing
	Response.End
End If

If iStrPreNextError = "900011" Or iStrPreNextError = "900012" Then
	Call DisplayMsgBox(iStrPreNextError, vbOKOnly, "", "", I_MKSCRIPT)
	Set PY2G101 = Nothing
	Response.End
End If
	
Set PY2G101 = Nothing

%>
<Script Language=vbscript>
With Parent.frm1
    .txtarReqNo.Value		   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_REQ_NO)))%>"
	.txtarItemCd.Value	       = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_CD)))%>"
	.txtarItemNm.Value	       = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_NM)))%>"
			
	.txtReqNo.Value		       = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_REQ_NO)))%>"
	.txtStatus.Value		   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_STATUS_NM)))%>"
						
	.cboItemAcct.Value		   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_ACCT)))%>"			
	.txtItemKind.Value		   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_KIND)))%>"
	.txtItemKindNm.Value	   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_KIND_NM)))%>"
	.txtItemLvl1.Value		   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_LVL1)))%>"	
	.txtItemLvl1Nm.Value	   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_LVL1_NM)))%>"
	.txtItemLvl2.Value		   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_LVL2)))%>"
	.txtItemLvl2Nm.Value	   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_LVL2_NM)))%>"			
	.txtItemLvl3.Value		   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_LVL3)))%>"
	.txtItemLvl3Nm.Value	   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_LVL3_NM)))%>"
	.txtSerialNo.Value		   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_SEQNO)))%>"
			
	If Trim("<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_DERIVE)))%>") = "Y" Then
		.rdoDerive1.CHECKED	= TRUE
		.rdoDerive2.CHECKED	= FALSE
		.hrdoDerive.value = "Y"
	Else
	    .rdoDerive1.CHECKED	= FALSE
		.rdoDerive2.CHECKED	= TRUE
		.hrdoDerive.value     = "N"
	End If
			
	.cboItemVer.Value	       = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_VER)))%>"			
	.txtBasicItem.Value	       = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_BASIC_CODE)))%>"			
	.txtBasicItemNm.Value	   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_BASIC_CODE_NM)))%>"
					
	.txtItemNm.Value	       = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_NM)))%>"
	.txtItemNm2.Value		   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_NM2)))%>"
	.txtSpec.Value	           = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_SPEC)))%>"
	.txtSpec2.Value	           = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_SPEC2)))%>"						
	.txtItemUnit.Value   	   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_UNIT)))%>"
	.cboPurType.Value		   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_PUR_TYPE)))%>"			
	.txtPurGroup.Value	       = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_PUR_GROUP)))%>"
	.txtPurGroupNm.Value	   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_PUR_GROUP_NM)))%>"
	.txtPurVendor.Value	       = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_PUR_VENDOR)))%>"
	.txtPurVendorNm.Value	   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_PUR_VENDOR_NM)))%>"
			
	If Trim("<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_UNIFY_PUR_FLAG)))%>") = "Y" Then
		.rdoUnifyPurFlg1.CHECKED	= TRUE
		.rdoUnifyPurFlg2.CHECKED	= FALSE
		.hrdoUnifyPurFlg.value = "Y"
	Else
	    .rdoUnifyPurFlg1.CHECKED	= FALSE
		.rdoUnifyPurFlg2.CHECKED	= TRUE
		.hrdoUnifyPurFlg.value     = "N"
	End If 	
			
	.txtNetWeight.Value	      = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_UNIT_WEIGHT)))%>"
	.txtNetWeightUnit.Value   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_UNIT_OF_WEIGHT)))%>"			
	.txtGrossWeight.Value	  = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_GROSS_WEIGHT)))%>"
	.txtGrossWeightUnit.Value = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_GROSS_UNIT)))%>"
	.txtCBM.Value             = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_CBM)))%>"	
	.txtCBMInfo.Value		  = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_CBM_DESCRIPTION)))%>"
	.txtHSCd.Value            = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_HS_CODE)))%>"
	.txtHSNm.Value            = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_HS_CODE_NM)))%>"		
	.txtValidFromDt.Text      = "<%=UniConvDateDbToCompany(Trim(iExportNewItemReq(Y101_E1_VALID_FROM_DT)),"")%>"
	.txtValidToDt.Text	      = "<%=UniConvDateDbToCompany(Trim(iExportNewItemReq(Y101_E1_VALID_TO_DT)),"")%>"
	.txtDocNo.Value	          = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_DOC_NO)))%>"
	.txtFileNm.Value	      = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_FILE_NM)))%>"
	.htxtFilePath.Value	      = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_FILE_PATH)))%>"
	.htxtIdFile.Value	      = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ID_FILE)))%>"
	.txtreq_user.Value	          = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_REQ_ID)))%>"
	.txtreq_user_Nm.Value		  = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_REQ_ID_NM)))%>"
	//.txtReqDt.Text            = "<%=UniConvDateDbToCompany(Trim(iExportNewItemReq(Y101_E1_REQ_DT)),"")%>"
	.txtReqDt.Text            = "<%=UniConvDateDbToCompany(iExportNewItemReq(Y101_E1_REQ_DT),"")%>"
			
	If "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_REQ_END_DT)))%>" = "1900-01-01" Then
		.txtEndReqDt.Text   = ""
	Else
		.txtEndReqDt.Text = "<%=UniConvDateDbToCompany(Trim(iExportNewItemReq(Y101_E1_REQ_END_DT)),"")%>"
	End If
		.txtEndReqDt.Text = "<%=UniConvDateDbToCompany(Trim(iExportNewItemReq(Y101_E1_REQ_END_DT)),"")%>"	
	.htxtItemCd.Value	       = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_ITEM_CD)))%>"		
	.htxtReqReason.Value	   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_REQ_REASON)))%>"			
	.htxtRemark.Value		   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_REMARK)))%>"			
	.htxtStatus.Value		   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_STATUS)))%>"
	.htxtInternalCd.Value	   = "<%=ConvSPChars(Trim(iExportNewItemReq(Y101_E1_INTERNAL_CD)))%>"
			
End with

Call Parent.DbQueryOk()

</Script>