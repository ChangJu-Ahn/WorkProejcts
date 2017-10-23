<%@ LANGUAGE="VBSCRIPT" %>
<% Option explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : b1252mb1
'*  4. Program Name         : 구매조직등록 
'*  5. Program Desc         : 구매조직등록 
'*  6. Component List       : PB2G629.cBLookupPurOrgS / PB2G621.cBMaintPurOrgS
'*  7. Modified date(First) : 2000/06/08
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Dim lgOpModeCRUD
	
	Call LoadBasisGlobalInf()
    Call HideStatusWnd                                                               '☜: Hide Processing message
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    lgOpModeCRUD  = Request("txtMode")                                               '☜: Read Operation Mode (CRUD)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)				                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	Dim B26029

	Dim I1_b_pur_org
	Dim I2_b_pur_org
	Dim E1_b_pur_org
	
	Dim strDefFrDate
	Dim strDefToDate
	
	'EXPORT View 상수 - 구매조직 
	Const M003_e1_pur_org_nm = 0
	Const M003_e1_usage_flg = 1
	Const M003_e1_valid_fr_dt = 2
	Const M003_e1_valid_to_dt = 3
	Const M003_e1_ext1_cd = 4
	Const M003_e1_ext2_cd = 5
	Const M003_e1_ext3_cd = 6
	Const M003_e1_ext4_cd = 7

	strDefFrDate = "1900-01-01"
	strDefToDate = "2999-12-31"

	I1_b_pur_org = UCase(Trim(Request("txtORGCd1")))
	I2_b_pur_org = Trim(Request("txtORGCd1"))
    Set B26029 = Server.CreateObject("PB2G629.cBLookupPurOrgS")    
	
    if CheckSYSTEMError(Err,True) = true then 
		Exit Sub
	End If

	Call B26029.B_LOOKUP_PUR_ORG(gStrGlobalCollection, _
								 I1_b_pur_org, _
								 E1_b_pur_org)

 	if CheckSYSTEMError2(Err,True,"","","","","") = true then 		
	    Set B26029 = Nothing															'☜: Unload Comproxy
		
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "	parent.frm1.txtORGNm1.value	= """" " & vbCr
		Response.Write "</Script>" & vbCr
		
		Exit Sub
	end if

    Set B26029 = Nothing															'☜: Unload Comproxy
	
	Response.Write "<Script Language=vbscript>"																	& vbCr
	Response.Write "With parent"																				& vbCr
	Response.Write "	.frm1.txtORGCd1.value	= """ & ConvSPChars(I2_b_pur_org) & """" 						& vbCr
	Response.Write "	.frm1.txtORGNm1.value	= """ & ConvSPChars(E1_b_pur_org(M003_e1_pur_org_nm)) & """" 	& vbCr
	Response.Write "	.frm1.txtORGCd2.value	= """ & ConvSPChars(I2_b_pur_org) & """" 						& vbCr
	Response.Write "	.frm1.txtORGNm2.value	= """ & ConvSPChars(E1_b_pur_org(M003_e1_pur_org_nm)) & """" 	& vbCr
	Response.Write "	.frm1.txtFrDt.text		= """ & UNIDateClientFormat(E1_b_pur_org(M003_e1_valid_fr_dt)) & """" & vbCr
	Response.Write "	.frm1.txtToDt.text		= """ & UNIDateClientFormat(E1_b_pur_org(M003_e1_valid_to_dt)) & """" & vbCr
	If Trim(E1_b_pur_org(M003_e1_usage_flg)) = "Y" Then
		Response.Write "	.frm1.rdoUseflg(0).checked=true " 	& vbCr
	Else
		Response.Write "	.frm1.rdoUseflg(1).checked=true " 	& vbCr	
	End If
	Response.Write "	parent.lgNextNo = """"" & vbCr	
	Response.Write "	parent.lgPrevNo = """"" & vbCr	
	Response.Write "	parent.DbQueryOk	"	& vbCr	
    Response.Write "End With " & vbCr
    Response.Write "</Script>" & vbCr

End Sub    

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear    				   					                                 '☜: Clear Error status
	
	Const M013_I1_PurOrg = 0
	Const M013_I1_PurOrgNm = 1
	Const M013_I1_UsageFlg = 2
	Const M013_I1_ValidFrDt = 3
	Const M013_I1_ValidToDt = 4
	Const M013_I1_EXT1_CD = 5
	Const M013_I1_EXT2_CD = 6
	Const M013_I1_EXT3_CD = 7
	Const M013_I1_EXT4_CD = 8

	Dim PB2G621
	Dim I1_pur_org
	Dim lgIntFlgMode
										
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
	
	Redim I1_pur_org(8)
	
	I1_pur_org(M013_I1_PurOrg) 		= UCase(Trim(Request("txtORGCd2")))
	I1_pur_org(M013_I1_PurOrgNm) 	= Trim(Request("txtORGNm2"))
	I1_pur_org(M013_I1_UsageFlg) 	= Trim(Request("txtuseflg"))
	I1_pur_org(M013_I1_ValidFrDt) 	= UniconvDate(Request("txtFrDt"))
	I1_pur_org(M013_I1_ValidToDt) 	= UniconvDate(Request("txtToDt"))
	I1_pur_org(M013_I1_EXT1_CD) 	= ""
	I1_pur_org(M013_I1_EXT2_CD) 	= ""
	I1_pur_org(M013_I1_EXT3_CD) 	= ""
	I1_pur_org(M013_I1_EXT4_CD) 	= ""
		
    Set PB2G621 = Server.CreateObject("PB2G621.cBMaintPurOrgS")    

    If CheckSYSTEMError(Err,True) = true Then 
		Exit Sub
	End If

    If lgIntFlgMode = OPMD_CMODE Then
		Call PB2G621.B_MAINT_PUR_ORG_SVR(gStrGlobalCollection,"CREATE",I1_pur_org)
	Else
		Call PB2G621.B_MAINT_PUR_ORG_SVR(gStrGlobalCollection,"UPDATE",I1_pur_org)
    End If

 	If CheckSYSTEMError2(Err,True,"","","","","") = true Then 		
	    Set b26021 = Nothing                                                    '☜: Unload Comproxy
		Exit Sub
 	End if

    Set b26021 = Nothing                                                 		'☜: Unload Comproxy

	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent"		 		& vbCr	
	Response.Write "	.DbSaveOk"		 		& vbCr
    Response.Write "End With "		 			& vbCr
    Response.Write "</Script>"		 			& vbCr

End Sub

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

	On Error Resume Next  				
    Err.Clear																		'☜: Protect system from crashing
	
	Const M013_I1_PurOrg = 0

	Dim b26021
	Dim I1_pur_org
	
	Redim I1_pur_org(0)
	
	I1_pur_org(M013_I1_PurOrg) = UCase(Trim(Request("txtORGCd1")))

    Set b26021 = Server.CreateObject("PB2G621.cBMaintPurOrgS")

    if CheckSYSTEMError(Err,True) = true then 
		Exit Sub
	End If

	Call b26021.B_MAINT_PUR_ORG_SVR(gStrGlobalCollection,"DELETE",I1_pur_org)
	
 	if CheckSYSTEMError2(Err,true,"","","","","") = true then 		
	    Set b26021 = Nothing                                                    '☜: Unload Comproxy
		Exit Sub
 	end if

    Set b26021 = Nothing                                                   '☜: Unload Comproxy

	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "Call parent.DbDeleteOk()"	& vbCr
    Response.Write "</Script>"		 			& vbCr

End Sub

%>

