<%@ LANGUAGE=VBSCript%>
<% Option explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : b1252qb1
'*  4. Program Name         : 구매조직조회 
'*  5. Program Desc         : 구매조직조회 
'*  6. Component List       : PB2G628.cBListPurOrgS
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

    lgOpModeCRUD  = Request("txtMode")                                               '☜: Read Operation Mode (CRUD)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call  SubBizQueryMulti()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	Const C_SHEETMAXROWS_D  = 100
	
    Const M001_E1_pur_org = 0
    Const M001_E1_pur_org_nm = 1

    Const M001_EG1_E1_pur_org = 0
    Const M001_EG1_E1_usage_flg = 1
    Const M001_EG1_E1_pur_org_nm = 2
    Const M001_EG1_E1_valid_fr_dt = 3
    Const M001_EG1_E1_valid_to_dt = 4
    Const M001_EG1_E1_ext1_cd = 5
    Const M001_EG1_E1_ext2_cd = 6
    Const M001_EG1_E1_ext3_cd = 7
    Const M001_EG1_E1_ext4_cd = 8

	Dim pB26028
	Dim iStrPrevKey	
	Dim iStrOrgCd
	Dim iStrUsageFlg
    Dim E1_b_pur_org
    Dim EG1_export_group
	Dim LngMaxRow
	Dim strData
	Dim strDt
	Dim StrNextKey
	Dim LngRow
	Dim strDefFrDate
	Dim strDefToDate
    Dim PvArr
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iStrPrevKey     = Trim(Request("lgStrPrevKey"))                                        '☜: Next Key	    
	iStrOrgCd 		= UCase(Trim(Request("txtORGCd")))
	iStrUsageFlg 	= Trim(Request("rdoUseflg"))
	
    Set pB26028 = Server.CreateObject("PB2G628.cBListPurOrgS")

    if CheckSYSTEMError(Err,True) = true then 
		Exit Sub
	End If

    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    Call pB26028.B_LIST_PUR_ORG_SVR(gStrGlobalCollection, _
								 C_SHEETMAXROWS_D, _
								 iStrOrgCd, _
								 iStrUsageFlg,_ 
								 iStrPrevKey, _
								 E1_b_pur_org, _
								 EG1_export_group)
    
 	if CheckSYSTEMError2(Err,True,"","","","","") = true then 		
	    set pB26028 = nothing
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "	Parent.frm1.txtORGCd.focus " & vbCr
		Response.Write "	Set Parent.gActiveElement = Parent.document.activeElement    " & vbCr
		Response.Write "</Script>" & vbCr
		response.End														'☜: 비지니스 로직 처리를 종료함 
 	end if

    set pB26028 = nothing
    
	Response.Write "<Script Language=vbscript>"														& vbCr
	Response.Write "With parent"																	& vbCr
	Response.Write "	If Trim(.frm1.txtOrgCd.Value) <> """" Then"									& vbCr
	Response.Write "	  	.frm1.txtORGNm.Value = """ & ConvSPChars(E1_b_pur_org(M001_E1_pur_org_nm)) & """"	& vbCr
    Response.Write "	End If " 																	& vbCr
    Response.Write "End With " 																		& vbCr
    Response.Write "</Script>" 																		& vbCr
 
 	LngMaxRow = Request("txtMaxRows")											'Save previous Maxrow  

	strDefFrDate = uniConvDate(strDefFrDate)
	strDefToDate = uniConvDate(strDefToDate)
    ReDim PvArr(ubound(EG1_export_group,1))
	For LngRow = 0 To ubound(EG1_export_group,1)
		
		If  LngRow >= C_SHEETMAXROWS_D  Then			
			StrNextKey = ConvSPChars(EG1_export_group(LngRow,M001_EG1_E1_pur_org))
			Exit For
		End If
		
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(LngRow,M001_EG1_E1_pur_org))
        strData = strData & Chr(11) & ConvSPChars(EG1_export_group(LngRow,M001_EG1_E1_pur_org_nm))
		if Trim(UCase(EG1_export_group(LngRow,M001_EG1_E1_usage_flg))) = "Y" then
			strData = strData & Chr(11) & "Yes"
		Else
			strData = strData & Chr(11) & "No"
		End if 

		strDt = uniConvDate(EG1_export_group(LngRow,M001_EG1_E1_valid_fr_dt))
		If strDt <> uniConvDate(strDefFrDate) Then
			strData = strData & Chr(11) & UNIDateClientFormat(EG1_export_group(LngRow,M001_EG1_E1_valid_fr_dt))
		else
			strData = strData & Chr(11) & ""
		End If 

		strDt = uniConvDate(EG1_export_group(LngRow,M001_EG1_E1_valid_to_dt))
		
		If strDt <> uniConvDate(strDefToDate) Then
			strData = strData & Chr(11) & UNIDateClientFormat(EG1_export_group(LngRow,M001_EG1_E1_valid_to_dt))
		else
			strData = strData & Chr(11) & ""
		End If
        strData = strData & Chr(11) & LngMaxRow + LngRow+1
        strData = strData & Chr(11) & Chr(12)
		
		PvArr(LngRow) = strData
		strData=""
    Next 
	
	strData = join(PvArr,"")
	
	Response.Write "<Script Language=vbscript>"														& vbCr
	Response.Write "With parent"																	& vbCr
	Response.Write "	.ggoSpread.Source = .frm1.vspdData"											& vbCr
	Response.Write "	.ggoSpread.SSShowData """ & strData & """"									& vbCr
	Response.Write "	.lgStrPrevKey = """ & StrNextKey & """"										& vbCr
    Response.Write "    If .frm1.vspdData.MaxRows < .C_SHEETMAXROWS And .lgStrPrevKey <> """" Then" & vbCr
	Response.Write "		.DbQuery"																& vbCr
	Response.Write "	End if" 																	& vbCr
    Response.Write "		.DbQueryOk"																& vbCr
    Response.Write "End With " 																		& vbCr
    Response.Write "</Script>" 																		& vbCr

End Sub    

%>

