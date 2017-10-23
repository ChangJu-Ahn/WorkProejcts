<%@ LANGUAGE="VBSCRIPT" %>

<!--
======================================================================================================
*  1. Module Name          : 인사/급여 
*  2. Function Name        : 조직도조회 
*  3. Program ID           : B2903mb2
*  4. Program Name         : 조직도 조회 
*  5. Program Desc         : 조직도를 트리뷰 형태로 보여준다 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001//
*  8. Modified date(Last)  : 2002/12/17
*  9. Modifier (First)     : 이석민 
* 10. Modifier (Last)      : Sim Hae Young
* 11. Comment              :
=======================================================================================================-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->


<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

    Dim ADOConn
    Dim ADORs
    Dim StrSql
    Dim ORGNM
    Dim ORGDT
    Dim DeptList
    Dim CoName
    
    Call HideStatusWnd 
    Call LoadBasisGlobalInf()
 
    Call SubOpenDB(ADOConn)                                                        '☜: Make  a DB Connection

	strSql = "SELECT ORGNM,ORGDT FROM HORG_ABS WHERE ORGID = " & FilterVar(Request("txtOrgId"), "''", "S")  
	If 	FncOpenRs("R",ADOConn,ADORs,strSql,"X","X") = False Then                    'If data not exists
		ORGNM = chr(12)
		Call DisplayMsgBox("900014", vbinformation, "", "", I_MKSCRIPT)   '검색된 데이터가 없습니다 
	Else
		ORGNM = ADORs("ORGNM")
		ORGDT = ADORs("ORGDT")

		Call SubCloseRs(ADORs)																	'☜: Release RecordSSet

		'법인명을(최상위 부서-법인명)를 제외한 부서정보 쿼리 
		strSql = "SELECT rTrim(PAR_DEPT_CD) PAR_DEPT_CD, rTrim(DEPT_CD) DEPT_CD, "
		strSql = strSql & " rtrim(CASE WHEN (SELECT ORGID FROM HORG_ABS WHERE ORGDT=(SELECT MAX(ORGDT) FROM HORG_ABS)) <> " & FilterVar(Request("txtOrgId"),"''","S") & " THEN "
		strSql = strSql & "                 CASE WHEN (SELECT WORK_FLAG FROM HORG_WORK_LIST WHERE ORG_CHANGE_ID = "
		strSql = strSql & "                           (SELECT ORGID FROM HORG_ABS WHERE ORGDT=(SELECT TOP 1 ORGDT FROM HORG_ABS WHERE ORGDT > " & FilterVar(ORGDT,"''","S") & " "
		strSql = strSql & "                                   		                            ORDER BY ORGDT ASC ))) = ''  THEN INTERNAL_CD "
		strSql = strSql & "                 ELSE OLD_INTERNAL_CD END  "
		strSql = strSql & " ELSE INTERNAL_CD END) INTERNAL_CD , "
		strSql = strSql & " rTrim(DEPT_FULL_NM) DEPT_FULL_NM "
		strSql = strSql & " FROM  B_ACCT_DEPT "
		strSql = strSql & " WHERE rTrim(PAR_DEPT_CD) <> '' AND ORG_CHANGE_ID =  " & FilterVar(Request("txtOrgId"),"", "S") & " "
		strSql = strSql & " ORDER BY INTERNAL_CD, PAR_DEPT_CD, DEPT_CD"
 
		If 	FncOpenRs("R",ADOConn,ADORs,strSql,"X","X") = False Then                    'If data not exists
		    DeptList =  chr(12)
		Else
		    While Not ADORs.EOF        
		       DeptList = DeptList & ADORs("PAR_DEPT_CD") & chr(11) & ADORs("DEPT_CD") & chr(11) & ADORs("INTERNAL_CD")
		       DeptList = DeptList & chr(11) & ADORs("DEPT_FULL_NM") & chr(11) & Chr(12)
		       
		       ADORs.MoveNext
		    Wend        
		End If

		Call SubCloseRs(ADORs)                                                          '☜: Release RecordSSet
    
		'법인명 쿼리 
		strSql =  "SELECT rTrim(INTERNAL_CD) INTERNAL_CD, rTrim(DEPT_FULL_NM) DEPT_FULL_NM "
		strSql =  strSql & " FROM B_ACCT_DEPT, B_COMPANY "
		strSql =  strSql & " WHERE (PAR_DEPT_CD is null OR rTrim(PAR_DEPT_CD) = '') AND ORG_CHANGE_ID = " & FilterVar(Request("txtOrgId"),"", "S") & " "
    
		If FncOpenRs("R",ADOConn,ADORs,strSql,"X","X") = False Then
			CoName = ""
		Else
			CoName = ADORs("DEPT_FULL_NM")                    ' 법인명 
		End If
	
		Call SubCloseRs(ADORs)                                                          '☜: Release RecordSSet
	End If

    ORGNM = ConvSPChars(ORGNM)
    DeptList = ConvSPChars(DeptList)
    CoName = ConvSPChars(Coname)
    
    Call SubCloseDB(ADOConn)                                                       '☜: Colse a DB Connection
	
%>

<Script Language=vbscript>	
	
	parent.LayerShowHide(0)
		
	If "<%=ORGNM%>" <> chr(12) And "<%=Deptlist%>" <> chr(12) Then
		
		With parent																		'☜: 화면 처리 ASP 를 지칭함 
			.frm1.txtOrgNm.value = "<%=OrgNM%>"

			call .MakeTree("<%=Deptlist%>", "<%=CoName%>")
		 	
		 	.frm1.btnCb_allCollapse.disabled = False
			.frm1.btnCb_allExpand.disabled   = False
		End With
	Else
	 	parent.frm1.btnCb_allCollapse.disabled = True
		parent.frm1.btnCb_allExpand.disabled   = True
		
	End If
</Script>

