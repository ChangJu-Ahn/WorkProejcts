<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : 인사/급여 
'*  2. Function Name        : 조직도 조회 
'*  3. Program ID           : b2903mb1
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 이석민 
'* 10. Modifier (Last)      : 이석민 
'* 11. Comment              : 트리뷰의 이벤트를 처리한다 
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/03/22 : ..........
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

    Dim ADOConn
    Dim ADORs
    Dim StrSql
    Dim EmpList
    Dim DeptList
	Dim CoName
	Dim Current  '현재 조직 개편 아이디 
	Dim iRow 
    
    Call LoadBasisGlobalInf()
    
    Call SubOpenDB(ADOConn)                                                        '☜: Make  a DB Connection
		select case UCase(request("fnc"))
			case "TREE"
				call tree()
			case "EMP"
				call emp()
		end select

    Call SubCloseDB(ADOConn)                                                       '☜: Colse a DB Connection													
    
    
    
    sub tree()

		'****************************************************************************************************
		'	 조직개편중이고 최종부서확정이 되지 않은 경우는 OLD_internal_cd를 가져와야 함(2005-11-22 JYK)
		'****************************************************************************************************
		'법인명을(최상위부서)를 제외한 부서정보 쿼리 
		strSql = "SELECT rTrim(PAR_DEPT_CD) PAR_DEPT_CD, rTrim(DEPT_CD) DEPT_CD, "
		strSql = strSql & " rtrim(CASE WHEN (SELECT WORK_FLAG FROM HORG_WORK_LIST WHERE ORG_CHANGE_ID = " 
		strSql = strSql & " (SELECT ORGID FROM HORG_ABS WHERE ORGDT=(SELECT MAX(ORGDT) FROM HORG_ABS))) IN ('A','B') "
		strSql = strSql & " THEN OLD_INTERNAL_CD ELSE INTERNAL_CD END) INTERNAL_CD, "
		strSql = strSql & " rTrim(DEPT_FULL_NM) DEPT_FULL_NM "
		strSql = strSql & " FROM  B_ACCT_DEPT, B_COMPANY "
		strSql = strSql & " WHERE rTrim(PAR_DEPT_CD) <> '' AND ORG_CHANGE_ID = CUR_ORG_CHANGE_ID "
		strSql = strSql & " ORDER BY INTERNAL_CD, PAR_DEPT_CD, DEPT_CD"

		If 	FncOpenRs("R",ADOConn,ADORs,strSql,"X","X") = False Then                    'If data not exists
			DeptList =  ""
		Else
			While Not ADORs.EOF        
				DeptList = DeptList & ADORs("PAR_DEPT_CD") & chr(11) & ADORs("DEPT_CD") & chr(11) & ADORs("INTERNAL_CD")
				DeptList = DeptList & chr(11) & ADORs("DEPT_FULL_NM") & chr(11) & Chr(12)
				   
				ADORs.MoveNext
			Wend        
		End If
		Call SubCloseRs(ADORs)   
															'☜: Release RecordSSet
		DeptList = ConvSPChars(DeptList)
		
'		Response.Write DeptList
'		Response.End 
	    
		'법인명 쿼리 
		strSql =  "SELECT rTrim(INTERNAL_CD) INTERNAL_CD, rTrim(DEPT_FULL_NM) DEPT_FULL_NM "
		strSql =  strSql & " FROM B_ACCT_DEPT, B_COMPANY "
		strSql =  strSql & " WHERE (PAR_DEPT_CD is null OR rTrim(PAR_DEPT_CD) = '') AND ORG_CHANGE_ID = CUR_ORG_CHANGE_ID "
	    
		if FncOpenRs("R",ADOConn,ADORs,strSql,"X","X") = false then
			CoName = ""
		Else
			CoName = ADORs("DEPT_FULL_NM")                    ' 법인명 
		end if
		CoName = ConvSPChars(Coname)
		Call SubCloseRs(ADORs)                                                          '☜: Release RecordSSet
		Call SubCloseDB(ADOConn)                                                       '☜: Colse a DB Connection

    end sub
    
    
	sub emp()  
		strSql = "SELECT NAME, EMP_NO, dbo.ufn_GetCodeName(" & FilterVar("H0026", "''", "S") & " , ROLE_CD) ROLE_NM, dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & " ,PAY_GRD1) PAY_NM, TEL_NO "
		strSql = strSql & " FROM HAA010T "
		strSql = strSql & " WHERE (" & FilterVar("ID", "''", "S") & " + rTrim(INTERNAL_CD) =  " & FilterVar(Request("nodeKey"), "''", "S") & " ) "
		strSql = strSql &   " AND (RETIRE_DT IS NULL OR RETIRE_RESN = " & FilterVar("6", "''", "S") & " ) "
	'    strSql = strSql &       " AND (ROLE_CD  = ROLE.MINOR_CD AND ROLE.MAJOR_CD = 'H0026') "
	'   strSql = strSql &       " AND (PAY_GRD1 = PAY.MINOR_CD  AND PAY.MAJOR_CD  = 'H0001') "
		strSql = strSql & " ORDER BY ROLE_CD "
	    
	    iRow = 0   
		If 	FncOpenRs("R",ADOConn,ADORs,strSql,"X","X") = False Then                    'If data not exists
			EmpList =  ""
		Else
			While Not ADORs.EOF        
			    iRow = iRow + 1
				EmpList = EmpList & "" & chr(11) & ADORs("NAME") & chr(11) & ADORs("EMP_NO") & chr(11) & ADORs("ROLE_NM") 
				EmpList = EmpList & chr(11) & ADORs("PAY_NM") & chr(11) & ADORs("TEL_NO") & chr(11) & iRow & chr(11) & Chr(12)
				ADORs.MoveNext
			WEnd        
		End If
		Call SubCloseRs(ADORs)                                                          '☜: Release RecordSSet	    
    end Sub
    

													
%>

<!-- #Include file="../../inc/uni2KCM.inc" -->	

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Operation.vbs"> </SCRIPT>	 

<Script Language="VBScript">
	
	select case "<%=UCase(request("fnc"))%>"
		case "EMP"
			parent.frm1.vspdData.maxrows = 0
			parent.ggoSpread.source = parent.frm1.vspdData
			parent.ggoSpread.SSShowData "<%=ConvSPChars(EmpList)%>"
		case "TREE"
			parent.frm1.DeptList.value = "<%=DeptList%>"
			parent.frm1.CoName.value = "<%=CoName%>"
			call parent.DbqueryOk()
	end select
	
</Script>

