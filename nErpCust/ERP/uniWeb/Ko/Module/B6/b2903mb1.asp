<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : �λ�/�޿� 
'*  2. Function Name        : ������ ��ȸ 
'*  3. Program ID           : b2903mb1
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : �̼��� 
'* 10. Modifier (Last)      : �̼��� 
'* 11. Comment              : Ʈ������ �̺�Ʈ�� ó���Ѵ� 
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/03/22 : ..........
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

    Dim ADOConn
    Dim ADORs
    Dim StrSql
    Dim EmpList
    Dim DeptList
	Dim CoName
	Dim Current  '���� ���� ���� ���̵� 
	Dim iRow 
    
    Call LoadBasisGlobalInf()
    
    Call SubOpenDB(ADOConn)                                                        '��: Make  a DB Connection
		select case UCase(request("fnc"))
			case "TREE"
				call tree()
			case "EMP"
				call emp()
		end select

    Call SubCloseDB(ADOConn)                                                       '��: Colse a DB Connection													
    
    
    
    sub tree()

		'****************************************************************************************************
		'	 �����������̰� �����μ�Ȯ���� ���� ���� ���� OLD_internal_cd�� �����;� ��(2005-11-22 JYK)
		'****************************************************************************************************
		'���θ���(�ֻ����μ�)�� ������ �μ����� ���� 
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
															'��: Release RecordSSet
		DeptList = ConvSPChars(DeptList)
		
'		Response.Write DeptList
'		Response.End 
	    
		'���θ� ���� 
		strSql =  "SELECT rTrim(INTERNAL_CD) INTERNAL_CD, rTrim(DEPT_FULL_NM) DEPT_FULL_NM "
		strSql =  strSql & " FROM B_ACCT_DEPT, B_COMPANY "
		strSql =  strSql & " WHERE (PAR_DEPT_CD is null OR rTrim(PAR_DEPT_CD) = '') AND ORG_CHANGE_ID = CUR_ORG_CHANGE_ID "
	    
		if FncOpenRs("R",ADOConn,ADORs,strSql,"X","X") = false then
			CoName = ""
		Else
			CoName = ADORs("DEPT_FULL_NM")                    ' ���θ� 
		end if
		CoName = ConvSPChars(Coname)
		Call SubCloseRs(ADORs)                                                          '��: Release RecordSSet
		Call SubCloseDB(ADOConn)                                                       '��: Colse a DB Connection

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
		Call SubCloseRs(ADORs)                                                          '��: Release RecordSSet	    
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

