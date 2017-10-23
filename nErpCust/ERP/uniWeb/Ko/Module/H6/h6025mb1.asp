<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%
	Dim lgStrPrevKey
   	Const C_SHEETMAXROWS_D = 100

    Call HideStatusWnd                                                               '☜: Hide Processing message 
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB")
                                                                   

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)

	lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection
'=============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    Dim iDx
	Dim strWhere

    dim tempArr(100,14)
    dim old_dept_nm
    dim row_cnt,col_cnt
	dim sum
	dim i,j

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strWhere = FilterVar(lgKeyStream(0), "''", "S")  'frm1.pau_yymm
    strWhere = strWhere & " AND a.emp_no LIKE " & FilterVar(Trim(lgKeyStream(1)),"" & FilterVar("%", "''", "S") & "","S")   
    strWhere = strWhere & " AND b.prov_type LIKE " & FilterVar(Trim(lgKeyStream(3)),"" & FilterVar("%", "''", "S") & "","S")    
    strWhere = strWhere & " AND a.pay_grd1 LIKE " & FilterVar(Trim(lgKeyStream(2)),"" & FilterVar("%", "''", "S") & "","S")   
    
    strWhere = strWhere & " And dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(7),"'%'", "S") & ") >= " & FilterVar(lgKeyStream(4), "''", "S")
    strWhere = strWhere & " And dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(7),"'%'", "S") & ") <= " & FilterVar(lgKeyStream(5), "''", "S")    
    strWhere = strWhere & " AND dbo.ufn_H_get_internal_cd(a.emp_no," & FilterVar(lgKeyStream(7),"'%'", "S") & ") LIKE  " & FilterVar(lgKeyStream(6) & "%", "''", "S") & ""    '  자료권한 Internal_cd = '?'    

    'strWhere = strWhere & " AND a.internal_cd >=  " & FilterVar(lgKeyStream(4), "''", "S") & ""       '  internal_cd = min
    'strWhere = strWhere & " AND a.internal_cd <=  " & FilterVar(lgKeyStream(5), "''", "S") & ""       '  internal_cd = max
    'strWhere = strWhere & " AND a.internal_cd LIKE  " & FilterVar(lgKeyStream(6) & "%", "''", "S") & ""    '  자료권한 Internal_cd = '?'
    strWhere = strWhere & " and a.emp_no = b.emp_no "
    strWhere = strWhere & " and b.pay_cd in (" & FilterVar("0", "''", "S") & " ," & FilterVar("1", "''", "S") & " ," & FilterVar("2", "''", "S") & "," & FilterVar("3", "''", "S") & ")"
	Call SubMakeSQLStatements("MR",strWhere,"1")                                   '☆ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call Displaymsgbox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜ : No data is found.
        Call SetErrorStatus()
    Else
        iDx       = 1

        old_dept_nm = ""
		col_cnt = 0
		row_cnt = 0

		for i=1 to C_SHEETMAXROWS_D
			for j=0 to 13
				tempArr(i,j) =0
			next
		next        

        Do While Not lgObjRs.EOF
		
			if lgObjRs("dept_nm") <> old_dept_nm then
				row_cnt = row_cnt+1				
				tempArr(row_cnt,0) = lgObjRs("DEPT_NM")
			end if
				
			col_cnt = ConvSPChars(lgObjRs("PAY_CD")) *2 +1

			tempArr(row_cnt,col_cnt) = ConvSPChars(lgObjRs("COUNT"))
			tempArr(row_cnt,col_cnt+1) = ConvSPChars(lgObjRs("PAY_TOT_AMT"))

			old_dept_nm = lgObjRs("dept_nm")
            
	        lgObjRs.MoveNext

            iDx =  iDx + 1
        Loop 
    End If

	Call SubMakeSQLStatements("MR",strWhere,"2")

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
    Else

        iDx       = 1
		col_cnt = 0

        Do While Not lgObjRs.EOF
			for i=1 to row_cnt

				if ConvSPChars(lgObjRs("dept_nm")) = tempArr(i,0) then
					col_cnt = 8+ConvSPChars(lgObjRs("num"))*2
					tempArr(i,col_cnt) = ConvSPChars(lgObjRs("COUNT"))
					tempArr(i,col_cnt+1) = ConvSPChars(lgObjRs("PAY_TOT_AMT"))
				end if
			Next
	        lgObjRs.MoveNext

            iDx =  iDx + 1
        Loop 
    End If

	lgstrData = ""		
	for i=1 to row_cnt
		for j=0 to 13
				
			if j=0 then
				lgstrData = lgstrData & Chr(11) & ConvSPChars(tempArr(i,j))
			elseif j=9 then
				sum =0+tempArr(i,2)+tempArr(i,4)+tempArr(i,6)+tempArr(i,8)
				lgstrData = lgstrData & Chr(11) & UNINumClientFormat(sum, ggAmtOfMoney.DecPoint,0)
			else
				lgstrData = lgstrData & Chr(11) & UNINumClientFormat(tempArr(i,j), ggAmtOfMoney.DecPoint,0)
			end if
		next

		lgstrData = lgstrData & Chr(11) & lgLngMaxRow + i+1
		lgstrData = lgstrData & Chr(11) & Chr(12)
	next

    Call SubHandleError("MR",lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode1,pCode2)
    Dim iSelCount

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                '☜: Clear Error status
    lgStrSQL = ""

    Select Case Mid(pDataType,1,1)
        Case "M"
        
           Select Case Mid(pDataType,2,1)
               Case "R"
                    Select Case pCode2
                       Case "1"

							lgStrSQL = "Select  DEPT_NM,PAY_CD,isnull(count( EMP_NO),0) COUNT, isnull(sum(PAY_TOT_AMT),0) PAY_TOT_AMT "
							lgStrSQL = lgStrSQL & " from ("                       
							lgStrSQL = lgStrSQL & " Select  DEPT_NM,PAY_CD, EMP_NO , isnull(sum(PAY_TOT_AMT),0) PAY_TOT_AMT "
							lgStrSQL = lgStrSQL & " From  ( Select  a.EMP_NO, b.DEPT_CD, b.DEPT_NM, b.PROV_TYPE, b.pay_cd,"
							lgStrSQL = lgStrSQL & " CASE b.PROV_TYPE WHEN " & FilterVar("1", "''", "S") & "  THEN b.PAY_TOT_AMT WHEN " & FilterVar("P", "''", "S") & "  THEN b.PAY_TOT_AMT ELSE b.BONUS_TOT_AMT END AS PAY_TOT_AMT  "
							lgStrSQL = lgStrSQL & " From HDF020T a ,HDF070T b "
							lgStrSQL = lgStrSQL & " Where  b.PAY_YYMM = " & pCode1
							lgStrSQL = lgStrSQL & " ) temp "
							lgStrSQL = lgStrSQL & " group by DEPT_NM,PAY_CD, EMP_NO ) temp2"
							lgStrSQL = lgStrSQL & " group by dept_nm,pay_cd "
							lgStrSQL = lgStrSQL & " order by dept_nm,pay_cd "

                       Case "2"
							lgStrSQL = "select  1 num,DEPT_NM,isnull(count(emp_no),0) count ,isnull(sum(PAY_TOT_AMT),0) PAY_TOT_AMT "
							lgStrSQL = lgStrSQL & " from ("
							lgStrSQL = lgStrSQL & " select  DEPT_NM,emp_no,isnull(sum(PAY_TOT_AMT),0) PAY_TOT_AMT "							
                            lgStrSQL = lgStrSQL & " From  (Select  a.EMP_NO, b.DEPT_CD, b.DEPT_NM, b.PROV_TYPE, b.pay_cd,"
                            lgStrSQL = lgStrSQL & " CASE b.PROV_TYPE WHEN " & FilterVar("1", "''", "S") & "   THEN b.PAY_TOT_AMT WHEN " & FilterVar("P", "''", "S") & "  THEN b.PAY_TOT_AMT ELSE b.BONUS_TOT_AMT END AS PAY_TOT_AMT "
                            lgStrSQL = lgStrSQL & " From HDF020T a ,HDF070T b  "
			                lgStrSQL = lgStrSQL & " Where  b.PAY_YYMM = " & pCode1
			                lgStrSQL = lgStrSQL & " and except_type in (" & FilterVar("1", "''", "S") & " ," & FilterVar("3", "''", "S") & ") "
			                lgStrSQL = lgStrSQL & " ) t group by dept_nm,emp_no "
			                lgStrSQL = lgStrSQL & " ) t2 group by dept_nm "
			                lgStrSQL = lgStrSQL & " union "
			                lgStrSQL = lgStrSQL & " select 2 num,DEPT_NM,isnull(count(emp_no),0) count ,isnull(sum(PAY_TOT_AMT),0) PAY_TOT_AMT "
							lgStrSQL = lgStrSQL & " from ("			                
			                lgStrSQL = lgStrSQL & " select 2 num,DEPT_NM,emp_no ,isnull(sum(PAY_TOT_AMT),0) PAY_TOT_AMT "			                
			                lgStrSQL = lgStrSQL & " from ( Select  a.EMP_NO, b.DEPT_CD, b.DEPT_NM, b.PROV_TYPE, b.pay_cd,"
			                lgStrSQL = lgStrSQL & " CASE b.PROV_TYPE WHEN " & FilterVar("1", "''", "S") & "   THEN b.PAY_TOT_AMT WHEN " & FilterVar("P", "''", "S") & "  THEN b.PAY_TOT_AMT ELSE b.BONUS_TOT_AMT END AS PAY_TOT_AMT"
			                lgStrSQL = lgStrSQL & " From HDF020T a ,HDF070T b  "
			                lgStrSQL = lgStrSQL & " Where  b.PAY_YYMM = " & pCode1
			                lgStrSQL = lgStrSQL & " and except_type in (" & FilterVar("2", "''", "S") & "," & FilterVar("3", "''", "S") & ") "
			                lgStrSQL = lgStrSQL & " ) s group by dept_nm,emp_no "
							lgStrSQL = lgStrSQL & " ) s2 group by dept_nm "			                
			                lgStrSQL = lgStrSQL & " order by dept_nm,num "
			                
					End Select
           End Select 
    End Select
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub

%>
<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             With Parent

					.ggoSpread.Source     = .frm1.vspdData1
					.ggoSpread.SSShowData "<%=lgstrData%>"
					.lgStrPrevKey    = "<%=lgStrPrevKey%>"
					.DBQueryOk
	         End with
	      Else
             Parent.DBQueryNo
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
    
       
</Script>	

