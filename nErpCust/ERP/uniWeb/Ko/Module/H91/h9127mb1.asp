<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->


<%
 	Const C_SHEETMAXROWS_D = 100
	Dim lgStrPrevKey  
    
    Call LoadBasisGlobalInf()
    Call LoadinfTb19029B("Q", "H","NOCOOKIE","MB")                                                                   '☜: Clear Error status
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgCurrentSpd      = Request("lgCurrentSpd")                                      '☜: "A"(Spread #1) "B"(Spread #2)

	lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                      

 
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
 
    'call svrmsgbox("mb...into" & lgCurrentSpd , vbinformation,i_mkscript) 

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()        
	Dim strWhere    
	Dim pComp
    Dim iDx
    Dim DFnm        
    Dim li_biz_own_rgst_no
    Dim Oldres_no,Cwork_no
    Dim i,strDNO
    Dim c_per_sub, c_spouse_sub, c_fam_sub, c_old_sub, c_paria_sub, c_lady_sub, c_chl_rear_sub
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim AiDx
    	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
  	        
	Call SubMakeSQLStatements("MR","","x",pComp)                              '☆: Make sql statements    
	Call SubBizQueryMulti()    
 
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""    
        lgStrPrevKey1 = ""    

    Else    

        lgstrData = ""
        Oldres_no = ""
        Cwork_no = 0

        li_biz_own_rgst_no = Trim(lgKeyStream(4))        

        iDx = 1
        AiDx = 1       '일련번호 
        Do While Not lgObjRs.EOF
 
            If Trim(li_biz_own_rgst_no) = "" Or Trim(li_biz_own_rgst_no) <> Trim(replace(ConvSPChars(lgObjRs("biz_own_rgst_no")),"-","")) Then 
                li_biz_own_rgst_no = Trim(replace(ConvSPChars(lgObjRs("biz_own_rgst_no")),"-",""))
                li_biz_own_rgst_no = Left(li_biz_own_rgst_no,7) & "." & Right(li_biz_own_rgst_no,3)
            End If

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("record_type"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("data_type"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tax"))
            lgstrData = lgstrData & Chr(11) & AiDx
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("present_dt"))
            lgstrData = lgstrData & Chr(11) & replace(ConvSPChars(lgObjRs("biz_own_rgst_no")),"-","")
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("hometax_id"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tax_cd"))
            lgstrData = lgstrData & Chr(11) & replace(ConvSPChars(lgObjRs("biz_own_rgst_no2")),"-","")
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("biz_area_nm"))
 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("res_no"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("nat_flag"))	 '2005
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("name"))
            
            lgstrData = lgstrData & Chr(11) & replace(ConvSPChars(lgObjRs("contr_rgst_no")),"-","")
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("contr_name"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("contr_code"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("cnt"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("contr_amt"))         
            lgstrData = lgstrData & Chr(11) & ""

            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext
			AiDx = AiDx + 1
			
            iDx =  iDx + 1   
'            If iDx > C_SHEETMAXROWS_D Then
'				lgStrPrevKey = lgStrPrevKey + 1
'               Exit Do
'            End If                       
                       
        Loop         
        If Trim(lgCurrentSpd) = "A" then
            DFnm = "C:\H" & li_biz_own_rgst_no       
%>
<SCRIPT LANGUAGE=VBSCRIPT>
    parent.frm1.txtFile.value = "<%=DFnm%>"
</SCRIPT>
<%      End If
    End If   	
 
    If iDx <= C_SHEETMAXROWS_D Then
 		  lgStrPrevKey = ""
    End If   
    
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)             
End Sub
 
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case Mid(pDataType,2,1)
        Case "R"
 
			iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1  
					              
            lgStrSQL = " SELECT 'A'  record_type, '27'  data_type,"							' 레코드구분(A)/ 자료구분(26)
            lgStrSQL = lgStrSQL & " tax_biz_cd  tax,"										' 세무서 
			lgStrSQL = lgStrSQL & FilterVar(Trim(replace(lgKeyStream(2),"-","")), "''", "S") & "  present_dt,"               ' 제출연월일 -> 입력변수                    
            lgStrSQL = lgStrSQL & FilterVar(Trim(replace(lgKeyStream(1),"-","")), "", "S") & "  biz_own_rgst_no,"
            lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(3)), "''", "S") & "  hometax_id,"
            lgStrSQL = lgStrSQL & "'9000' tax_cd,"
            lgStrSQL = lgStrSQL & " f.own_rgst_no  biz_own_rgst_no2,"                                        ' 원천징수의무자사업자등록번호 
            lgStrSQL = lgStrSQL & " CONVERT(VARCHAR(40), year_area_nm)  biz_area_nm,"                      ' 법인명(상호)
            lgStrSQL = lgStrSQL & " A.res_no , case when A.nat_cd ='kr' then '1' else '9' end  nat_flag, A.name ,"
            lgStrSQL = lgStrSQL & " isnull(B.contr_rgst_no,own_rgst_no) contr_rgst_no, isnull(B.contr_name,year_area_nm) contr_name,B.contr_code, B.contr_amt, B.cnt "
            lgStrSQL = lgStrSQL & " FROM (	SELECT year_yy,emp_no,contr_rgst_no,contr_name,contr_code,sum(contr_amt) contr_amt,sum(PROV_CNT)  cnt"
			lgStrSQL = lgStrSQL & "			FROM hfa140t "
			lgStrSQL = lgStrSQL & "			WHERE year_flag = 'Y' "            
			lgStrSQL = lgStrSQL & "				AND year_yy = " & FilterVar(left(lgKeyStream(5),4),"''", "S")
			lgStrSQL = lgStrSQL & "				AND contr_dt between " & FilterVar(left(lgKeyStream(5),4) & "0101","''", "S")  
			lgStrSQL = lgStrSQL & "						AND " &  FilterVar(lgKeyStream(5),"''", "S")  											
			lgStrSQL = lgStrSQL & "			GROUP BY  year_yy,emp_no, contr_rgst_no , contr_name ,contr_code "			
            lgStrSQL = lgStrSQL & "		   UNION ALL "
            lgStrSQL = lgStrSQL & "			SELECT yy,emp_no, NULL contr_rgst_no, NULL contr_name,'40' contr_code, priv_contr contr_amt,1 cnt "
            lgStrSQL = lgStrSQL & "			FROM hfa030t "
            lgStrSQL = lgStrSQL & "			WHERE yy = " & FilterVar(left(lgKeyStream(5),4),"''", "S")
            lgStrSQL = lgStrSQL & "				and priv_contr >0 "
            lgStrSQL = lgStrSQL & "		) B  left outer JOIN hfa050t C ON  B.emp_no = C.emp_no  and B.year_yy = C.year_yy"
            lgStrSQL = lgStrSQL & "			JOIN haa010t A ON  B.emp_no = A.emp_no "
            lgStrSQL = lgStrSQL & "			JOIN hfa100t F ON  A.year_area_cd = F.year_area_cd "
            
            lgStrSQL = lgStrSQL & " WHERE f.year_area_cd Like"  & FilterVar(lgKeyStream(6), "''", "S")
			lgStrSQL = lgStrSQL & "		AND A.entr_dt < " & FilterVar(lgKeyStream(5),"''", "S") 
			lgStrSQL = lgStrSQL & "		AND C.contr_sub >=  " & lgKeyStream(7)
			lgStrSQL = lgStrSQL & " ORDER BY B.emp_no,B.contr_rgst_no, B.contr_code "

'Response.Write  lgStrSQL
'Response.End 

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
    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status

    Select Case pOpCode
        Case "MR"
    End Select
End Sub
%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                Select Case Trim("<%=lgCurrentSpd%>")
                    Case "A"
                        .ggoSpread.Source     = .frm1.vspdData
                        .lgStrPrevKey    = "<%=lgStrPrevKey%>"
						if .topleftOK then
							.DBQueryOk
						end if
                      
                End Select
                .ggoSpread.SSShowData "<%=lgstrData%>"
	         End with
          End If   
 
    End Select     
       
</Script>	
