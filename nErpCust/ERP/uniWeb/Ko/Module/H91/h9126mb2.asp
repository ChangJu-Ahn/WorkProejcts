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
    Call LoadBasisGlobalInf()
    Call LoadinfTb19029B("Q", "H","NOCOOKIE","MB")                                                                   '☜: Clear Error status
	
    Dim AlgObjRs 
    Dim BiDx,CiDx
    Dim strFilePath,strMode
    Dim Fnm,CFnm,FPnm      
    Dim Fso,DFnm,CTFnm    
    Dim ARowData, ARowData2, ARowData3
    Dim AQuery
  
    Call HideStatusWnd                                                              '☜: Hide Processing message
   
    BiDx = 1

    lgErrorStatus     = "NO"													    
    lgCurrentSpd      = Request("lgCurrentSpd")                                     '☜: "M"(Spread #1) "S"(Spread #2)
    
    strMode      = Request("txtMode")                                               '☜: Read Operation Mode (CRUD)
    lgKeyStream  = Split(Request("txtKeyStream"),gColSep)
    
    lgstrData = ""
 
    Call SubOpenDB(lgObjConn)      
            
    Select Case strMode
	    Case CStr(UID_M0001)                                                            'data select and create File on server     	        
            Set Fso = CreateObject("Scripting.FileSystemObject")  
                Fnm = Fso.GetFileName(Trim(lgKeyStream(4)))                
    
                FPnm = Server.MapPath("../../files/u2000/" & Fnm)  '2002.02.01 /files 에는 현재 u2000만 존재:나중에 공통쪽 변경되면 수정해야 함.

 
                Call SubBizQuery("")

            If UCase(Trim(lgErrorStatus)) <> "YES" Then
                Set CTFnm = Fso.CreateTextFile (FPnm,True)                              'text를 저장할 File을 생성            
                
                CTFnm.Write lgstrData                                                   'Text 내용부분                       
                DFnm = Fso.GetFileName(FPnm)            
                CTFnm.close    
                Set CTFnm = nothing
            Else
                Call DisplayMsgBox("800004", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
                Call SetErrorStatus() 
            End If
            Set Fso = nothing           
            
            Call HideStatusWnd           
            
%>
    <SCRIPT LANGUAGE=VBSCRIPT>
				parent.subVatDiskOK("<%=DFnm%>")
	</SCRIPT>
<%
'-----------------------------------------------------------------------------------------------
'            File DownLoad(With B.A)
'-----------------------------------------------------------------------------------------------
	
    Case CStr(UID_M0002)

	    Err.Clear 

	    Call HideStatusWnd

	    strFilePath = "http://" & Request.ServerVariables("LOCAL_ADDR") & ":" _
	    			   & Request.ServerVariables("SERVER_PORT")
        If Instr(1, Request.ServerVariables("URL"), "Module") <> 0 Then
            strFilePath = strFilePath & Mid(Request.ServerVariables("URL"), 1, InStr(1, Request.ServerVariables("URL"), "Module") - 1)     
        End If
	    strFilePath = strFilePath  & "files/u2000/"    '2002.02.01 /files 에는 현재 u2000만 존재:나중에 공통쪽 변경되면 수정해야 함.
	    strFilePath = strFilePath & Request("txtFileName")

	End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery(AQuery)   
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim lgstrData2 , AiDx     
	Dim strWhere    
	Dim pComp
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
             
    Call SubMakeSQLStatements("MR",strWhere,"x",pComp)                              '☆: Make sql statements       

    If FncOpenRs("R",lgObjConn,AlgObjRs,lgStrSQL,"X","X") = False Then
       Call SetErrorStatus("")
    Else        
        AiDx = 1    
        Do While Not AlgObjRs.EOF
        
            lgstrData = lgstrData & SetFixSrting(ConvSPChars(AlgObjRs("record_type")),"","",1,"")
            lgstrData = lgstrData & SetFixSrting(ConvSPChars(AlgObjRs("data_type")),"","0",2,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ConvSPChars(AlgObjRs("tax")),"","",3,"")

            lgstrData = lgstrData & SetFixSrting(AiDx,"","0",6,"RIGHT")
 
            lgstrData = lgstrData & SetFixSrting(ConvSPChars(AlgObjRs("present_dt")),"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(replace(ConvSPChars(AlgObjRs("biz_own_rgst_no")),"-",""),"","",10,"")
            lgstrData = lgstrData & SetFixSrting(ConvSPChars(AlgObjRs("hometax_id")),"","",20,"")
            lgstrData = lgstrData & SetFixSrting(ConvSPChars(AlgObjRs("tax_cd")),"","",4,"")
            lgstrData = lgstrData & SetFixSrting(replace(ConvSPChars(AlgObjRs("biz_own_rgst_no2")),"-",""),"","",10,"")
            lgstrData = lgstrData & SetFixSrting(ConvSPChars(AlgObjRs("biz_area_nm")),"","",40,"")

            lgstrData = lgstrData & SetFixSrting(ConvSPChars(AlgObjRs("res_no")),"","",13,"")
            lgstrData = lgstrData & SetFixSrting(ConvSPChars(AlgObjRs("nat_flag")),"","",1,"")
            lgstrData = lgstrData & SetFixSrting(ConvSPChars(AlgObjRs("name")),"","",30,"")
            lgstrData = lgstrData & SetFixSrting(replace(ConvSPChars(AlgObjRs("med_rgst_no")),"-",""),"","",10,"")
            lgstrData = lgstrData & SetFixSrting(ConvSPChars(AlgObjRs("med_name")),"","",40,"")
            lgstrData = lgstrData & SetFixSrting(ConvSPChars(AlgObjRs("cnt")),"","0",5,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(ConvSPChars(AlgObjRs("tot")),"","0",11,"RIGHT")
            
            lgstrData = lgstrData & SetFixSrting(0,"","0",5,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(0,"","0",11,"RIGHT")
            
            lgstrData = lgstrData & SetFixSrting(ConvSPChars(AlgObjRs("family_rel")),"","",1,"")
            lgstrData = lgstrData & SetFixSrting(ConvSPChars(AlgObjRs("family_res_no")),"","",13,"")
            lgstrData = lgstrData & SetFixSrting(ConvSPChars(AlgObjRs("nat_flag")),"","",1,"")
            lgstrData = lgstrData & SetFixSrting(ConvSPChars(AlgObjRs("family_type")),"","",1,"")
            lgstrData = lgstrData & "    "
            lgstrData = lgstrData & Chr(13) & Chr(10)
 
            AlgObjRs.MoveNext

   			AiDx = AiDx + 1
        Loop
    End If
    Call SubHandleError("MR",lgObjConn,AlgObjRs,Err)
    Call SubCloseRs(AlgObjRs)    
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
            lgStrSQL = " SELECT 'A' record_type,'26' data_type, tax_biz_cd  tax,1  no,"       ' 레코드구분(A)/자료구분(26)/세무서 
            lgStrSQL = lgStrSQL & " 1  no,"
			lgStrSQL = lgStrSQL & FilterVar(Trim(replace(lgKeyStream(2),"-","")), "''", "S") & "  present_dt,"                 ' 제출연월일 -> 입력변수                    
            lgStrSQL = lgStrSQL & FilterVar(Trim(replace(lgKeyStream(1),"-","")), "", "S") & "  biz_own_rgst_no,"
            lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(3)), "''", "S") & "  hometax_id,"
            lgStrSQL = lgStrSQL & FilterVar("9000", "''", "S") & " tax_cd,"
            lgStrSQL = lgStrSQL & " f.own_rgst_no  biz_own_rgst_no2,"                                        ' 원천징수의무자사업자등록번호 
            lgStrSQL = lgStrSQL & " CONVERT(VARCHAR(40), year_area_nm)  biz_area_nm,"                      ' 법인명(상호)
            
            lgStrSQL = lgStrSQL & " A.res_no as res_no,"
            lgStrSQL = lgStrSQL & " case when A.nat_cd ='kr' then '1' else '9' end  nat_flag,"            
            lgStrSQL = lgStrSQL & " A.name as name,"
            lgStrSQL = lgStrSQL & " B.med_rgst_no as med_rgst_no,"
            lgStrSQL = lgStrSQL & " B.med_name as med_name,"
            lgStrSQL = lgStrSQL & " B.family_rel as family_rel," 
            lgStrSQL = lgStrSQL & " sum(B.med_amt) as tot ,"
            lgStrSQL = lgStrSQL & " sum(B.PROV_CNT) as cnt ,"
            lgStrSQL = lgStrSQL & " B.family_res_no as family_res_no,"
            lgStrSQL = lgStrSQL & " case when B.family_type ='A' or B.family_type ='B' or B.family_rel='0' then '1' else '2' end  family_type"

            lgStrSQL = lgStrSQL & " FROM hfa100t f, haa010t A , hfa130t B , hfa050t C " 
            lgStrSQL = lgStrSQL & " WHERE f.year_area_cd Like"  & FilterVar(lgKeyStream(6), "''", "S")
 
			lgStrSQL = lgStrSQL & " AND f.year_area_cd = a.year_area_cd "	
			lgStrSQL = lgStrSQL & " AND A.emp_no = B.emp_no "
			lgStrSQL = lgStrSQL & " AND A.emp_no = C.emp_no "
			lgStrSQL = lgStrSQL & " AND B.year_yy = C.year_yy "
			lgStrSQL = lgStrSQL & " AND B.year_flag = 'Y' "            
			lgStrSQL = lgStrSQL & " AND B.year_yy = " & FilterVar(Year(UNIConvDateCompanyToDB(lgKeyStream(5),NULL)),"NULL", "S")
			lgStrSQL = lgStrSQL & " AND B.med_dt between " & " CONVERT(VARCHAR(4) , DATEPART(year," & replace(FilterVar(UNIConvDateCompanyToDB(lgKeyStream(5),NULL),"NULL", "S"),gComDateType,"") & ")) + '0101'"
			lgStrSQL = lgStrSQL & " AND " & replace(FilterVar(UNIConvDateCompanyToDB(lgKeyStream(5),NULL),"NULL", "S"),gComDateType,"")
			lgStrSQL = lgStrSQL & " AND A.entr_dt < " & replace(FilterVar(UNIConvDateCompanyToDB(lgKeyStream(5),NULL),"NULL", "S"),gComDateType,"")
			lgStrSQL = lgStrSQL & " AND C.med_sub >=  " & FilterVar(lgKeyStream(7), "''", "S")
			lgStrSQL = lgStrSQL & " GROUP BY  f.TAX_BIZ_CD,f.OWN_RGST_NO,f.YEAR_AREA_NM, B.emp_no , A.res_no, A.name , A.nat_cd , B.med_rgst_no , B.med_name ,B.family_rel , B.family_res_no , B.family_type "
			lgStrSQL = lgStrSQL & " ORDER BY B.emp_no,B.med_rgst_no, B.family_rel, B.med_name "
'Response.Write lgStrSQL
'Response.end		
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
'============================================================================================================
' Name : SetFixSrting(입력값,비교문자,대체문자,고정길이,문자정렬방향)
' Desc : This Function return srting
'============================================================================================================
' SetFixSrting(replace(ConvSPChars(BlgObjRs("med_rgst_no")),"-",""),"","",10,"")

Function SetFixSrting(InValue, ComSymbol, strFix, InPos, direct)
    Dim Cnt,MCnt
    Dim ExSymbol,strSplit,strMid
    Dim iDx,i,strTemp
    
    If InValue = "" OR IsNull(InValue) Then                                         '입력값이 존재하지않을경우 입력값의 길이를 0으로 한다.
        Cnt = 0     
    Else
        
        Cnt = Len(InValue)              
        For i = 1 To Cnt
            strMid = Mid(InValue,i,1)
            If Asc(strMid) > 255 OR Asc(strMid) < 0 Then
                If Trim(strMid) = ")" Or Trim(strMid) = "(" Then
                    MCnt = MCnt + 1
                Else
                    MCnt = MCnt + 2                                                  '한글부분만 길이를 각각 2로한다.
                End If                                                 '한글부분만 길이를 각각 2로한다.
            Else
                MCnt = MCnt + 1
            End If
        Next
        Cnt = MCnt
                 
        If ComSymbol = "" OR IsNull(ComSymbol) Then                                  '비교문자가 없을경우 
        Else                                                                         '비교문자가 존재할경우 비교문자를 뺀 나머지를 입력값으로한다.
            ExSymbol = Split(InValue,ComSymbol)
            If UBound(ExSymbol) > 0 Then
                iDx = UBound(ExSymbol)
                For i = 0 To iDx
                    strSplit = strSplit & ExSymbol(i)
                Next
                InValue = strSplit
            End If               
        End If        
    End If        
    
    If InPos = "" Then                                                              '고정길이가 정해지지 않을 경우 입력문자 길이가 고정길이가 된다.
        InPos = Cnt  
    End If
    
    If UCase(Trim(direct)) = "LEFT" OR UCase(Trim(direct)) = "" Then                '왼쪽정렬일경우(defalut) 고정길이 보다 작은 길이의 문자가 입력되면 나머지 공백(default)부분을 대체문자로 체운다.
        If InPos > Cnt Then                                                         ' ex:hi    
            If strFix = "" Then
               strFix = Chr(32)
            End if 
            For i = (Cnt+1) To InPos        
                InValue = InValue & strFix
            Next         

		ElseIf InPos < Cnt Then '입력문자가 지정길이를 초과할경우 뒷자리는 잘라버림. cyc
			InValue = Left(InValue , InPos)

        End If
        
    ElseIf UCase(Trim(direct)) = "RIGHT" Then                                       '오른쪽정렬 
        If InPos > Cnt Then                                                         ' ex:     hi
            If strFix = "" Then
               strFix = Chr(32)
            End if 
            For i = 1 To (InPos - Cnt)
                strTemp = strTemp & strFix         
            Next
            InValue = strTemp & InValue
        End If
    End If
    SetFixSrting = InValue 
End Function

%>

<script language="vbscript">
		Dim SF
		On Error Resume Next
		Set SF = CreateObject("uni2kCM.SaveFile")
		Call SF.SaveTextFile("<%= strFilePath %>")

		Set SF = Nothing
</script>
