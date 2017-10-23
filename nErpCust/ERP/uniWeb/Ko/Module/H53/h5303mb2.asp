<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%
	Dim lgStrPrevKey, strTab
	Const C_SHEETMAXROWS_D = 100
	strTab			  = Request("txtTab")                                           '��: Read Operation Mode (CRUD)
	
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "MB")

    Call HideStatusWnd                                                               '��: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Dim strFilePath,strMode,Pinfo,iDx
    Dim Fnm,CFnm,FPnm      
    Dim Fso,DFnm,CTFnm   
    
    lgCurrentSpd      = Request("lgCurrentSpd")                                     '��: "M"(Spread #1) "S"(Spread #2)
    lgstrData = ""

    Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection
    Select Case lgOpModeCRUD
	    Case CStr(UID_M0001)                                                            'data select and create File on server     	        
            Set Fso = CreateObject("Scripting.FileSystemObject")
				iDx = 1
                Pinfo = Request.ServerVariables ("PATH_INFO")
   
                
	            Fnm = Mid(Pinfo,InstrRev(Pinfo,"/")+1,InstrRev(Pinfo,".")-InstrRev(Pinfo,"/")-1)    'File�� ����� File Name�� ���� 
				FPnm = Server.MapPath("../../files/u2000/" & Fnm & "_" & iDx)           '��θ� System ���丮�� �ٲ۴�.
 
				Do While Fso.FileExists (Fpnm)                                                      'Server�ʿ� ������ File Name�� �ߺ����� 
           
				    iDx = Mid(FPnm,InstrRev(FPnm,"_")+1)                                            
				    iDx = iDx + 1        
				    FPnm = Server.MapPath("../../files/u2000/" & Fnm & "_" & iDx)       '"_" & ���� �� �ٿ� ȭ���� ��ü ���丮��θ� ����         
           
				Loop  
				         
                Call SubBizQueryMulti()

            If UCase(Trim(lgErrorStatus)) <> "YES" Then
                Set CTFnm = Fso.CreateTextFile (FPnm,True)                              'text�� ������ File�� ����            
              
                CTFnm.Write lgstrData                                                   'Text ����κ�                       
                DFnm = Fso.GetFileName(FPnm)
                CTFnm.close    
                Set CTFnm = nothing
                
            Else
                Call DisplayMsgBox("800004", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
                Call SetErrorStatus() 
            End If
            Set Fso = nothing           

            Call HideStatusWnd           
            
%>
    <SCRIPT LANGUAGE=VBSCRIPT>
    
				parent.subVatDiskOK("<%=DFnm%>")
	</SCRIPT>
<%
	    Case CStr(UID_M0002)
		    Err.Clear 
		    Call HideStatusWnd
		    strFilePath = "http://" & Request.ServerVariables("SERVER_NAME") & ":" _
		    			   & Request.ServerVariables("SERVER_PORT")
	        If Instr(1, Request.ServerVariables("URL"), "Module") <> 0 Then
	            strFilePath = strFilePath & Mid(Request.ServerVariables("URL"), 1, InStr(1, Request.ServerVariables("URL"), "Module") - 1)     
	        End If
		    strFilePath = strFilePath  & "files/u2000/" 
		    strFilePath = strFilePath & Request("txtFileName")
	End Select
    
    Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strSect_cd
    Dim strWhere
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call SubMakeSQLStatements("MR",strWhere,iKey1,C_EQ)                              '�� : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)              '�� : No data is found. 
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & SetFixSrting(lgObjRs("NUM")		,"","0",6,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(lgObjRs("BIZ_AREA_NUM")		,"","0",8,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(lgObjRs("SUB_NUM")		,"","0",1,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(lgObjRs("HOIGEI")		,"","0",2,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(lgObjRs("UNIT_BIZ_AREA")		,"","0",3,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(lgObjRs("CERTI_NUM")		,"","0",11,"")
            lgstrData = lgstrData & SetFixSrting(lgObjRs("NAME"),""," ",16,"LEFT")
            lgstrData = lgstrData & SetFixSrting(replace(lgObjRs("res_no"),"-",""),"","",13,"")
            lgstrData = lgstrData & SetFixSrting(lgObjRs("ACQ_DT"),"","",8,"") 
            lgstrData = lgstrData & SetFixSrting(lgObjRs("LAST_PAY_MONTH"),"","0",2,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(lgObjRs("LAST_PAY_TOT"),"","0",13,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(lgObjRs("LAST_BOSU_TOT"),"","0",15,"RIGHT")
            lgstrData = lgstrData & SetFixSrting(lgObjRs("DUTY_MONTH"),"","0",2,"RIGHT")
            
            lgstrData = lgstrData & Chr(13) & Chr(10)
            lgObjRs.MoveNext

            iDx = iDx + 1
        Loop
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

    Select Case Mid(pDataType,1,1)
        Case "M"
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
                     
               Case "R"
					lgStrSQL = "SELECT NUM, BIZ_AREA_NUM, SUB_NUM, HOIGEI, UNIT_BIZ_AREA, CERTI_NUM, NAME, RES_NO, ACQ_DT, LAST_PAY_MONTH, LAST_PAY_TOT, LAST_BOSU_TOT, DUTY_MONTH "
					lgStrSQL = lgStrSQL & " FROM HDB030T "
					lgStrSQL = lgStrSQL & " WHERE DIV=" & FilterVar(strTab, "''", "S") & " AND YEAR_YY =" & FilterVar(lgKeyStream(0), "''", "S") & " AND BIZ_AREA_CD = " & FilterVar(lgKeyStream(1), "''", "S")
					lgStrSQL = lgStrSQL & " ORDER BY NUM "
'Response.Write "lgStrSQL:" & lgStrSQL
'Response.End                  
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
    lgErrorStatus     = "YES"                                                         '��: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)

    On Error Resume Next                                                              '��: Protect system from crashing
    Err.Clear                                                                         '��: Clear Error status

    Select Case pOpCode
        Case "MR"
 
    End Select
End Sub
'============================================================================================================
' Name : SetFixSrting(�Է°�,�񱳹���,��ü����,��������,�������Ĺ���)
' Desc : This Function return srting
'============================================================================================================
Function SetFixSrting(InValue, ComSymbol, strFix, InPos, direct)
    Dim Cnt,MCnt
    Dim ExSymbol,strSplit,strMid
    Dim iDx,i,strTemp
    
    If InValue = "" OR IsNull(InValue) Then                                         '�Է°��� ��������������� �Է°��� ���̸� 0���� �Ѵ�.
        Cnt = 0     
    Else                                  '�Է°��� �����ϸ鼭 �ѱ��ϰ�� 
        Cnt = Len(InValue)              
        For i = 1 To Cnt
            strMid = Mid(InValue,i,1)
            If Asc(strMid) > 255 OR Asc(strMid) < 0 Then
                MCnt = MCnt + 2                                                  '�ѱۺκи� ���̸� ���� 2���Ѵ�.
            Else
                MCnt = MCnt + 1    
            End If
        Next
        Cnt = MCnt
                         
        If ComSymbol = "" OR IsNull(ComSymbol) Then                                  '�񱳹��ڰ� ������� 
        Else                                                                         '�񱳹��ڰ� �����Ұ�� �񱳹��ڸ� �� �������� �Է°������Ѵ�.
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
    
    If InPos = "" Then                                                              '�������̰� �������� ���� ��� �Է¹��� ���̰� �������̰� �ȴ�.
        InPos = Cnt  
    End If
    
    If UCase(Trim(direct)) = "LEFT" OR UCase(Trim(direct)) = "" Then                '���������ϰ��(defalut) �������� ���� ���� ������ ���ڰ� �ԷµǸ� ������ ����(default)�κ��� ��ü���ڷ� ü���.
        If InPos > Cnt Then                                                         ' ex:hi    
            If strFix = "" Then
               strFix = Chr(32)
            End if 
            For i = (Cnt+1) To InPos        
                InValue = InValue & strFix
            Next         
        End If
    ElseIf UCase(Trim(direct)) = "RIGHT" Then                                         '���������� 
        If InPos > Cnt Then                                                           ' ex:     hi
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
