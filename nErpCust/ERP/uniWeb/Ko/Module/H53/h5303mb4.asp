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
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
	
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "MB")

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Dim strFilePath,strMode,Pinfo,iDx
    Dim Fnm,CFnm,FPnm      
    Dim Fso,DFnm,CTFnm   
    
    lgCurrentSpd      = Request("lgCurrentSpd")                                     '☜: "M"(Spread #1) "S"(Spread #2)
    lgstrData = ""


    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    Select Case lgOpModeCRUD
	    Case CStr(UID_M0001)                                                            'data select and create File on server     	        
            Set Fso = CreateObject("Scripting.FileSystemObject")
				iDx = 1
                Pinfo = Request.ServerVariables ("PATH_INFO")
   
             
	            Fnm = Mid(Pinfo,InstrRev(Pinfo,"/")+1,InstrRev(Pinfo,".")-InstrRev(Pinfo,"/")-1)    'File의 경로중 File Name만 저장 
				FPnm = Server.MapPath("../../files/u2000/" & Fnm & "_" & iDx)           '경로를 System 디렉토리로 바꾼다.


				Do While Fso.FileExists (Fpnm)                                                      'Server쪽에 생성될 File Name이 중복방지 
           
				    iDx = Mid(FPnm,InstrRev(FPnm,"_")+1)                                            
				    iDx = iDx + 1        
				    FPnm = Server.MapPath("../../files/u2000/" & Fnm & "_" & iDx)       '"_" & 숫자 를 붙여 화일의 전체 디렉토리경로를 저장         
           
				Loop  
				         
                Call SubBizQuery()

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
	    Case CStr(UID_M0002)
		    Err.Clear 
		    Call HideStatusWnd
		    strFilePath = "http://" & Request.ServerVariables("LOCAL_ADDR") & ":" _
		    			   & Request.ServerVariables("SERVER_PORT")
	        If Instr(1, Request.ServerVariables("URL"), "Module") <> 0 Then
	            strFilePath = strFilePath & Mid(Request.ServerVariables("URL"), 1, InStr(1, Request.ServerVariables("URL"), "Module") - 1)     
	        End If
		    strFilePath = strFilePath  & "files/u2000/" 
		    strFilePath = strFilePath & Request("txtFileName")
	End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
	Dim itxtSpread
	Dim itxtSpreadArr
	Dim itxtSpreadArrCount
 
	Dim iCUCount
	Dim iDCount
 
	Dim ii   
	
	itxtSpread = ""
	             
	iCUCount = Request.Form("txtCUSpread").Count

	itxtSpreadArrCount = -1
	             
	ReDim itxtSpreadArr(iCUCount)
	             
	For ii = 1 To iCUCount
	    itxtSpreadArrCount = itxtSpreadArrCount + 1
	    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
	Next
 
	itxtSpread = Join(itxtSpreadArr,"")	 
	lgstrData = lgstrData & replace(itxtSpread,chr(11),Chr(13) & Chr(10))

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
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
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

<script language="vbscript">
		Dim SF
		On Error Resume Next
		Set SF = CreateObject("uni2kCM.SaveFile")
	
		Call SF.SaveTextFile("<%= strFilePath %>")

		Set SF = Nothing
</script>
