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
	Dim strFilePath,strMode
    Dim Fnm,CFnm,FPnm      
    Dim Fso,DFnm,CTFnm    

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "MB")
    Call HideStatusWnd                                                              '☜: Hide Processing message

    lgErrorStatus     = "NO"													    
    strMode			= Request("txtMode")                                               '☜: Read Operation Mode (CRUD)
	lgKeyStream		= Split(Request("txtKeyStream"),gColSep)

	
	Call SubOpenDB(lgObjConn)      

    Select Case strMode
	    Case CStr(UID_M0001)                                                            'data select and create File on server     	        

            Set Fso = CreateObject("Scripting.FileSystemObject")  
            

                Fnm = Fso.GetFileName("C:\e" &  Request("txtFileName"))                
                FPnm = Server.MapPath("../../files/u2000/" & Fnm)  '2002.02.01 /files 에는 현재 u2000만 존재:나중에 공통쪽 변경되면 수정해야 함.
         
                Call SubBizQuery()
 
            If UCase(Trim(lgErrorStatus)) <> "YES" Then

                Set CTFnm = Fso.CreateTextFile(FPnm,True)                              'text를 저장할 File을 생성            
             
                CTFnm.Write lgstrData                                                   'Text 내용부분                       
                DFnm = Fso.GetFileName(FPnm)            

                CTFnm.close    
                Set CTFnm = nothing
                
            Else
                Call DisplayMsgBox("800004", vbInformation, "", "", I_MKSCRIPT)      '☜ :파일생성오류 
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
	    strFilePath = strFilePath  & "files/u2000/"    '2002.02.01 /files 에는 현재 u2000만 존재:나중에 공통쪽 변경되면 수정해야 함.
	    strFilePath = strFilePath & Request("txtFileName")

End Select
 
'============================================================================================================
' Name : ASubBizQueryMulti()
' Desc : Query ASheet Data from Db
'============================================================================================================
Sub SubBizQuery()
	Dim lgStrSQL
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
        
    lgstrData = ""
    
    lgStrSQL = "SELECT  REPLACE(DATA,CHAR(11),'') data	FROM H0000T ORDER BY YEAR_AREA_CD,EMP_NO,TYPE ,right(rtrim(DATA) ,2) "
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
       Call SetErrorStatus("")
    Else        
        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & lgObjRs("DATA")
            lgstrData = lgstrData & Chr(13) & Chr(10)
            lgObjRs.MoveNext
        Loop
    End If
    
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)    
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

<script language="vbscript">
		Dim SF
		On Error Resume Next
		Set SF = CreateObject("uni2kCM.SaveFile")
		Call SF.SaveTextFile("<%= strFilePath %>")

		Set SF = Nothing
</script>
