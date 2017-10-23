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
			
    Dim strMode

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "MB")

    lgErrorStatus     = "NO"													    
    lgCurrentSpd      = Request("lgCurrentSpd")                                      '☜: "M"(Spread #1) "S"(Spread #2)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    strMode = Request("txtMode")                                                                '☜ : 현재 상태를 받음 
    lgKeyStream = Request("txtComp_cd") & gColSep
    lgKeyStream = lgKeyStream & Request("txtFr_acq_dt") & gColSep
    lgKeyStream = lgKeyStream & Request("txtTo_acq_dt") & gColSep
    lgKeyStream = lgKeyStream & Request("txtReportDt")  & gColSep
    lgKeyStream = lgKeyStream & Request("gSelframeFlg") & gColSep

   Call SubOpenDB(lgObjConn)      
    
   Dim strFilePath
   Dim Pinfo,Fnm,CFnm,Pnm,FPnm      
   Dim SFnm,iDx,Fso,DFnm,CTFnm
   Dim xdn
    
   Select Case strMode
   
        Case CStr(UID_M0001)        
	                                                                                            '☜: Protect system from crashing	        
           Set Fso = CreateObject("Scripting.FileSystemObject")  

           SFnm = "_"
           iDx = 1
           Pinfo = Request.ServerVariables ("PATH_INFO")                                       '현재File의 경로를 받는다.
           
           Fnm = Mid(Pinfo,InstrRev(Pinfo,"/")+1,InstrRev(Pinfo,".")-InstrRev(Pinfo,"/")-1)    'File의 경로중 File Name만 저장 
           Pnm = Mid(Pinfo,1,InstrRev(Pinfo,"/")+1)                                            'File Name 부분을 뺀 나머지 경로를 저장 
           FPnm = Server.MapPath("../../files/u2000/" & Fnm & SFnm & iDx)           '경로를 System 디렉토리로 바꾼다.
           Do While Fso.FileExists (Fpnm)                                                      'Server쪽에 생성될 File Name이 중복방지 
           
               iDx = Mid(FPnm,InstrRev(FPnm,"_")+1)                                            
               iDx = iDx + 1        
               FPnm = Server.MapPath("../../files/u2000/" & Fnm & SFnm & iDx)       '"_" & 숫자 를 붙여 화일의 전체 디렉토리경로를 저장         
           
           Loop

           Call SubBizQueryMulti(lgKeyStream)
   
           If UCase(Trim(lgErrorStatus)) <> "YES" Then
               Set CTFnm = Fso.CreateTextFile (Fpnm,False)                                         'text를 저장할 File을 생성            
               CTFnm.Write lgstrData                                                                'Text 내용부분                       
               DFnm = Fso.GetFileName(FPnm)            
               CTFnm.close    
               Set CTFnm = nothing
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
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti(arrKey)

    Dim iDx,arrSplitKey
    Dim stryymm, strWhere   
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear             
    
    arrSplitKey = Split(arrKey,gColSep)
    
    stryymm = FilterVar(UNIConvDateToYYYYMM(arrSplitKey(3),gServerDateFormat,""),"''" ,"S")                                                    '☜: Clear Error status    

    If arrSplitKey(4)="1" Then
        lgCurrentSpd = "M"
        strWhere = stryymm
        strWhere = strWhere & " And ( hdf020t.emp_no = hdf070t.emp_no ) "
        strWhere = strWhere & " And ( hdf070t.prov_type = " & FilterVar("1", "''", "S") & "  ) "
        strWhere = strWhere & " And ( IsNull(hdf020t.anut_acq_dt,hdf020t.entr_dt) >= " & FilterVar(UNIConvDateCompanyToDB(arrSplitKey(1),NULL),"NULL", "S") & ") "
        strWhere = strWhere & " And ( IsNull(hdf020t.anut_acq_dt,hdf020t.entr_dt) <= " & FilterVar(UNIConvDateCompanyToDB(arrSplitKey(2),NULL),"NULL", "S") & ") "        
        
        Call SubMakeSQLStatements("MR",strWhere,"X",C_EQ,arrSplitKey(0),arrSplitKey(3))                              '☆: Make sql statements
    Else
        lgCurrentSpd = "S"
        strWhere = FilterVar(UNIConvDateCompanyToDB(arrSplitKey(1),NULL),"NULL", "S")
        strWhere = strWhere & " And ( hdf020t.anut_loss_dt <= " & FilterVar(UNIConvDateCompanyToDB(arrSplitKey(2),NULL),"NULL", "S") & ") "        
        
        Call SubMakeSQLStatements("MR",strWhere,"X",">=",arrSplitKey(0),arrSplitKey(3))                              '☆: Make sql statements        
    End If    

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then        
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else    
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""        
        iDx = 1         
        Do While Not lgObjRs.EOF
            Select Case lgCurrentSpd
               Case "M" 	
                    lgstrData = lgstrData & "KI05"
                    lgstrData = lgstrData & UCase(ConvSPChars(SetFixSrting(lgObjRs("comp_cd"),"","",8)))
                    lgstrData = lgstrData & (String((5-Len(Cstr(lgStrPrevKey+iDx))),"0") & (lgStrPrevKey+iDx))
                    lgstrData = lgstrData & ConvSPChars(SetFixSrting(lgObjRs("name"),"","",20))                   
                    lgstrData = lgstrData & ConvSPChars(SetFixSrting(lgObjRs("anut_no"),"","",14))
                    lgstrData = lgstrData & (String((8-Len(Cstr(lgObjRs("pay_tot_amt")))),"0") & lgObjRs("pay_tot_amt"))
                    lgstrData = lgstrData & ConvSPChars(SetFixSrting(lgObjRs("anut_grade"),"","",2))
                    lgstrData = lgstrData & "000000"
                    lgstrData = lgstrData & ConvSPChars(lgObjRs("anut_acq_dt"))
                    lgstrData = lgstrData & "00000"
              Case Else
                    lgstrData = lgstrData & "KI05"
                    lgstrData = lgstrData & UCase(ConvSPChars(SetFixSrting(lgObjRs("comp_cd"),"","",8)))
                    lgstrData = lgstrData & (String((5-Len(Cstr(lgStrPrevKey+iDx))),"0") & (lgStrPrevKey+iDx))
                    lgstrData = lgstrData & ConvSPChars(SetFixSrting(lgObjRs("name"),"","",20))
                    lgstrData = lgstrData & ConvSPChars(SetFixSrting(lgObjRs("anut_no"),"","",14))
                    lgstrData = lgstrData & SetFixSrting("","","",14) 
                    lgstrData = lgstrData & "00"
                    lgstrData = lgstrData & lgObjRs("anut_loss_dt")
                    lgstrData = lgstrData & "00000"
               End Select      
            
            lgstrData = lgstrData & Chr(13) & Chr(10)
		    lgObjRs.MoveNext
            iDx =  iDx + 1               
        Loop         
    End If 
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
   Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
End Sub    
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp,txtComp_cd,txtReportDt)    

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    If lgCurrentSpd = "M" Then
       lgStrSQL = "Select hdf020t.name, hdf020t.anut_no + " & FilterVar("0", "''", "S") & "  anut_no, hdf070t.pay_tot_amt, hdf020t.anut_grade, CONVERT(VARCHAR(8), IsNull(hdf020t.anut_acq_dt,hdf020t.entr_dt), 112) anut_acq_dt , " & FilterVar(txtComp_cd, "''", "S") & " AS comp_cd "
       lgStrSQL = lgStrSQL & " From  hdf020t, hdf070t "
       lgStrSQL = lgStrSQL & " Where hdf070t.pay_yymm " & pComp & pCode
    Else
       lgStrSQL = "Select hdf020t.name,  hdf020t.anut_no + " & FilterVar("0", "''", "S") & "  anut_no, hdf020t.anut_grade, CONVERT(VARCHAR(8), hdf020t.anut_loss_dt, 112)  anut_loss_dt, " & FilterVar(txtComp_cd, "''", "S") & " AS comp_cd "
       lgStrSQL = lgStrSQL & " From  hdf020t   "
       lgStrSQL = lgStrSQL & " Where anut_loss_dt " & pComp & pCode
    End If   
End Sub
'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
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
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
'    On Error Resume Next                                                              '☜: Protect system from crashing
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
'============================================================================================================
' Name : SetFixSrting
' Desc : This Function return srting
'============================================================================================================
Function SetFixSrting(InValue, ComSymbol, strFix, InPos)
    Dim Cnt,MCnt
    Dim ExSymbol,strSplit,strMid
    Dim iDx,i
    
    If InValue = "" OR IsNull(InValue) Then        
        Cnt = 0     
    Else

        Cnt = Len(InValue)
        For i = 1 To Cnt
            strMid = Mid(InValue,i,1)
            If Asc(strMid) > 255 OR Asc(strMid) < 0 Then
                MCnt = MCnt + 2
            Else
                MCnt = MCnt + 1
            End If
        Next
        Cnt = MCnt
                 
        If ComSymbol = "" OR IsNull(ComSymbol) Then
        Else
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
    
    If InPos = "" Then
        InPos = Cnt  
    End If
        
    If InPos > Cnt Then    
        If strFix = "" Then
           strFix = Chr(32)
        End if 
        For i = (Cnt+1) To InPos        
            InValue = InValue & strFix
        Next         
    End If
    
    SetFixSrting = InValue
End Function


%>
<script language="vbscript">
		Dim SF
		On Error Resume Next

		Set SF = CreateObject("uni2kcm.SaveFile")
		
		Call SF.SaveTextFile("<%= strFilePath %>")

		Set SF = Nothing
</script>

