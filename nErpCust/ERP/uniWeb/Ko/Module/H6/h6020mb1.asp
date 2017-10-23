<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%		
    Dim strMode
    			
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "BB")

    Call HideStatusWnd                                                               '��: Hide Processing message

    lgErrorStatus     = "NO"													
    
    strMode = Request("txtMode")                                                                '�� : ���� ���¸� ����    
    lgKeyStream = Request("txtPay_yymm_dt")                  & gColSep
    lgKeyStream = lgKeyStream & Request("txtProv_type")      & gColSep
    lgKeyStream = lgKeyStream & Request("txtYy_mm_dd_dt")    & gColSep
    lgKeyStream = lgKeyStream & Request("txtcboPay_cd")      & gColSep
    lgKeyStream = lgKeyStream & Request("txtSect_cd")        & gColSep
    lgKeyStream = lgKeyStream & Request("txtOcpt_type")      & gColSep
    lgKeyStream = lgKeyStream & Request("txtFr_dept_cd")     & gColSep
    lgKeyStream = lgKeyStream & Request("txtTo_dept_cd")     & gColSep
    lgKeyStream = lgKeyStream & Request("txtGigup_type")     & gColSep
    lgKeyStream = lgKeyStream & UniConvNum(Request("txtStand_amt"),0)      & gColSep

    Call SubOpenDB(lgObjConn)      

    Dim strFilePath
    Dim Pinfo,Fnm,CFnm,Pnm,FPnm      
    Dim SFnm,iDx,Fso,DFnm,CTFnm
    Dim xdn

    Select Case strMode
	    Case CStr(UID_M0001)        
                                                                                       '��: Protect system from crashing	        
            Set Fso = CreateObject("Scripting.FileSystemObject")  

            SFnm = "_"
            iDx = 1
            Pinfo = Request.ServerVariables ("PATH_INFO")                         'request vitual path(����File�� ��θ� �޴´�.)
            Fnm = Mid(Pinfo,InstrRev(Pinfo,"/")+1,InstrRev(Pinfo,".")-InstrRev(Pinfo,"/")-1)    'find only file name(File�� ����� File Name�� ����)
            Pnm = Mid(Pinfo,1,InstrRev(Pinfo,"/")+1)                                           'File Name �κ��� �� ������ ��θ� ���� 
'            FPnm = Server.MapPath("../../files/" & gCompany & "/" & Fnm & SFnm & iDx)   'change vitual path into physical path(��θ� System ���丮�� �ٲ۴�.)
            FPnm = Server.MapPath("../../files/u2000/" & Fnm & SFnm & iDx)   'change vitual path into physical path(��θ� System ���丮�� �ٲ۴�.)
            Do While Fso.FileExists (Fpnm)                                 'Server�ʿ� ������ File Name �ߺ����� 
                iDx = Mid(FPnm,InstrRev(FPnm,"_")+1)                                            
                iDx = iDx + 1        
'                FPnm = Server.MapPath("../../files/" & gCompany & "/" & Fnm & SFnm & iDx)       '"_" & ���� �� �ٿ� ȭ���� ��ü ���丮��θ� ����         
                FPnm = Server.MapPath("../../files/u2000/" & Fnm & SFnm & iDx)       '"_" & ���� �� �ٿ� ȭ���� ��ü ���丮��θ� ����         
            Loop

            Call SubBizQueryMulti(lgKeyStream)


            If UCase(Trim(lgErrorStatus)) <> "YES" Then
                Set CTFnm = Fso.CreateTextFile (Fpnm,true)                                         'text�� ������ File�� ����           
   
            
                CTFnm.Write lgstrData                                                                'Text ����κ�           
          
                DFnm = Fso.GetFileName(FPnm)            
                CTFnm.close    
   
              
                Set CTFnm = nothing
            Else
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

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti(arrKey)
    Dim iDx,arrSplitKey,iPos
    Dim strWhere,ld_count2,id_real_prov_amt
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear             
       
    arrSplitKey = Split(arrKey,gColSep)
    
    If IsNull(arrSplitKey(3))   OR Trim(arrSplitKey(3)) = "" Then   arrSplitKey(3)  = "%"
	If IsNull(arrSplitKey(4))   OR Trim(arrSplitKey(4)) = "" Then   arrSplitKey(4)  = "%"
    If IsNull(arrSplitKey(5))   OR Trim(arrSplitKey(5)) = "" Then 	arrSplitKey(5)  = "%"
    If IsNull(arrSplitKey(6))   OR Trim(arrSplitKey(6)) = "" Then   arrSplitKey(6)  = ""
    If IsNull(arrSplitKey(7))   OR Trim(arrSplitKey(7)) = "" Then   arrSplitKey(7)  = "zzzzzzz"
    
    strWhere = FilterVar(arrSplitKey(0), "''", "S")
	strWhere = strWhere & " AND b.prov_type	= "        & FilterVar(arrSplitKey(1), "''", "S")
	strWhere = strWhere & " AND b.real_prov_amt > 0"
	strWhere = strWhere & " AND b.internal_cd >= "     & FilterVar(arrSplitKey(6), "''", "S")
	strWhere = strWhere & " AND b.internal_cd <= "     & FilterVar(arrSplitKey(7), "''", "S")
	strWhere = strWhere & " AND a.ocpt_type	like "     & FilterVar(arrSplitKey(5), "''", "S")
	strWhere = strWhere & " AND c.sect_cd LIKE "       & FilterVar(arrSplitKey(4), "''", "S")
	strWhere = strWhere & " AND a.pay_cd LIKE "        & FilterVar(arrSplitKey(3), "''", "S")
        
    Call SubMakeSQLStatements("MR",strWhere,"X","=")                              '��: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then        
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
        Call SetErrorStatus()
    Else    

        lgstrData = ""
        ld_count2 = 0
        iDx = 1                        
        
        Call HeadEndWrite ("HEAD","","",arrKey)        
        Do While Not lgObjRs.EOF
        
            If Cdbl(lgObjRs("real_prov_amt")) > 0 Then
                lgstrData = lgstrData & "D"
                lgstrData = lgstrData & "21"
                lgstrData = lgstrData & SetFixSrting(iDx,"","0",6,"RIGHT")
                lgstrData = lgstrData & "1"
            
                If isNull(lgObjRs("bank_accnt")) OR Trim(ConvSPChars(lgObjRs("bank_accnt"))) = "" Then                
                    Call DisplayMsgBox("800434", vbInformation,ConvSPChars(lgObjRs("emp_no")),"x",I_MKSCRIPT)   '�� �ٲ�κ� 
			        lgstrData = lgstrData & SetFixSrting("error","","",14,"")
			    Else
			        lgstrData = lgstrData & SetFixSrting(ConvSPChars(lgObjRs("bank_accnt")),"-","",14,"")
		        End If
		        
		        id_real_prov_amt = 0
		        id_real_prov_amt = Cdbl(lgObjRs("real_prov_amt"))
		        
		        ld_count2 = Cdbl(ld_count2) + id_real_prov_amt
		        If arrSplitKey(8) = "1" Then
		            lgstrData = lgstrData & SetFixSrting((id_real_prov_amt - CDbl(arrSplitKey(9))),"","0",11,"RIGHT")
		        ElseIf arrSplitKey(8) = "2" Then
			        If Cdbl(id_real_prov_amt) > arrSplitKey(9) Then
		   	            lgstrData = lgstrData & SetFixSrting(CDbl(arrSplitKey(9)),"","0",11,"RIGHT")
		   	        Else
		   	            lgstrData = lgstrData & SetFixSrting(id_real_prov_amt,"","0",11,"RIGHT")
			        End If
			    Else
			        lgstrData = lgstrData & SetFixSrting(id_real_prov_amt,"","0",11,"RIGHT")
	            End If

	            lgstrData = lgstrData & SetFixSrting("xxxxxxxxxxxxx","","",13,"")	             
                lgstrData = lgstrData & "0"
                lgstrData = lgstrData & "00"
                lgstrData = lgstrData & SetFixSrting("","","",9,"")
                lgstrData = lgstrData & SetFixSrting("","","",17,"")
                lgstrData = lgstrData & "*"
            Else
                Call DisplayMsgBox("800076", vbInformation,"x","x",I_MKSCRIPT)   '�� �ٲ�κ� 
            End If				 

            lgstrData = lgstrData & Chr(13) & Chr(10)
		    lgObjRs.MoveNext
            iDx =  iDx + 1       

        Loop       

        Call HeadEndWrite ("END",iDx-1,ld_count2,arrKey)
    
    End If     
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '��: Release RecordSSet
End Sub    
'============================================================================================================
' Name :  
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)    

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    
    lgStrSQL = " SELECT"
    lgStrSQL = lgStrSQL & " a.emp_no, a.name, a.bank, a.bank_accnt, b.real_prov_amt"
    lgStrSQL = lgStrSQL	& " FROM hdf020t a, hdf070t b, haa010t c"
	lgStrSQL = lgStrSQL	& " WHERE  a.emp_no = b.emp_no"
	lgStrSQL = lgStrSQL	& " AND a.emp_no = c.emp_no"
	lgStrSQL = lgStrSQL	& " AND b.pay_yymm " & pComp & pCode
End Sub
'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '��: Set error status
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
    On Error Resume Next                                                              '��: Protect system from crashing
    Err.Clear                                                                         '��: Clear Error status

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
Function SetFixSrting(InValue, ComSymbol, strFix, InPos, direct)
    Dim Cnt,MCnt
    Dim ExSymbol,strSplit,strMid
    Dim iDx,i,strTemp
    
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
    
    If UCase(Trim(direct)) = "LEFT" OR UCase(Trim(direct)) = "" Then   
        If InPos > Cnt Then    
            If strFix = "" Then
               strFix = Chr(32)
            End if 
            For i = (Cnt+1) To InPos        
                InValue = InValue & strFix
            Next         
        End If
    ElseIf UCase(Trim(direct)) = "RIGHT" Then
        If InPos > Cnt Then    
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
'============================================================================================================
' Name : Head_End_TitleWrite
' Desc : text file�� ����κа� ��� �κ��� ����.
'============================================================================================================
Sub HeadEndWrite(HEType,li_count1,ld_count2,arrKey)
Dim arrSplitKey,TempSplit,i,iPos
Dim EndDate
Dim strHead
    
    If UCase(Trim(HEType)) = "HEAD" Then        

        arrSplitKey = Split(arrKey,gColSep)    
        EndDate = Mid(Year(Date),3,2) & Right("0" & Month(Date),2) & Right("0" & Day(Date),2)       
        If Instr(arrSplitKey(2),"-") > 0 Or Len(arrSplitKey(2)) > 6 Then
            TempSplit = Split(arrSplitKey(2),"-")
            If UBound(TempSplit) > 0 Then
                For i=0 To Ubound(TempSplit)
                    arrSplitKey(2) = arrSplitKey(2) & TempSplit(i)
                Next
            End If
            arrSplitKey(2) = Right(arrSplitKey(2),6)
        ElseIf Instr(arrSplitKey(2),"/") > 0 Or Len(arrSplitKey(2)) > 6 Then
            TempSplit = Split(arrSplitKey(2),"/")
            If UBound(TempSplit) > 0 Then
                For i=0 To Ubound(TempSplit)
                    arrSplitKey(2) = arrSplitKey(2) & TempSplit(i)
                Next
            End If
            arrSplitKey(2) = Right(arrSplitKey(2),6)        
        End IF        
        
        lgstrData = lgstrData & "S"
        lgstrData = lgstrData & "21"
        lgstrData = lgstrData & "XXXXXXX"
        lgstrData = lgstrData & "1"
        lgstrData = lgstrData & "1"
        lgstrData = lgstrData & "1"
        lgstrData = lgstrData & arrSplitKey(1)
	    lgstrData = lgstrData & SetFixSrting("9999","","",4,"")
        lgstrData = lgstrData & EndDate
        lgstrData = lgstrData & arrSplitKey(2)
	    lgstrData = lgstrData & "3"
	    lgstrData = lgstrData & SetFixSrting("JH","","",8,"")
	    lgstrData = lgstrData & SetFixSrting("XXXXXXX","","",8,"")
	    lgstrData = lgstrData & SetFixSrting("XXX","","",8,"")
	    lgstrData = lgstrData & SetFixSrting("XXXXXXXXXXX","","",11,"RIGHT")
	    lgstrData = lgstrData & SetFixSrting("","","",11,"")
	    lgstrData = lgstrData & "*"
        lgstrData = lgstrData & Chr(13) & Chr(10)
    Else
'        If Len(ld_count2) > 11 Then
'            iPos = instr(ld_count2,".")
'            ld_count2 =Round(ld_count2,11-iPos)
'        End If
            
        lgstrData = lgstrData & "E"
        lgstrData = lgstrData & SetFixSrting(li_count1+2,"","0",5,"RIGHT")
		lgstrData = lgstrData & SetFixSrting(li_count1,"","0",5,"RIGHT")
	    lgstrData = lgstrData & SetFixSrting(ld_count2,"","0",11,"RIGHT")
		lgstrData = lgstrData & SetFixSrting("0","","0",5,"RIGHT")
		lgstrData = lgstrData & SetFixSrting("0","","0",11,"RIGHT")
	    lgstrData = lgstrData & SetFixSrting("0","","0",5,"RIGHT")
	    lgstrData = lgstrData & SetFixSrting("0","","0",11,"")
	    lgstrData = lgstrData & SetFixSrting("","","",5,"")
	    lgstrData = lgstrData & SetFixSrting("","","",18,"")
        lgstrData = lgstrData & "*"            
    End If        
End Sub

%>
<script language="vbscript">
		Dim SF
		On Error Resume Next
		Set SF = CreateObject("uni2kCM.SaveFile")
		Call SF.SaveTextFile("<%= strFilePath %>")

		Set SF = Nothing
</script>

