<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../wcm/inc_SvrDebug.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
    
    '---------------------------------------Common-----------------------------------------------------------                                                              '☜: Hide Processing message
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgIntFlgMode = CInt(Request("txtFlgMode"))		

    'Multi SpreadSheet
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)    
    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Dim lgSoSeq
    Dim L1_auto_code
    Dim lgQueryChain
    Dim lgDataError
	Dim iArrTotal
'    ReDim L1_auto_code(lgLngMaxRow)

    Function RtnQueryVal(strField,strFrom,strWhere)
        Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	    RtnQueryVal = ""
	    Call CommonQueryRs(strField,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	    RtnQueryVal = Replace(lgF0,Chr(11),"")
	    If RtnQueryVal = "X" Or trim(RtnQueryVal) = "" Or IsNull(RtnQueryVal) Then
            Call DisplayMsgBox("970000", vbInformation, strWhere & strField, "", I_MKSCRIPT)
           
            ObjectContext.SetAbort
            Call SetErrorStatus
		End If
    End Function
    
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    Call SubOpenDB(lgObjConn)
    
    Call CheckVersion(lgKeyStream(1), lgKeyStream(2))	' 2005-03-11 버전관리기능 추가 
     
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query        
        
     
             Call SubBizQueryMulti()
             Call SubBizQueryMulti2()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
             
             
             if lgErrorStatus <> "YES" then
                 Call SubBizSaveMulti()
             end if  
            
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
        Case CStr(UID_M0005)
             Call SubBizAutoQuery()       
        Case CStr(UID_M0006)
             Call SubBizSaveMultiDeleteBtn()
    End Select
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizSaveMultiDeleteBtn
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDeleteBtn()
                                                                    '☜: Clear Error status
       
    
End Sub

    

Sub SubBizQueryMulti()
'''''''''''''''''''''''''''''''''''''''''''''''''''''
	Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere
   
    Dim iClsRs
    Dim iTemp,i
    Dim k
    
    On Error Resume Next
    Err.Clear                                                               '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'iKey1 = FilterVar(lgKeyStream(0),"''", "S")
    

    strWhere = " a.co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
	strWhere = strWhere & "  and a.fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
	strWhere = strWhere & "  and a.rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")


    Call SubMakeSQLStatements("MR",strWhere,"X","")                              '☜ : Make sql statements
    
    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        iClsRs = 1
        'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else
  
         Response.Write "<Script Language=vbscript>" & vbCr
	     Response.Write "With parent.frm1" & vbCr
         If lgObjRs("W1") = "1" then
		     Response.Write "   .chkW1.checked = true " & vbCr 
		 Else
		     Response.Write "   .chkW1.checked = False " & vbCr    
		 end if
		 
		 If lgObjRs("W2") = "1" then
		     Response.Write "   .chkW2.checked = true " & vbCr 
		 Else
		     Response.Write "   .chkW2.checked = False " & vbCr    
		 end if
		 
		 If lgObjRs("W3") = "1" then
		    Response.Write "   .chkW3.checked = true " & vbCr 
		 Else
		     Response.Write "   .chkW3.checked = False " & vbCr    
		 end if
  		 
  		 If lgObjRs("W4") = "1" then
		    Response.Write "    .chkW4.checked = true " & vbCr 
		 Else
		     Response.Write "   .chkW4.checked = False " & vbCr    
		 end if
		 
		 If lgObjRs("W5") = "1" then
		    Response.Write "    .chkW5.checked = true " & vbCr 
		 Else
		     Response.Write "   .chkW5.checked = False " & vbCr    
		 end if
		     
  		 Response.Write " .txtChange_DT.text    = """& ConvSPChars(lgObjRs("w7"))  &"""" & vbCr   
  		 Response.Write " .cboW13_A.Value    = """& ConvSPChars(lgObjRs("w13_A"))  &"""" & vbCr
  		 Response.Write " .cboW13_B.Value    = """& ConvSPChars(lgObjRs("w13_B"))  &"""" & vbCr 
  		 Response.Write " .txtW13_C.Value    = """& ConvSPChars(lgObjRs("w13_C"))  &"""" & vbCr
  		 Response.Write " .cboW14_A.Value    = """& ConvSPChars(lgObjRs("w14_A"))  &"""" & vbCr
  		 Response.Write " .cboW14_B.Value    = """& ConvSPChars(lgObjRs("w14_B"))  &"""" & vbCr 
  		 Response.Write " .txtW14_C.Value    = """& ConvSPChars(lgObjRs("w14_C"))  &"""" & vbCr  
  		 Response.Write " .cboW15_A.Value    = """& ConvSPChars(lgObjRs("w15_A"))  &"""" & vbCr
  		 Response.Write " .cboW15_B.Value    = """& ConvSPChars(lgObjRs("w15_B"))  &"""" & vbCr 
  		 Response.Write " .txtW15_C.Value    = """& ConvSPChars(lgObjRs("w15_C"))  &"""" & vbCr 
  		   
  		   
	     Response.Write " End With "	& vbCr
         Response.Write "</Script>"  & vbCr
   
    

  end if  

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs) 
End Sub
Sub SubBizQueryMulti2()
'''''''''''''''''''''''''''''''''''''''''''''''''''''
	Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere
   
    Dim iClsRs
    Dim iTemp,i
    Dim k
    
    On Error Resume Next
    Err.Clear                                                               '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'iKey1 = FilterVar(lgKeyStream(0),"''", "S")
    
            	
		strWhere = " co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
		strWhere = strWhere & "  and fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
		strWhere = strWhere & "  and rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")

    Call SubMakeSQLStatements("MD",strWhere,"X","")                              '☜ : Make sql statements
    
    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        iClsRs = 1
      '  Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
      '  Call SetErrorStatus()
    Else
  
        
		
             
       lgstrData = ""
        iDx       = 1
       

         
        Do While Not lgObjRs.EOF
           
          
                lgstrData = lgstrData & Chr(11) & UNIConvNum(lgObjRs("seq_no"),0)
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("w8"))
                lgstrData = lgstrData & Chr(11) &  REPLACE(ConvSPChars(lgObjRs("w9_Fr")),0,"")
                lgstrData = lgstrData & Chr(11) & "~"
                lgstrData = lgstrData & Chr(11) & REPLACE(ConvSPChars(lgObjRs("w9_To")),0,"")
                lgstrData = lgstrData & Chr(11) & REPLACE(ConvSPChars(lgObjRs("w10")),0,"")
                lgstrData = lgstrData & Chr(11) & REPLACE(ConvSPChars(lgObjRs("w11")),0,"")
                lgstrData = lgstrData & Chr(11) & REPLACE(ConvSPChars(lgObjRs("w12")),0,"")

 
 
	
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
				lgstrData = lgstrData & Chr(11) &  iDx
				lgstrData = lgstrData & Chr(11) & Chr(12)
			
		    lgObjRs.MoveNext
		 



        Loop 
  

    

  end if  

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs) 
End Sub



Sub SubBizAutoQuery()

                                         '☜: Release RecordSSet

End Sub    


'============================================================================================================
' Name : SubBizSave
' Desc : Save Data
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear
   Dim strWhere
  
     lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
	                        '☜ : Make sql statements

	
	Select Case lgIntFlgMode
				    Case  OPMD_CMODE    
				          	
							strWhere = " a.co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
							strWhere = strWhere & "  and a.fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
							strWhere = strWhere & "  and a.rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")

    
							Call SubMakeSQLStatements("MR",strWhere,"X","")                                                               '☜: Create
				          If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
							    lgStrPrevKeyIndex = ""
							    
							    Call DisplayMsgbox("WC0001","X",Trim(UCase(Request("txtFISC_YEAR"))) ,"X" ,I_MKSCRIPT)      '☜ : No data is found. 
							    
							    
		
							    Call SetErrorStatus()
							    lgErrorStatus = "YES"
							   
							Else
                                Call SubBizSaveSingleCreate()
								    if lgErrorStatus <> "YES" then
									    Call SubBizSaveSingleCreate2()
									end if   	             	
							end if	
		
				          
				    Case  OPMD_UMODE  
				                                                                  '☜: Update
				           Call SubBizSaveSingleUpdate()
				           if lgErrorStatus <> "YES" then
							   Call SubBizSaveSingleUpdate2()
						   end if  
				    
	End Select
	
			
	    Call SubCloseRs(lgObjRs)		
End Sub

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
   On Error Resume Next                                                             '☜: Protect system from crashing
   Err.Clear
        lgStrSQL = "delete from TB_63H " 
		lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
		lgStrSQL = lgStrSQL & "  and fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
		lgStrSQL = lgStrSQL & "  and rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")
		lgStrSQL = lgStrSQL & " delete from TB_63B " 
		lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
		lgStrSQL = lgStrSQL & "  and fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
		lgStrSQL = lgStrSQL & "  and rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")
		lgStrSQL = lgStrSQL & " delete from TB_63A " 
		lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
		lgStrSQL = lgStrSQL & "  and fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
		lgStrSQL = lgStrSQL & "  and rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")
         '---------- Developer Coding part (End  ) ---------------------------------------------------------------
 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

	Dim arrRowVal
    Dim arrColVal
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere

    On Error Resume Next
    Err.Clear                                                                        '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	
            	
	strWhere = " a.co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
	strWhere = strWhere & "  and a.fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
	strWhere = strWhere & "  and a.rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")
    
    Call SubMakeSQLStatements("MR",strWhere,"X","")                              '☜ : Make sql statements
    
    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = false  and lgIntFlgMode =OPMD_CMODE Then
        lgStrPrevKeyIndex = ""
        
        Call DisplayMsgbox("WC0001","X",Trim(UCase(Request("txtFISC_YEAR"))) ,"X" ,I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
        lgQueryChain = 0
    Else
        lgQueryChain = 1
   
            
	        arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data

            For iDx = 1 To lgLngMaxRow
                arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data

                Select Case arrColVal(0)
                    Case "C"
                        Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
                    Case "U"
                        Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
                    Case "D"
                        Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
                End Select
                
                If lgErrorStatus    = "YES" Then
                   lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
                   Exit For
                End If
                
            Next
          
end if
            
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    dim iObjPS5G115
    On Error Resume Next
    Err.Clear                                                                        '☜: Clear Error status
   
    
    lgStrSQL = "INSERT INTO TB_63A("
    lgStrSQL = lgStrSQL & " co_cd, "
    lgStrSQL = lgStrSQL & " FISC_YEAR, "
    lgStrSQL = lgStrSQL & " rep_type, "
    lgStrSQL = lgStrSQL & " seq_no, "
    lgStrSQL = lgStrSQL & " W8, "
    lgStrSQL = lgStrSQL & " W9_Fr, "
    lgStrSQL = lgStrSQL & " W9_To, "
    lgStrSQL = lgStrSQL & " W10, "
    lgStrSQL = lgStrSQL & " W11, "
    lgStrSQL = lgStrSQL & " W12, "
    
    lgStrSQL = lgStrSQL & " INSRT_USER_ID, "
    lgStrSQL = lgStrSQL & " INSRT_DT, "
    lgStrSQL = lgStrSQL & " UPDT_USER_ID, "
    lgStrSQL = lgStrSQL & " UPDT_DT)"
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL &   FilterVar(Trim(UCase(lgKeyStream(0))),"","S")      & ","                      
    lgStrSQL = lgStrSQL &   FilterVar(Trim(lgKeyStream(1)),"","S")    & ","           
    lgStrSQL = lgStrSQL &   FilterVar(Trim(lgKeyStream(2)),"","S")      & ","  
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(1),0) & ","
    lgStrSQL = lgStrSQL  &  FilterVar(Trim((arrColVal(2))),"","S") & ","             
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(3),0) & ","
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(4),0) & ","
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(5),0) & ","    
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(6),0) & ","    
    lgStrSQL = lgStrSQL  &  FilterVar(Trim(arrColVal(7)),"","S") & ","   
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                       
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","                                     
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                     
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S")                                            
    lgStrSQL = lgStrSQL & ")"


	PrintLog lgStrSQL
	
	
     '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : 
'============================================================================================================
Sub SubBizSaveSingleUpdate()

    On Error Resume Next
    Err.Clear                                                                        '☜: Clear Error status
   

     lgStrSQL = lgStrSQL & " Update TB_63H set "
     lgStrSQL = lgStrSQL & " w1				= "  &  FilterVar(Trim(UCase(lgKeyStream(3))),"","S")  & ","
     lgStrSQL = lgStrSQL & " w2				= "  &  FilterVar(Trim(UCase(lgKeyStream(4))),"","S")  & ","      
     lgStrSQL = lgStrSQL & " w3				= "  &  FilterVar(Trim(UCase(lgKeyStream(5))),"","S")  & ","      
     lgStrSQL = lgStrSQL & " w4				= "  &  FilterVar(Trim(UCase(lgKeyStream(6))),"","S")  & ","      
     lgStrSQL = lgStrSQL & " w5				= "  &  FilterVar(Trim(UCase(lgKeyStream(7))),"","S")  & ","            
	 
     lgStrSQL = lgStrSQL & " UPDT_USER_ID	= "  & FilterVar(gUsrId,"''","S")                        & ","       
     lgStrSQL = lgStrSQL & " UPDT_DT		= "  & FilterVar(GetSvrDateTime,"","S")           
     lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
	 lgStrSQL = lgStrSQL & "        and fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
	 lgStrSQL = lgStrSQL & "        and rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")   
	 

	PrintLog lgStrSQL
     '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub



'============================================================================================================
' Name : SubBizSaveSingleUpdate2
' Desc : 
'============================================================================================================
Sub SubBizSaveSingleUpdate2()

    On Error Resume Next
    Err.Clear                                                                        '☜: Clear Error status
   

     lgStrSQL = lgStrSQL & " Update TB_63B set "
     lgStrSQL = lgStrSQL & " W7					= "  &  FilterVar(Trim(UCase(lgKeyStream(8))),"","S")  & ","
     lgStrSQL = lgStrSQL & " W13_A				= "  &  FilterVar(Trim(UCase(lgKeyStream(9))),"","S")  & ","      
     lgStrSQL = lgStrSQL & " W13_B				= "  &  FilterVar(Trim(UCase(lgKeyStream(10))),"","S")  & ","      
     lgStrSQL = lgStrSQL & " W13_C				= "  &  FilterVar(Trim(lgKeyStream(11)),"","S")  & ","      
     lgStrSQL = lgStrSQL & " W14_A				= "  &  FilterVar(Trim(UCase(lgKeyStream(12))),"","S")  & ","      
     lgStrSQL = lgStrSQL & " W14_B				= "  &  FilterVar(Trim(UCase(lgKeyStream(13))),"","S")  & ","      
     lgStrSQL = lgStrSQL & " W14_C				= "  &  FilterVar(Trim(lgKeyStream(14)),"","S")  & ","      
     lgStrSQL = lgStrSQL & " W15_A				= "  &  FilterVar(Trim(UCase(lgKeyStream(15))),"","S")  & ","      
     lgStrSQL = lgStrSQL & " W15_B				= "  &  FilterVar(Trim(UCase(lgKeyStream(16))),"","S")  & ","      
     lgStrSQL = lgStrSQL & " W15_C				= "  &  FilterVar(Trim(lgKeyStream(17)),"","S")  & ","      
          
          
          
	 
     lgStrSQL = lgStrSQL & " UPDT_USER_ID	= "  & FilterVar(gUsrId,"''","S")                        & ","       
     lgStrSQL = lgStrSQL & " UPDT_DT		= "  & FilterVar(GetSvrDateTime,"","S")           
     lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
	 lgStrSQL = lgStrSQL & "        and fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
	 lgStrSQL = lgStrSQL & "        and rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")   
	 
	
	
     '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveSingleCreate
' Desc : 
'============================================================================================================
Sub SubBizSaveSingleCreate()
    dim iObjPS5G115
    On Error Resume Next
    Err.Clear                                                                        '☜: Clear Error status
   

     lgStrSQL = "INSERT INTO TB_63H("
     lgStrSQL = lgStrSQL & " co_cd, "         '0
     lgStrSQL = lgStrSQL & " FISC_YEAR, "		'1
     lgStrSQL = lgStrSQL & " rep_type, "		'2
     lgStrSQL = lgStrSQL & " w1  , "		'3
	 lgStrSQL = lgStrSQL & " w2  , "		'4
	 lgStrSQL = lgStrSQL & " w3  , "		'5
	 lgStrSQL = lgStrSQL & " w4  , "		'6
	 lgStrSQL = lgStrSQL & " w5  , "		'7
	 lgStrSQL = lgStrSQL & " INSRT_USER_ID, "
     lgStrSQL = lgStrSQL & " INSRT_DT, "
     lgStrSQL = lgStrSQL & " UPDT_USER_ID, "
     lgStrSQL = lgStrSQL & " UPDT_DT)"
     lgStrSQL = lgStrSQL & " VALUES("
     lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")     & ","                      
     lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")     & ","           
     lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")     & ","   
     lgStrSQL = lgStrSQL &  FilterVar(Trim(UCase(lgKeyStream(3))),"","S")      & ","             
     lgStrSQL = lgStrSQL &  FilterVar(Trim(UCase(lgKeyStream(4))),"","S")    & ","   
          
     lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(5)),"","S")      & ","      
     lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(6))),"","S")      & ","   
     lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(7))),"","S")      & ","   
   
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                       
	lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"","S") & ","                                     
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                     
	lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"","S")                                            
	lgStrSQL = lgStrSQL & ")"


     '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)


End Sub


'============================================================================================================
' Name : SubBizSaveSingleCreate
' Desc : 
'============================================================================================================
Sub SubBizSaveSingleCreate2()
    dim iObjPS5G115
    On Error Resume Next
    Err.Clear                                                                        '☜: Clear Error status
   


    
     lgStrSQL = "INSERT INTO TB_63B("
     lgStrSQL = lgStrSQL & " co_cd, "         '0
     lgStrSQL = lgStrSQL & " FISC_YEAR, "		'1
     lgStrSQL = lgStrSQL & " rep_type, "		'2
     lgStrSQL = lgStrSQL & " W7  , "		'8
	 lgStrSQL = lgStrSQL & " W13_A  , "		'9
	 lgStrSQL = lgStrSQL & " W13_B  , "		'10
	 lgStrSQL = lgStrSQL & " W13_C  , "		'11
	 lgStrSQL = lgStrSQL & " W14_A  , "		'12
	 lgStrSQL = lgStrSQL & " W14_B  , "		'13
	 lgStrSQL = lgStrSQL & " W14_C  , "		'14
	 lgStrSQL = lgStrSQL & " W15_A  , "		'15
	 lgStrSQL = lgStrSQL & " W15_B  , "		'16
	 lgStrSQL = lgStrSQL & " W15_C  , "		'17
	
	 lgStrSQL = lgStrSQL & " INSRT_USER_ID, "
     lgStrSQL = lgStrSQL & " INSRT_DT, "
     lgStrSQL = lgStrSQL & " UPDT_USER_ID, "
     lgStrSQL = lgStrSQL & " UPDT_DT)"
     lgStrSQL = lgStrSQL & " VALUES("
     lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")     & ","                      
     lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")     & ","           
     lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")     & ","   
           
     lgStrSQL = lgStrSQL &  FilterVar(Trim(UCase(lgKeyStream(8))),"","S")    & ","   
     lgStrSQL = lgStrSQL &  FilterVar(Trim(UCase(lgKeyStream(9))),"","S")    & ","    
     lgStrSQL = lgStrSQL &  FilterVar(Trim(UCase(lgKeyStream(10))),"","S")    & ","   
     lgStrSQL = lgStrSQL &  FilterVar(Trim((lgKeyStream(11))),"","S")    & ","   
     lgStrSQL = lgStrSQL &  FilterVar(Trim(UCase(lgKeyStream(12))),"","S")    & ","   
     lgStrSQL = lgStrSQL &  FilterVar(Trim(UCase(lgKeyStream(13))),"","S")    & ","   
     lgStrSQL = lgStrSQL &  FilterVar(Trim((lgKeyStream(14))),"","S")    & ","   
     lgStrSQL = lgStrSQL &  FilterVar(Trim(UCase(lgKeyStream(15))),"","S")    & ","   
     lgStrSQL = lgStrSQL &  FilterVar(Trim(UCase(lgKeyStream(16))),"","S")    & ","   
     lgStrSQL = lgStrSQL &  FilterVar(Trim((lgKeyStream(17))),"","S")    & ","   
   
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                       
	lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"","S") & ","                                     
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                     
	lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"","S")                                            
	lgStrSQL = lgStrSQL & ")"


     '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)


End Sub


'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

On Error Resume Next
Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    lgStrSQL =  " Update TB_63A set"
    lgStrSQL = lgStrSQL & " W8		= " & FilterVar(Trim((arrColVal(2))),"''","S")    & ","  
    lgStrSQL = lgStrSQL & " W9_Fr	= " & UNIConvNum(arrColVal(3),0)   & ","  
    lgStrSQL = lgStrSQL & " W9_To	= " & UNIConvNum(arrColVal(4),0)   & ","  
    lgStrSQL = lgStrSQL & " W10		= " & UNIConvNum(arrColVal(5),0)   & ","  
    lgStrSQL = lgStrSQL & " W11		= " & UNIConvNum(arrColVal(6),0)   & ","  
    lgStrSQL = lgStrSQL & " W12		= " & FilterVar(Trim(arrColVal(7)),"","S")   & ","  

    lgStrSQL = lgStrSQL & " UPDT_USER_ID   = " & FilterVar(gUsrId,"''","S")   & ","  
    lgStrSQL = lgStrSQL & " UPDT_DT   = " &  FilterVar(GetSvrDateTime,"","S")   
    lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
	lgStrSQL = lgStrSQL & "        and fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
	lgStrSQL = lgStrSQL & "        and rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")      
    lgStrSQL = lgStrSQL & "  and SEQ_NO =" & UNIConvNum(arrColVal(1),0) 
    
    
    
    
   
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db

'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "DELETE  TB_63A"
   
    lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
	lgStrSQL = lgStrSQL & "       and fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
	lgStrSQL = lgStrSQL & "        and rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")      
    lgStrSQL = lgStrSQL & "  and SEQ_NO =" & UNIConvNum(arrColVal(1),0) 

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)

	Dim iSelCount
	 
    '------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case Mid(pDataType,1,1)
        Case "S"

	        Select Case  lgPrevNext
                Case " "
                Case "P"
                Case "N"
            End Select
        Case "M"
            iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1
           
            Select Case Mid(pDataType,2,1)
                Case "C"
                
                Case "D"
                       lgStrSQL = "SELECT   SEQ_NO , w8, w9_Fr, w9_To, w10, w11, w12  "
                       lgStrSQL = lgStrSQL & " FROM  TB_63A "
                       lgStrSQL = lgStrSQL & " where "
                       lgStrSQL = lgStrSQL &   pComp & pCode  
                        lgStrSQL = lgStrSQL & "order by  SEQ_NO "
                  
                Case "R"
                       lgStrSQL = "SELECT top 1  "
                        lgStrSQL = lgStrSQL & "w1  , "   
					   lgStrSQL = lgStrSQL & " w2  , "    
				       lgStrSQL = lgStrSQL & " w3  , "   
				       lgStrSQL = lgStrSQL & " w4  , "   
				       lgStrSQL = lgStrSQL & " w5  , "   
				       lgStrSQL = lgStrSQL & " w7  , "   
                       lgStrSQL = lgStrSQL & " w13_A  , "   
					   lgStrSQL = lgStrSQL & " w13_B  , "    
				       lgStrSQL = lgStrSQL & " w13_C  , "   
				       lgStrSQL = lgStrSQL & " w14_A  , "   
					   lgStrSQL = lgStrSQL & " w14_B  , "    
				       lgStrSQL = lgStrSQL & " w14_C  , "     
				       lgStrSQL = lgStrSQL & " w15_A  , "   
					   lgStrSQL = lgStrSQL & " w15_B  , "    
				       lgStrSQL = lgStrSQL & " w15_C   "    
					  
	
                       lgStrSQL = lgStrSQL & " FROM  TB_63H a,TB_63b b"
                       lgStrSQL = lgStrSQL & " where a.co_cd = b.co_cd and a.fisc_year = b.fisc_year and a.rep_type = b.rep_type and"
                       lgStrSQL = lgStrSQL &   pComp & pCode             

	
	
	
                Case "Y"
                      
 
                Case "Z"
                      
                   
                       
				Case "U"
                
            End Select
           
    End Select
	'------ Developer Coding part (End   ) ------------------------------------------------------------------

End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
    
 
                .DBQueryOk
    
                         
	          End With
          End If
        Case "<%=UID_M0002%>"                                                         '☜ : Save
    
                If Trim("<%=lgErrorStatus%>") = "NO" Then
                   
                    Parent.DBSaveOk
                Else
                   
                End If
     
        Case "<%=UID_M0005%>"                                                         '☜ : Save
            If Trim("<%=lgErrorStatus%>") = "NO" Then
                Parent.DBBtnSaveOk
            Else
            End If
        Case "<%=UID_M0006%>"                                                         '☜ : Save
            If Trim("<%=lgErrorStatus%>") = "NO" Then
                Parent.DBBtnSaveOk
            Else
            End If
        Case "<%=UID_M0003%>"                                                         '☜ : Delete
            If Trim("<%=lgErrorStatus%>") = "NO" Then
                Call parent.CancelRestoreToolBar()									<%'버튼 툴바 제어 %>

                Parent.DbDeleteOk
            Else
            End If
    End Select
    
</Script>
