<%@ LANGUAGE=VBSCript CODEPAGE=949 TRANSACTION=Required%>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
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

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)    
    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Dim lgSoSeq
    Dim L1_auto_code
    Dim lgQueryChain
    Dim lgDataError
	Dim iArrTotal
'    ReDim L1_auto_code(lgLngMaxRow)
   
   const W8 =1
   const W9 =2
   const W10 =3
   const W12_1 =4
   const W12_2 =5
   const W13 =6
   
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
             Call SubBizQuerySingle
             if lgErrorStatus <> "YES" then
                 Call SubBizQueryMulti2()
             end if
     
  
        Case CStr(UID_M0002)       
                                               '☜: Save,Update
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

    

Sub SubBizQuerySingle()
'''''''''''''''''''''''''''''''''''''''''''''''''''''
	Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere
   
    Dim iClsRs
    Dim iTemp,i
    Dim k
    
    'On Error Resume Next
    Err.Clear                                                               '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'iKey1 = FilterVar(lgKeyStream(0),"''", "S")
    
    strWhere = " co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
    strWhere = strWhere & "  and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")  
    strWhere = strWhere & "  and rep_type =" & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")  


    Call SubMakeSQLStatements("MR",strWhere,"X","")                              '☜ : Make sql statements
    
    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        iClsRs = 1
        'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write "Call Parent.FncNew"  &  vbCrLf
        Response.Write " </Script>"        
    Else

             
         Response.Write "<Script Language=vbscript>" & vbCr
	     Response.Write "on error resume next" & vbCr
		 Response.Write "With parent" & vbCr
	     Response.Write "	.frm1.txtW2.value       = """ & ConvSPChars(lgObjRs("w2"))       	& """" & vbCr
	     Response.Write "	.frm1.txtW3.value       = """ & ConvSPChars(lgObjRs("w3"))       	& """" & vbCr
	     Response.Write "	.frm1.txtW4.value       = """ & ConvSPChars(lgObjRs("w4"))       	& """" & vbCr
	     Response.Write "	.frm1.txtW5.value      = """ & ConvSPChars(lgObjRs("w5"))       	& """" & vbCr
	     Response.Write "	.frm1.txtW6.value      = """ & ConvSPChars(lgObjRs("w6"))       	& """" & vbCr	

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
    DIM iStrData
    Dim iClsRs
    Dim iTemp,i
    Dim k,Htmp
    
    'On Error Resume Next
    Err.Clear                                                               '☜: Clear Error status
    iDx=0
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'iKey1 = FilterVar(lgKeyStream(0),"''", "S")
    
    strWhere = " co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
    strWhere = strWhere & "  and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")  
    strWhere = strWhere & "  and rep_type =" & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")  

    Call SubMakeSQLStatements("MD",strWhere,"X","")                              '☜ : Make sql statements
    
    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        iClsRs = 1
      '  Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
       Call SetErrorStatus()
    Else
  		
        Do While Not lgObjRs.EOF
		
            if iDx=0 then 				Htmp = split(lgObjRs("W8"),"::")

			iStrData = iStrData & Chr(11)	
			iStrData = iStrData & Chr(11)							
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W9"))
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W10"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W12_1"))	
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W12_2"))		
			iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W13"))	
			iStrData = iStrData & Chr(11) & iDx
			iStrData = iStrData & Chr(11) & Chr(12)
	
		    lgObjRs.MoveNext
		    
		    iDx = iDx + 1

        Loop 
		lgObjRs.Close
        Set lgObjRs = Nothing
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
	    Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr
	    Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
	    Response.Write "	.DbQueryOk                                      " & vbCr
	    Response.Write " End With                                           " & vbCr
	    Response.Write " </Script>                                          " & vbCr

                  Response.Write "<Script Language=vbscript>" & vbCr
				 Response.Write "With parent.frm1.vspdData" & vbCr
				 'Response.Write "	.maxrows = 9+4  " & vbCr 
				 Response.Write "		.Row = 1 " & vbCr 
				 Response.Write "		.Col = 3	: .CellType = 1	 " & vbCr 
				 Response.Write "		.text = "" "& Htmp(1)&" "" " & vbCr 
				 Response.Write "		.Row = 1 " & vbCr 
				 Response.Write "		.Col = 4	: .CellType = 1	" & vbCr 
				 Response.Write "		.text = "" "& Htmp(2)&" "" " & vbCr 
				 Response.Write "		.Row = 1 " & vbCr 
				 Response.Write "		.Col = 5	: .CellType = 1	 " & vbCr 
				 Response.Write "		.text = "" "& Htmp(3)&" "" " & vbCr 
				 Response.Write "		.Row = 1 " & vbCr 
				 Response.Write "		.Col = 6	: .CellType = 1	  " & vbCr 
				 Response.Write "		.text = "" "& Htmp(4)&" "" " & vbCr 
				 Response.Write "		 Row = 1 " & vbCr 
				 Response.Write "		.Col = 7	: .CellType = 1	: " & vbCr 
				 Response.Write "		.text = """" " & vbCr 
				 Response.Write " End With "	& vbCr
				 Response.Write "</Script>"  & vbCr

          
       
    If iDx <= lgMaxCount Then
       lgStrPrevKeyIndex = ""
    End If   
            
    

  end if  

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs) 
End Sub



Sub SubBizAutoQuery()

    Dim strWhere
   
   ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
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
				              strWhere = " co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
							 strWhere = strWhere & "  and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")  
							 strWhere = strWhere & "  and rep_type =" & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")  

    
							Call SubMakeSQLStatements("MR",strWhere,"X","")                                                               '☜: Create
				          If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
							    lgStrPrevKeyIndex = ""
							 
							    
							    Call DisplayMsgbox("WC0001","X", FilterVar(Trim(UCase(lgKeyStream(1))),"","S")   ,"X" ,I_MKSCRIPT)  
		
							    Call SetErrorStatus()
							    lgErrorStatus = "YES"
							   
							Else
                                Call SubBizSaveSingleCreate()
									             	
							end if	
		
				          
				    Case  OPMD_UMODE  
				                                                                  '☜: Update
							  Call SubBizSaveSingleUpdate()
				    
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
        lgStrSQL = "delete from TB_23BD " 
        lgStrSQL = lgStrSQL & " WHERE co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
		lgStrSQL = lgStrSQL & "       and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")  
		lgStrSQL = lgStrSQL & "        and rep_type =" & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")  
		lgStrSQL = lgStrSQL & " delete from TB_23BH " 
	     lgStrSQL = lgStrSQL & " WHERE co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
		lgStrSQL = lgStrSQL & "       and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")  
		lgStrSQL = lgStrSQL & "        and rep_type =" & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")  
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
	
    strWhere = " co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
	strWhere = strWhere & "  and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")  
	 strWhere = strWhere & "  and rep_type =" & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")  

    
    Call SubMakeSQLStatements("MR",strWhere,"X","")                              '☜ : Make sql statements
    
    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = false  and lgIntFlgMode =OPMD_CMODE Then
        lgStrPrevKeyIndex = ""
        
        Call DisplayMsgbox("WC0001","X",Trim(UCase(Request("txtFISC_YEAR"))) ,"X" ,I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
        lgQueryChain = 0
    Else
        lgQueryChain = 1

            
	        arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data

            For iDx = 1 To uBound(arrRowVal,1) 
                arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
                
                Select Case arrColVal(0)
                    Case "C" ,"U"
                        Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
               
                                             '☜: Update
                    Case "D"
                                 '☜: Delete
                End Select
                
                If lgErrorStatus    = "YES" Then
                   lgErrorPos = lgErrorPos & arrColVal(W8) & gColSep
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
  
    dim Htmp

    if arrColVal(8)=1 then
		
		arrColVal(W8) =arrColVal(W8) & "::" & arrColVal(W9) &"::"&  arrColVal(W10) &"::"&  arrColVal(W12_1) &"::"&  arrColVal(W12_2)
		arrColVal(W9)=0
		arrColVal(W10)=0
		arrColVal(W12_1)=0
		arrColVal(W12_2)=0
		
	end if
    
    lgStrSQL =  " if not EXISTS (select * from  TB_23BD "
    lgStrSQL = lgStrSQL & "		WHERE	co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
	lgStrSQL = lgStrSQL & "				and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")  
	lgStrSQL = lgStrSQL & "				and rep_type =" & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")  
	lgStrSQL = lgStrSQL & "				and SEQ_NO =" & UNIConvNum(arrColVal(8),0)
    lgStrSQL = lgStrSQL & "     )"
    lgStrSQL = lgStrSQL & " Begin " & chr(13)
    lgStrSQL = lgStrSQL & "	INSERT INTO TB_23BD(" & chr(13)
    lgStrSQL = lgStrSQL & " co_cd, " & chr(13)
    lgStrSQL = lgStrSQL & " FISC_YEAR, " & chr(13)
    lgStrSQL = lgStrSQL & " rep_type, " & chr(13)
    lgStrSQL = lgStrSQL & " SEQ_NO, " & chr(13)
    lgStrSQL = lgStrSQL & " W8, " & chr(13)
    lgStrSQL = lgStrSQL & " W9, " & chr(13)
    lgStrSQL = lgStrSQL & " W10, " & chr(13)
    lgStrSQL = lgStrSQL & " W12_1, " & chr(13)
    lgStrSQL = lgStrSQL & " W12_2, " & chr(13)
	lgStrSQL = lgStrSQL & " W13, " & chr(13)
    lgStrSQL = lgStrSQL & " INSRT_USER_ID, " & chr(13)
    lgStrSQL = lgStrSQL & " INSRT_DT, " & chr(13)
    lgStrSQL = lgStrSQL & " UPDT_USER_ID, " & chr(13)
    lgStrSQL = lgStrSQL & " UPDT_DT)" & chr(13)
    lgStrSQL = lgStrSQL & " VALUES(" & chr(13)
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")      & ","                      
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")       & ","           
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")      & ","      
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(8),0) & "," 
    lgStrSQL = lgStrSQL  &  FilterVar(Trim(UCase(arrColVal(W8))),"","S")  & ","           
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(W9),0) & ","    
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(W10),0) & ","    
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(W12_1),0) & ","    
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(W12_2),0) & ","     & chr(13)
	lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(W13),0) & ","     & chr(13)
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                       
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","                                     
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                     
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S")                                            
    lgStrSQL = lgStrSQL & ")"
    lgStrSQL = lgStrSQL & " END "
    lgStrSQL = lgStrSQL & " ELSE "
    lgStrSQL = lgStrSQL & " BEGIN "
     lgStrSQL = lgStrSQL & "Update TB_23BD set" & chr(13)

    lgStrSQL = lgStrSQL & " W8		= " & FilterVar(arrColVal(W8),"","S")    & "," 
    lgStrSQL = lgStrSQL & " W9		= " & UNIConvNum(arrColVal(W9),0)   & ","  
    lgStrSQL = lgStrSQL & " W10		= " & UNIConvNum(arrColVal(W10),0)   & ","   & chr(13)
    lgStrSQL = lgStrSQL & " W12_1	= " & UNIConvNum(arrColVal(W12_1),0)   & ","  
    lgStrSQL = lgStrSQL & " W12_2	= " & UNIConvNum(arrColVal(W12_2),0)   & ","   & chr(13)
	lgStrSQL = lgStrSQL & " W13	= " & UNIConvNum(arrColVal(W13),0)   & ","   & chr(13)
    lgStrSQL = lgStrSQL & " UPDT_USER_ID   = " & FilterVar(gUsrId,"''","S")   & ","  
    lgStrSQL = lgStrSQL & " UPDT_DT   = " &  FilterVar(GetSvrDateTime,"","S")    & chr(13)
    lgStrSQL = lgStrSQL & " WHERE co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
	lgStrSQL = lgStrSQL & "       and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")  & chr(13) 
	lgStrSQL = lgStrSQL & "        and rep_type =" & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")  
     '---------- Developer Coding part (End  ) --------------------------------------------------------
    lgStrSQL = lgStrSQL & "  and SEQ_NO =" &  UNIConvNum(arrColVal(8),0) & chr(13)
    lgStrSQL = lgStrSQL & " END "

     '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveSingleCreate
' Desc : 
'============================================================================================================
Sub SubBizSaveSingleUpdate()

    On Error Resume Next
    Err.Clear                                                                        '☜: Clear Error status
   


    
    lgStrSQL = lgStrSQL & " Update TB_23BH set"

    lgStrSQL = lgStrSQL & " W1  = " & UNIConvNum(Request("txtW1"),0)   & ","  
    lgStrSQL = lgStrSQL & " W2  = " & UNIConvNum(Request("txtW2"),0)   & ","  
    lgStrSQL = lgStrSQL & " W3  = " & UNIConvNum(Request("txtW3"),0)   & ","  
    lgStrSQL = lgStrSQL & " W4		 = " & UNIConvNum(Request("txtW4"),0) & ","  
    lgStrSQL = lgStrSQL & " W5		 = " & UNIConvNum(Request("txtW5"),0)  & ","  	
    lgStrSQL = lgStrSQL & " W6		 = " & UNIConvNum(Request("txtW6"),0)   & ","  
    lgStrSQL = lgStrSQL & " UPDT_USER_ID   = " & FilterVar(gUsrId,"''","S")   & ","  
    lgStrSQL = lgStrSQL & " UPDT_DT   = " &  FilterVar(GetSvrDateTime,"","S")   
    lgStrSQL = lgStrSQL & " WHERE co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
	lgStrSQL = lgStrSQL & "       and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")  
	lgStrSQL = lgStrSQL & "        and rep_type =" & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")  
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
   


    
     lgStrSQL = lgStrSQL & "INSERT INTO TB_23BH("
    lgStrSQL = lgStrSQL & " co_cd, "
    lgStrSQL = lgStrSQL & " FISC_YEAR, "
    lgStrSQL = lgStrSQL & " rep_type, "

    lgStrSQL = lgStrSQL & " W1, "
    lgStrSQL = lgStrSQL & " W2, "
    lgStrSQL = lgStrSQL & " W3, "
    lgStrSQL = lgStrSQL & " W4, "
    lgStrSQL = lgStrSQL & " W5, "
    lgStrSQL = lgStrSQL & " W6, "
   
    lgStrSQL = lgStrSQL & " INSRT_USER_ID, "
    lgStrSQL = lgStrSQL & " INSRT_DT, "
    lgStrSQL = lgStrSQL & " UPDT_USER_ID, "
    lgStrSQL = lgStrSQL & " UPDT_DT)"
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")      & ","                      
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")       & ","           
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")      & ","             
   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW1"),0)     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW2"),0)     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW3"),0)     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW4"),0)     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW5"),0)     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW6"),0)     & ","   

                                          
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
    
    lgStrSQL =  " Update TB_23BD set"


    lgStrSQL = lgStrSQL & " W9		= " & UNIConvNum(arrColVal(W9),0)   & ","  
    lgStrSQL = lgStrSQL & " W10		= " & UNIConvNum(arrColVal(W10),0)   & ","  
    lgStrSQL = lgStrSQL & " W12_1	= " & UNIConvNum(arrColVal(W12_1),0)   & ","  
    lgStrSQL = lgStrSQL & " W12_2	= " & UNIConvNum(arrColVal(W12_2),0)   & ","  
    lgStrSQL = lgStrSQL & " W13		= " & UNIConvNum(arrColVal(W13),0)   & ","  
    lgStrSQL = lgStrSQL & " W14		= " & UNIConvNum(arrColVal(7),0)   & ","  
    lgStrSQL = lgStrSQL & " W15		= " & UNIConvNum(arrColVal(8),0)   & ","  
    lgStrSQL = lgStrSQL & " UPDT_USER_ID   = " & FilterVar(gUsrId,"''","S")   & ","  
    lgStrSQL = lgStrSQL & " UPDT_DT   = " &  FilterVar(GetSvrDateTime,"","S")   
    lgStrSQL = lgStrSQL & " WHERE co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
	lgStrSQL = lgStrSQL & "       and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")  
	lgStrSQL = lgStrSQL & "        and rep_type =" & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")  
     '---------- Developer Coding part (End  ) --------------------------------------------------------
    lgStrSQL = lgStrSQL & "  and W8 =" &  FilterVar(Trim(UCase(arrColVal(W8))),"''","S")  
   
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
                       lgStrSQL = "SELECT    w8,		w9,		w10,	w12_1	,w12_2		,w13,	w14,	w15 "
                       lgStrSQL = lgStrSQL & " FROM  TB_23BD "
                       lgStrSQL = lgStrSQL & " where "
                       lgStrSQL = lgStrSQL &   pComp & pCode  
                       lgStrSQL = lgStrSQL & "order by  seq_NO "
              
                Case "R"
                       lgStrSQL = "SELECT top 1  w2, w3, w4, w5, w6"
                       lgStrSQL = lgStrSQL & " FROM  TB_23BH "
                       lgStrSQL = lgStrSQL & " where "
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
                Parent.DbDeleteOk
            Else
            End If
    End Select
    
</Script>
