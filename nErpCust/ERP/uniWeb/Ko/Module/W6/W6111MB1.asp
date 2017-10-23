<%@ LANGUAGE="VBScript" CODEPAGE=949 TRANSACTION=Required%>

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
    

    strWhere = " co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
	strWhere = strWhere & "  and fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
	strWhere = strWhere & "  and rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")


    Call SubMakeSQLStatements("MR",strWhere,"X","")                              '☜ : Make sql statements
    
    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        iClsRs = 1
        'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else
  
         Response.Write "<Script Language=vbscript>" & vbCr
	     Response.Write "With parent.frm1.vspdData1" & vbCr
	 

		 Response.Write "   parent.ggoSpread.Source = parent.frm1.vspdData1 "  & vbCr
		 Response.Write "	 .row = 1  : .col = 3 :   .text = """ & ConvSPChars(lgObjRs("w8_d1_s"))       	& """" & vbCr 
		 Response.Write "	 .row = 1  : .col = 5 :   .text = """ & ConvSPChars(lgObjRs("w8_d1_e"))       	& """" & vbCr 
		 Response.Write "	 .row = 1  : .col = 6 :	  .text = """ & ConvSPChars(lgObjRs("w8_d2_s"))       	& """" & vbCr 
		 Response.Write "	 .row = 1  : .col = 8 :	  .text = """ & ConvSPChars(lgObjRs("w8_d2_e"))       	& """" & vbCr 
		 Response.Write "    .row = 1  : .col = 9 :   .text = """ & ConvSPChars(lgObjRs("w8_d3_s"))       	& """" & vbCr 
		 Response.Write "    .row = 1  : .col = 11 :  .text = """ & ConvSPChars(lgObjRs("w8_d3_e"))       	& """" & vbCr
		 Response.Write "    .row = 1  : .col = 12 :  .text = """ & ConvSPChars(lgObjRs("w8_d4_s"))       	& """" & vbCr
		 Response.Write "    .row = 1  : .col = 14 :  .text = """ & ConvSPChars(lgObjRs("w8_d4_e"))       	& """" & vbCr
					
		 Response.Write "    .row = 2  : .col = 3 :  .text = " & UNIConvNum(lgObjRs("w8_Amt1"),0)       	& "" & vbCr 
		 Response.Write "    .row = 2  : .col = 6 :  .text = " & UNIConvNum(lgObjRs("w8_Amt2"),0)       	& "" & vbCr 
		 Response.Write "    .row = 2  : .col = 9 :  .text = " & UNIConvNum(lgObjRs("w8_Amt3"),0)       	& "" & vbCr 
		 Response.Write "    .row = 2  : .col = 12 : .text = " & UNIConvNum(lgObjRs("w8_Amt4"),0)       	& "" & vbCr 
		 Response.Write "    .row = 2  : .col = 15 : .text = " & UNIConvNum(lgObjRs("w8_Sum"),0)       	& "" & vbCr 
					
		 Response.Write "    .row = 4  : .col = 1 :  .text = " & UNIConvNum(lgObjRs("w9"),0)       	& "" & vbCr 
  		 Response.Write "    .row = 5  : .col = 3 :  .text = " & UNIConvNum(lgObjRs("w10"),0)       	& "" & vbCr 
  		 Response.Write "    .row = 8  : .col = 2 :   .text = " & UNIConvNum(lgObjRs("w15_11"),0)       	& "" & vbCr 
  		 Response.Write "    .row = 8  : .col = 6 :   .text = """ & ConvSPChars(lgObjRs("w15_12_View"))       	& """" & vbCr 
  		 Response.Write "     parent.frm1.txt15_12Value.value =  """ & ConvSPChars(lgObjRs("w15_12_Value")) & """" & vbCr 
  		 Response.Write "     .row = 8  : .col = 9 : .text =  " &  UNIConvNum(lgObjRs("w15_13"),0)     	& "" & vbCr 
  		 Response.Write "     .row = 8  : .col = 12 : .text =  """ &   ConvSPChars(lgObjRs("w15_14"))   	& """" & vbCr   
  		 Response.Write "     .row = 9  : .col =  2 :  .text =  " &  UNIConvNum(lgObjRs("w16_11"),0)    & "" & vbCr 
  		 Response.Write "     .row = 9  : .col =6 :  .text =  """ &   ConvSPChars(lgObjRs("w16_12_View"))   	& """" & vbCr   
  		 Response.Write "     parent.frm1.txt16_12Value.value = """ & ConvSPChars(lgObjRs("w16_12_Value")) & """" & vbCr 
  		 Response.Write "    .row = 9  : .col = 9 : .text =  " &  UNIConvNum(lgObjRs("w16_13"),0)     	 & "" & vbCr 
  		 Response.Write "    .row = 9  : .col = 12 : .text =  """ &   ConvSPChars(lgObjRs("w16_14"))   	& """" & vbCr   
        ' Response.Write "  parent.frm1.txtCompType.value     =  """ & ConvSPChars(lgObjRs("COMP_TYPE1")) & """" & vbCr 
  		 Response.Write "  parent.frm1.txt4yearMth.value	=  " &  UNIConvNum(lgObjRs("W_CURR_YEAR_Mth"),0)     	 & "" & vbCr 	
  		 Response.Write "  parent.frm1.txtyearMth.value		=  " &  UNIConvNum(lgObjRs("W_4YEAR_Mth"),0)     	 & "" & vbCr   
  	 if  ConvSPChars(lgObjRs("COMP_TYPE1")) = 2 then   '중소기업 
  				       
  	     Response.Write "  .row = 10 : .col = 9 : .text =  """ &  UNIConvNum(lgObjRs("w17"),0)     	& """" & vbCr  
  	     Response.Write "  .row = 10  : .col = 12 :  .text =  """ &  ConvSPChars(lgObjRs("w17_14"))     	& """" & vbCr   
  				      
  	 else
  				       
  	     Response.Write "  .row = 11 : .col = 9 : .text =  """ &  UNIConvNum(lgObjRs("w17"),0)     	& """" & vbCr  
  	     Response.Write "  .row = 11  : .col = 12 :  .text =  """ &  ConvSPChars(lgObjRs("w17_14"))     	& """" & vbCr   
  	 end if
					
		 Response.Write "  .row = 12  : .col = 9  :  .text =  """ &  UNIConvNum(lgObjRs("w18_A13"),0)     	& """" & vbCr  
  	     Response.Write "  .row = 12  : .col = 12 :  .text =  """ &  ConvSPChars(lgObjRs("w18_A14"))     	& """" & vbCr   
  	     Response.Write "  .row = 13  : .col = 9  :  .text =  """ &  UNIConvNum(lgObjRs("w18_B13"),0)     	& """" & vbCr  
  	     Response.Write "  .row = 13  : .col = 12 :  .text =  """ &  ConvSPChars(lgObjRs("w18_B14"))     	& """" & vbCr   
  	     Response.Write "  .row = 14  : .col = 9  :  .text =  """ &  UNIConvNum(lgObjRs("w18_C13"),0)     	& """" & vbCr  
  	     Response.Write "  .row = 14  : .col = 12 :  .text =  """ &  ConvSPChars(lgObjRs("w18_C14"))     	& """" & vbCr  
  	     Response.Write "  .row = 4   : .col = 2 : .TypeHAlign = 2 :   .text =  """ &  ConvSPChars(lgObjRs("w_DESC"))     	& """" & vbCr   
  	    
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
                lgstrData = lgstrData & Chr(11) & ""
                lgstrData = lgstrData & Chr(11) & "계정"
				lgstrData = lgstrData & Chr(11) & ""
            	lgstrData = lgstrData & Chr(11) & ""
            	lgstrData = lgstrData & Chr(11) & ""
            	lgstrData = lgstrData & Chr(11) & ""
            	lgstrData = lgstrData & Chr(11) & ""
            	lgstrData = lgstrData & Chr(11) & ""
            	lgstrData = lgstrData & Chr(11) & ""
            	lgstrData = lgstrData & Chr(11) & ""
            	lgstrData = lgstrData & Chr(11) & ""
            	lgstrData = lgstrData & Chr(11) & ""
            	lgstrData = lgstrData & Chr(11) & iDx
                lgstrData = lgstrData & Chr(11) & Chr(12) 	
        
        Do While Not lgObjRs.EOF
           
            
        
		
                lgstrData = lgstrData & Chr(11) & UNIConvNum(lgObjRs("seq_no"),0)
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_NM"))
				lgstrData = lgstrData & Chr(11) & UNIConvNum(lgObjRs("w1"),0)
            	lgstrData = lgstrData & Chr(11) & UNIConvNum(lgObjRs("w2"),0)
            	lgstrData = lgstrData & Chr(11) & UNIConvNum(lgObjRs("w3"),0)
            	lgstrData = lgstrData & Chr(11) & UNIConvNum(lgObjRs("w4"),0)
            	lgstrData = lgstrData & Chr(11) & UNIConvNum(lgObjRs("w5"),0)
            	lgstrData = lgstrData & Chr(11) & UNIConvNum(lgObjRs("w6"),0)
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("w1_t"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("w2_t"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("w3_t"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("w4_t"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("w5_t"))
 
	
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

    Dim strWhere
   
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
  
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
             	
		strWhere = " co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
		strWhere = strWhere & "  and fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
		strWhere = strWhere & "  and rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")
    


    
    
    Call SubMakeSQLStatements("MK",strWhere,"X",C_EQ)                                 '☆ : Make sql statements
  
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
     
		lgStrPrevKeyIndex = ""
        iClsRs = 1
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
   ELSE

         Response.Write "<Script Language=vbscript>" & vbCr
	     Response.Write "With parent" & vbCr
	   
	     Response.Write "	.frm1.txtW4.text       = """ & UNIConvNum(lgObjRs("A_TYPE_AMT"),0)       	& """" & vbCr
	  
	     Response.Write " End With "	& vbCr
         Response.Write "</Script>"  & vbCr
	     

              
End if
	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
	
    Call SubCloseRs(lgObjRs)    
    Call SubCloseCommandObject(lgObjComm)                                                      '☜: Release RecordSSet

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
							strWhere = strWhere & "  and fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
							strWhere = strWhere & "  and rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")

    
							Call SubMakeSQLStatements("MR",strWhere,"X","")                                                               '☜: Create
				          If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
							    lgStrPrevKeyIndex = ""
							    
							    Call DisplayMsgbox("WC0001","X",Trim(UCase(Request("txtFISC_YEAR"))) ,"X" ,I_MKSCRIPT)      '☜ : No data is found. 
							    
							    
		
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
        lgStrSQL = "delete from TB_JT3B " 
		lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
		lgStrSQL = lgStrSQL & "  and fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
		lgStrSQL = lgStrSQL & "  and rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")
		lgStrSQL = lgStrSQL & " delete from TB_JT3A " 
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
	
            	
	strWhere = " co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
	strWhere = strWhere & "  and fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
	strWhere = strWhere & "  and rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")
    
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
   


    
    
    lgStrSQL = "INSERT INTO TB_JT3B("
    lgStrSQL = lgStrSQL & " co_cd, "
    lgStrSQL = lgStrSQL & " FISC_YEAR, "
    lgStrSQL = lgStrSQL & " rep_type, "
    lgStrSQL = lgStrSQL & " seq_no, "
    lgStrSQL = lgStrSQL & " ACCT_NM, "
    lgStrSQL = lgStrSQL & " W1, "
    lgStrSQL = lgStrSQL & " W2, "
    lgStrSQL = lgStrSQL & " W3, "
    lgStrSQL = lgStrSQL & " W4, "
    lgStrSQL = lgStrSQL & " W5, "
    lgStrSQL = lgStrSQL & " W6, "
    lgStrSQL = lgStrSQL & " W1_T, "
    lgStrSQL = lgStrSQL & " W2_T, "
    lgStrSQL = lgStrSQL & " W3_T, "
    lgStrSQL = lgStrSQL & " W4_T, "
    lgStrSQL = lgStrSQL & " W5_T, "
    
    lgStrSQL = lgStrSQL & " INSRT_USER_ID, "
    lgStrSQL = lgStrSQL & " INSRT_DT, "
    lgStrSQL = lgStrSQL & " UPDT_USER_ID, "
    lgStrSQL = lgStrSQL & " UPDT_DT)"
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S")     & ","                      
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S")     & ","           
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S")     & ","  
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(1),0) & ","             
    lgStrSQL = lgStrSQL  &  FilterVar(Trim(UCase(arrColVal(2))),"","S") & ","    

    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(3),0) & ","    
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(4),0) & ","    
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(5),0) & ","    
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(6),0) & ","    
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(7),0) & ","
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(8),0) & ","      
    lgStrSQL = lgStrSQL  &  FilterVar(Trim(UCase(arrColVal(9))),"","S") & ","   
    lgStrSQL = lgStrSQL  &  FilterVar(Trim(UCase(arrColVal(10))),"","S") & ","   
    lgStrSQL = lgStrSQL  &  FilterVar(Trim(UCase(arrColVal(11))),"","S") & ","   
    lgStrSQL = lgStrSQL  &  FilterVar(Trim(UCase(arrColVal(12))),"","S") & ","   
    lgStrSQL = lgStrSQL  &  FilterVar(Trim(UCase(arrColVal(13))),"","S") & ","   
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                       
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","                                     
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                     
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S")                                            
    lgStrSQL = lgStrSQL & ")"


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
   

     lgStrSQL = lgStrSQL & " Update TB_JT3A set "
     lgStrSQL = lgStrSQL & " co_cd			= "  &  FilterVar(Trim(UCase(lgKeyStream(0))),"","S")     & ","                      
     lgStrSQL = lgStrSQL & " FISC_YEAR		= "  &  FilterVar(Trim(UCase(lgKeyStream(1))),"","S")     & ","        
     lgStrSQL = lgStrSQL & " rep_type		= "  &  FilterVar(Trim(UCase(lgKeyStream(2))),"","S")     & ","       
     lgStrSQL = lgStrSQL & " w8_d1_s		= "  &  FilterVar(Trim(lgKeyStream(3)),"","S")      & ","      
	 lgStrSQL = lgStrSQL & " w8_d1_e		= "  &  FilterVar(Trim(lgKeyStream(4)),"","S")       & ","    
	 lgStrSQL = lgStrSQL & " w8_d2_s		= "  &  FilterVar(Trim(lgKeyStream(5)),"","S")      & ","    
	 lgStrSQL = lgStrSQL & " w8_d2_e		= "  &  FilterVar(Trim(lgKeyStream(6)),"","S")      & "," 
	 lgStrSQL = lgStrSQL & " w8_d3_s		= "  &  FilterVar(Trim(lgKeyStream(7)),"","S")      & "," 
	 lgStrSQL = lgStrSQL & " w8_d3_e		= "  &  FilterVar(Trim(lgKeyStream(8)),"","S")      & "," 
	 lgStrSQL = lgStrSQL & " w8_d4_s		= "  &  FilterVar(Trim(lgKeyStream(9)),"","S")      & "," 
	 lgStrSQL = lgStrSQL & " w8_d4_e		= "  &  FilterVar(Trim(lgKeyStream(10)),"","S")      & "," 
	
	 lgStrSQL = lgStrSQL & " w8_Amt1		= "  &  UNIConvNum(Trim(lgKeyStream(11)),0)     & ","   
	 lgStrSQL = lgStrSQL & " w8_Amt2		= "  &  UNIConvNum(Trim(lgKeyStream(12)),0)     & "," 
	 lgStrSQL = lgStrSQL & " w8_Amt3		= "  &  UNIConvNum(Trim(lgKeyStream(13)),0)     & ","    
	 lgStrSQL = lgStrSQL & " w8_Amt4		= "  &  UNIConvNum(Trim(lgKeyStream(14)),0)     & ","   
	 lgStrSQL = lgStrSQL & " w8_Sum			= "  &  UNIConvNum(Trim(lgKeyStream(15)),0)     & "," 
	 
	 lgStrSQL = lgStrSQL & " w9				= "  &  UNIConvNum(Trim(lgKeyStream(16)),0)     & "," 
	 lgStrSQL = lgStrSQL & " w10			= "  &  UNIConvNum(Trim(lgKeyStream(17)),0)     & "," 
		
	 lgStrSQL = lgStrSQL & " w15_11			= "  &  UNIConvNum(Trim(lgKeyStream(18)),0)     & "," 
	 lgStrSQL = lgStrSQL & " w15_12_View	= "  &  FilterVar(Trim(lgKeyStream(19)),"","S")      & ","    
	 lgStrSQL = lgStrSQL & " w15_12_Value	= "  & 	UNIConvNum(Trim(lgKeyStream(20)),0)     & ","
	 lgStrSQL = lgStrSQL & " w15_13			= "  &  UNIConvNum(Trim(lgKeyStream(21)),0)     & "," 
	 lgStrSQL = lgStrSQL & " w15_14			= "  &  FilterVar(Trim(lgKeyStream(22)),"","S")      & ","    
		
		
	 lgStrSQL = lgStrSQL & " w16_11			= "  & UNIConvNum(Trim(lgKeyStream(23)),0)     & "," 
	 lgStrSQL = lgStrSQL & " w16_12_View	= "  & FilterVar(Trim(lgKeyStream(24)),"","S")      & ","    
	 lgStrSQL = lgStrSQL & " w16_12_Value	= "  & UNIConvNum(Trim(lgKeyStream(25)),0)     & ","
	 lgStrSQL = lgStrSQL & " w16_13			= "  & UNIConvNum(Trim(lgKeyStream(26)),0)     & ","     
	 lgStrSQL = lgStrSQL & " w16_14			= "  & FilterVar(Trim(lgKeyStream(27)),"","S")      & ","    
	 
	 lgStrSQL = lgStrSQL & " COMP_TYPE1		= "  &  FilterVar(Trim(lgKeyStream(28)),"","S")      & ","    
	 lgStrSQL = lgStrSQL & " w17			= "  & UNIConvNum(Trim(lgKeyStream(29)),0)     & ","    
	 lgStrSQL = lgStrSQL & " w17_14			= "  & FilterVar(Trim(lgKeyStream(30)),"","S")      & ","  
	 lgStrSQL = lgStrSQL & " w18_A13		= "  & UNIConvNum(Trim(lgKeyStream(31)),0)     & ","    
	 lgStrSQL = lgStrSQL & " w18_A14		= "  & FilterVar(Trim(lgKeyStream(32)),"","S")      & "," 
     lgStrSQL = lgStrSQL & " w18_B13		= "  & UNIConvNum(Trim(lgKeyStream(33)),0)     & ","    
	 lgStrSQL = lgStrSQL & " w18_B14		= "  & FilterVar(Trim(lgKeyStream(34)),"","S")      & "," 
	 lgStrSQL = lgStrSQL & " w18_C13		= "  & UNIConvNum(Trim(lgKeyStream(35)),0)     & ","    
	 lgStrSQL = lgStrSQL & " w18_C14		= "  & FilterVar(Trim(lgKeyStream(36)),"","S")      & "," 
	 lgStrSQL = lgStrSQL & " w_DESC			= "  & FilterVar(Trim(lgKeyStream(37)),"","S")      & "," 
	 lgStrSQL = lgStrSQL & " W_CURR_YEAR_Mth			= "  & UNIConvNum(Trim(lgKeyStream(38)),0)     & "," 
	 lgStrSQL = lgStrSQL & " W_4YEAR_Mth			= "  & UNIConvNum(Trim(lgKeyStream(39)),0)     & "," 
	 
 
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
   


    
     lgStrSQL = "INSERT INTO TB_JT3A("
     lgStrSQL = lgStrSQL & " co_cd, "         '0
     lgStrSQL = lgStrSQL & " FISC_YEAR, "		'1
     lgStrSQL = lgStrSQL & " rep_type, "		'2
     lgStrSQL = lgStrSQL & " w8_d1_s  , "		'3
	 lgStrSQL = lgStrSQL & " w8_d1_e  , "		'4
	 lgStrSQL = lgStrSQL & " w8_d2_s  , "		'5
	 lgStrSQL = lgStrSQL & " w8_d2_e  , "		'6
	 lgStrSQL = lgStrSQL & " w8_d3_s  , "		'7
	 lgStrSQL = lgStrSQL & " w8_d3_e  , "		'8
	 lgStrSQL = lgStrSQL & " w8_d4_s  , "		'9
	 lgStrSQL = lgStrSQL & " w8_d4_e  , "		'10
	 
	 lgStrSQL = lgStrSQL & " w8_Amt1  , "		'11
	 lgStrSQL = lgStrSQL & " w8_Amt2  , "		'12
	 lgStrSQL = lgStrSQL & " w8_Amt3  , "		'13
	 lgStrSQL = lgStrSQL & " w8_Amt4  , "		'14
	 lgStrSQL = lgStrSQL & " w8_Sum   , "		'15
	 
	 lgStrSQL = lgStrSQL & " w9      , "		'16
	 lgStrSQL = lgStrSQL & " w10     , "		'17
	 
	 lgStrSQL = lgStrSQL & " w15_11  , "		'18
	 lgStrSQL = lgStrSQL & " w15_12_View  , "   '19 
	 lgStrSQL = lgStrSQL & " w15_12_Value  , "  '20 
	 lgStrSQL = lgStrSQL & " w15_13    , "		'21
	 lgStrSQL = lgStrSQL & " w15_14    , "		'22
		
		
	 lgStrSQL = lgStrSQL & " w16_11     , "     '23
	 lgStrSQL = lgStrSQL & " w16_12_View   , "  '24 
	 lgStrSQL = lgStrSQL & " w16_12_Value  , "  '25 
	 lgStrSQL = lgStrSQL & " w16_13     , "		'26
	 lgStrSQL = lgStrSQL & " w16_14		,"	'27
	 
	 lgStrSQL = lgStrSQL & " COMP_TYPE1  , "	'28
	 lgStrSQL = lgStrSQL & "  w17         , "	'29
	 lgStrSQL = lgStrSQL & "  w17_14      , "		'30
	 lgStrSQL = lgStrSQL & "  w18_A13         , "	'31
	 lgStrSQL = lgStrSQL & "  w18_A14      , "		'32
	 lgStrSQL = lgStrSQL & "  w18_B13         , "	'33
	 lgStrSQL = lgStrSQL & "  w18_B14      , "		'34
     lgStrSQL = lgStrSQL & "  w18_C13         , "	'35
	 lgStrSQL = lgStrSQL & "  w18_C14      , "		'36
     lgStrSQL = lgStrSQL & "  w_DESC      , "		'37
     lgStrSQL = lgStrSQL & "  W_CURR_YEAR_Mth      , "		'38
	 lgStrSQL = lgStrSQL & "  W_4YEAR_Mth      , "		'39

	 lgStrSQL = lgStrSQL & " INSRT_USER_ID, "
     lgStrSQL = lgStrSQL & " INSRT_DT, "
     lgStrSQL = lgStrSQL & " UPDT_USER_ID, "
     lgStrSQL = lgStrSQL & " UPDT_DT)"
     lgStrSQL = lgStrSQL & " VALUES("
     lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")     & ","                      
     lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")     & ","           
     lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")     & ","             
     lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(3)),"","S")      & ","      
     lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(4)),"","S")       & ","    
     lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(5)),"","S")      & ","    
     lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(6)),"","S")      & "," 
     lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(7)),"","S")      & "," 
     lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(8)),"","S")      & "," 
     lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(9)),"","S")      & "," 
     lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(10)),"","S")      & "," 
    
     lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(11)),0)     & ","   
     lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(12)),0)     & ","   
     lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(13)),0)     & ","   
     lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(14)),0)     & ","   
     lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(15)),0)     & "," 
    
     lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(16)),0)     & "," 
     lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(17)),0)     & "," 
     
     lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(18)),0)     & "," 
     lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(19)),"","S")      & ","    
     lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(20)),0)     & ","
     lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(21)),0)     & ","     
     lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(22)),"","S")      & ","    
  
     
     lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(23)),0)     & "," 
     lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(24)),"","S")      & ","    
     lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(25)),0)     & ","
     lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(26)),0)     & ","     
     lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(27)),"","S")      & ","    
     
     lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(28)),"","S")      & ","    
     lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(29)),0)     & ","    
     lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(30)),"","S")      & ","  
     lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(31)),0)     & ","    
     lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(32)),"","S")      & ","  
     lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(33)),0)     & ","    
     lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(34)),"","S")      & ","  
     lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(35)),0)     & ","    
     lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(36)),"","S")      & ","  
     lgStrSQL = lgStrSQL & FilterVar(Trim(lgKeyStream(37)),"","S")      & ","  
     lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(38)),0)     & ","    
     lgStrSQL = lgStrSQL & UNIConvNum(Trim(lgKeyStream(39)),0)     & ","    
    
                                          
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
    
    lgStrSQL =  " Update TB_JT3B set"
    lgStrSQL = lgStrSQL & " co_cd = " & FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S")  & ","   
    lgStrSQL = lgStrSQL & " FISC_YEAR = " & FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S")  & ","   
    lgStrSQL = lgStrSQL & " rep_type  = " & FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S")   & ","  
    lgStrSQL = lgStrSQL & " SEQ_NO= " & UNIConvNum(arrColVal(1),0)   & ","  
    lgStrSQL = lgStrSQL & " ACCT_NM= " & FilterVar(Trim(UCase(arrColVal(2))),"''","S")    & ","  
    lgStrSQL = lgStrSQL & " W1= " & UNIConvNum(arrColVal(3),0)   & ","  
    lgStrSQL = lgStrSQL & " W2= " & UNIConvNum(arrColVal(4),0)   & ","  
    lgStrSQL = lgStrSQL & " W3= " & UNIConvNum(arrColVal(5),0)   & ","  
    lgStrSQL = lgStrSQL & " W4= " & UNIConvNum(arrColVal(6),0)   & ","  
    lgStrSQL = lgStrSQL & " W5= " & UNIConvNum(arrColVal(7),0)   & ","  
    lgStrSQL = lgStrSQL & " W6= " & UNIConvNum(arrColVal(8),0)   & ","  
    lgStrSQL = lgStrSQL & " W1_T= " & FilterVar(Trim(UCase(arrColVal(9))),"''","S")   & ","
    lgStrSQL = lgStrSQL & " W2_T= " & FilterVar(Trim(UCase(arrColVal(10))),"''","S")   & ","    
    lgStrSQL = lgStrSQL & " W3_T= " & FilterVar(Trim(UCase(arrColVal(11))),"''","S")   & ","  
    lgStrSQL = lgStrSQL & " W4_T= " & FilterVar(Trim(UCase(arrColVal(12))),"''","S")   & ","  
    lgStrSQL = lgStrSQL & " W5_T= " & FilterVar(Trim(UCase(arrColVal(13))),"''","S")   & ","  
    
    lgStrSQL = lgStrSQL & " UPDT_USER_ID   = " & FilterVar(gUsrId,"''","S")   & ","  
    lgStrSQL = lgStrSQL & " UPDT_DT   = " &  FilterVar(GetSvrDateTime,"","S")   
    lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
	lgStrSQL = lgStrSQL & "       and fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
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
    lgStrSQL = "DELETE  TB_JT3B"
   
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
                       lgStrSQL = "SELECT   SEQ_NO , ACCT_NM, w1, w2, w3, w4, w5, w6 , w1_t, w2_t, w3_t, w4_t, w5_t  "
                       lgStrSQL = lgStrSQL & " FROM  TB_JT3B "
                       lgStrSQL = lgStrSQL & " where "
                       lgStrSQL = lgStrSQL &   pComp & pCode  
                        lgStrSQL = lgStrSQL & "order by  SEQ_NO "
                  
                Case "R"
                       lgStrSQL = "SELECT top 1  "
                       lgStrSQL = lgStrSQL & " w8_d1_s  , "   
					   lgStrSQL = lgStrSQL & " w8_d1_e  , "    
				       lgStrSQL = lgStrSQL & " w8_d2_s  , "     
					   lgStrSQL = lgStrSQL & " w8_d2_e  , "    
				       lgStrSQL = lgStrSQL & " w8_d3_s  , "    
					   lgStrSQL = lgStrSQL & " w8_d3_e  , "    
					   lgStrSQL = lgStrSQL & " w8_d4_s  , "     
				       lgStrSQL = lgStrSQL & " w8_d4_e  , "     
	
				       lgStrSQL = lgStrSQL & " w8_Amt1  , "     
				       lgStrSQL = lgStrSQL & " w8_Amt2  , "     
					   lgStrSQL = lgStrSQL & " w8_Amt3  , "    
				       lgStrSQL = lgStrSQL & " w8_Amt4  , "     
					   lgStrSQL = lgStrSQL & " w8_Sum   , "     
	
						lgStrSQL = lgStrSQL & " w9      , "     
						lgStrSQL = lgStrSQL & " w10     , "     
						lgStrSQL = lgStrSQL & " w15_11  , "      
						lgStrSQL = lgStrSQL & " w15_12_View  , "    
						lgStrSQL = lgStrSQL & " w15_12_Value  , "   
						lgStrSQL = lgStrSQL & " w15_13    , "    
						lgStrSQL = lgStrSQL & " w15_14    , "    
	
	
						lgStrSQL = lgStrSQL & " w16_11     , "   
						lgStrSQL = lgStrSQL & " w16_12_View   , "   
						lgStrSQL = lgStrSQL & " w16_12_Value  , "   
						lgStrSQL = lgStrSQL & " w16_13     , "   
						lgStrSQL = lgStrSQL & " w16_14     , "  
						lgStrSQL = lgStrSQL & " COMP_TYPE1  , "  
						lgStrSQL = lgStrSQL & " w17         , "  
						lgStrSQL = lgStrSQL & " w17_14      , "    
						lgStrSQL = lgStrSQL & " w18_A13     , "  
	  				    lgStrSQL = lgStrSQL & " w18_A14     , "    
	  				    lgStrSQL = lgStrSQL & " w18_B13     , "  
	  				    lgStrSQL = lgStrSQL & " w18_B14     ,  "    
	  				    lgStrSQL = lgStrSQL & " w18_C13     , "  
	  				    lgStrSQL = lgStrSQL & " w18_C14     ,  "    
	  				    lgStrSQL = lgStrSQL & " w_DESC      ,"    
						lgStrSQL = lgStrSQL & " W_CURR_YEAR_Mth  ,    "    
						lgStrSQL = lgStrSQL & " W_4YEAR_Mth      "    
	
                       lgStrSQL = lgStrSQL & " FROM  TB_JT3A "
                       lgStrSQL = lgStrSQL & " where "
                       lgStrSQL = lgStrSQL &   pComp & pCode             

                Case "Y"
                      
 
                Case "Z"
                       lgStrSQL = "SELECT  *  "
                       lgStrSQL = lgStrSQL & " FROM  TB_WORK_2 "
                       lgStrSQL = lgStrSQL & " where "
                       lgStrSQL = lgStrSQL &   pComp & pCode           
                   
                       
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
                .ggoSpread.Source     = .frm1.vspdData0
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                
 
                .DBQueryOk
    
                         
	          End With
          ELSE
			  With Parent
				.DBQueryFalse
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
