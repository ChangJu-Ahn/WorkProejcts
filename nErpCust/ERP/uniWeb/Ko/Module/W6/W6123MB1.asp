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
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<!-- #Include file="../wcm/inc_SvrDebug.asp" -->
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

        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
         
                 Call SubBizSaveMulti()
        
            
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

        'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else
  
        
		
             
       lgstrData = ""
        iDx       = 1
       

         
        Do While Not lgObjRs.EOF
           
            

		
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("w1"))
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("w1_nm"))
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("w1_1"))
                
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("w2"))
				
            	lgstrData = lgstrData & Chr(11) & ""
            	lgstrData = lgstrData & Chr(11) & UNIConvNum(lgObjRs("w3"),0)
            	lgstrData = lgstrData & Chr(11) & UNIConvNum(lgObjRs("w4"),0)
            	lgstrData = lgstrData & Chr(11) & UNIConvNum(lgObjRs("w5"),0)
            	lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("w6"))
            	lgstrData = lgstrData & Chr(11) & UNIConvNum(lgObjRs("w2_A"),0)
            	lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("w2_B"))
            	lgstrData = lgstrData & Chr(11) & UNIConvNum(lgObjRs("w2_C"),0)
            	lgstrData = lgstrData & Chr(11) & UNIConvNum(lgObjRs("w2_C_Val"),0)
            	lgstrData = lgstrData & Chr(11) & UNIConvNum(lgObjRs("w2_D"),0)
 

	'------ Developer Coding part (End   ) ------------------------------------------------------------------
				lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
				lgstrData = lgstrData & Chr(11) & Chr(12)
			
		        iDx = iDx + 1
		      
		        
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
 
End Sub

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
   On Error Resume Next                                                             '☜: Protect system from crashing
   Err.Clear
        lgStrSQL = "delete from TB_8_2 " 
	    lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S") 
		lgStrSQL = lgStrSQL & "		  and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S") 
		lgStrSQL = lgStrSQL & "		  and rep_type =" &FilterVar(Trim(UCase(lgKeyStream(2))),"","S") 
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

'    On Error Resume Next
    Err.Clear                                                                        '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	
	
	    arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data

            For iDx = 1 To lgLngMaxRow
                arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
	
	
	
					 strWhere = " co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
					 strWhere = strWhere & "  and fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
					 strWhere = strWhere & "  and rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")
					 strWhere = strWhere & "  and w1 =" &  FilterVar(Trim(UCase(arrColVal(1))),"","S")

    
					Call SubMakeSQLStatements("MR",strWhere,"X","")                              '☜ : Make sql statements
    
					If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True  and  iDx <> lgLngMaxRow  and  arrColVal(0) = "C" Then
					    lgStrPrevKeyIndex = ""
					    
					      Call DisplayMsgbox("970001","X",lgObjRs("w1_nm") ,"X" ,I_MKSCRIPT)      '☜ : No data is found. 
					      
					      Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
					      Call SetErrorStatus
					        
						  Call SubCloseRs(lgObjRs)

					      Exit sub
				    end if	
		   Next
					
            
	    
           For iDx = 1 To lgLngMaxRow
                arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data   
                'PrintLog "arrColVal=" & arrRowVal(iDx-1)
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

            
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    dim iObjPS5G115
    'On Error Resume Next
    Err.Clear                                                                        '☜: Clear Error status
   


    
    
    lgStrSQL = "INSERT INTO TB_8_2("
    lgStrSQL = lgStrSQL & " co_cd, "
    lgStrSQL = lgStrSQL & " FISC_YEAR, "
    lgStrSQL = lgStrSQL & " rep_type, "
    lgStrSQL = lgStrSQL & " w1, "
    lgStrSQL = lgStrSQL & " w1_nm, "
    lgStrSQL = lgStrSQL & " w3, "
    lgStrSQL = lgStrSQL & " w4, "
    lgStrSQL = lgStrSQL & " w5, "
    lgStrSQL = lgStrSQL & " w6, "
    lgStrSQL = lgStrSQL & " w2_a, "
    lgStrSQL = lgStrSQL & " w2_b, "
    lgStrSQL = lgStrSQL & " w2_c, "
    lgStrSQL = lgStrSQL & " w2_c_val, "
    lgStrSQL = lgStrSQL & " w2_d, "

	' -- 2006-3: 개정서식 반영  : 기존 팝업 방식을 제거하고 직접 입력하게 수정
    lgStrSQL = lgStrSQL & " w1_1, "
    lgStrSQL = lgStrSQL & " w2, "
    
    lgStrSQL = lgStrSQL & " INSRT_USER_ID, "
    lgStrSQL = lgStrSQL & " INSRT_DT, "
    lgStrSQL = lgStrSQL & " UPDT_USER_ID, "
    lgStrSQL = lgStrSQL & " UPDT_DT)"
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")     & ","                      
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")     & ","           
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")     & ","  
    lgStrSQL = lgStrSQL  &   FilterVar(Trim(UCase(arrColVal(1))),"","S")  & ","  
    lgStrSQL = lgStrSQL  &   FilterVar(Trim(UCase(arrColVal(2))),"","S")  & ","               
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(3),0) & ","    
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(4),0) & ","    
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(5),0) & ","    
    lgStrSQL = lgStrSQL  &  FilterVar(Trim(UCase(arrColVal(6))),"","S")  & ","  
    lgStrSQL = lgStrSQL  &   UNIConvNum(arrColVal(7),0)  & ","    
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(8),0) & ","    
    lgStrSQL = lgStrSQL  &  FilterVar(Trim(UCase(arrColVal(9))),"","S") & "," 
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(10),0) & ","    
    lgStrSQL = lgStrSQL  &  UNIConvNum(arrColVal(11),0) & "," 
    
	' -- 2006-3: 개정서식 반영  : 기존 팝업 방식을 제거하고 직접 입력하게 수정
    lgStrSQL = lgStrSQL  &  FilterVar(Trim(UCase(arrColVal(12))),"","S") & "," 
    lgStrSQL = lgStrSQL  &  FilterVar(Trim(UCase(arrColVal(13))),"","S") & "," 
    
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                       
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","                                     
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                     
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S")                                            
    lgStrSQL = lgStrSQL & ")"

	'PrintLog lgStrSQL
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
    
    lgStrSQL =  " Update TB_8_2 set"

    lgStrSQL = lgStrSQL & " w1_nm= " & FilterVar(Trim(UCase(arrColVal(2))),"''","S")    & ","  
    lgStrSQL = lgStrSQL & " w3= " & UNIConvNum(arrColVal(3),0)    & ","  
    lgStrSQL = lgStrSQL & " W4= " & UNIConvNum(arrColVal(4),0)   & ","  
    lgStrSQL = lgStrSQL & " W5= " & UNIConvNum(arrColVal(5),0)   & ","  
    lgStrSQL = lgStrSQL & " W6= " & FilterVar(Trim(UCase(arrColVal(6))),"''","S")    & ","  
    lgStrSQL = lgStrSQL & " W2_A= " & UNIConvNum(arrColVal(7),0)   & ","  
    lgStrSQL = lgStrSQL & " W2_B= " & UNIConvNum(arrColVal(8),0)   & ","  
    lgStrSQL = lgStrSQL & " W2_C= " & FilterVar(Trim(UCase(arrColVal(9))),"''","S")   & ","  
    lgStrSQL = lgStrSQL & " W2_C_val= " & UNIConvNum(arrColVal(10),0)   & ","  
    lgStrSQL = lgStrSQL & " W2_D= " & UNIConvNum(arrColVal(11),0)   & ","  

	' -- 2006-3: 개정서식 반영  : 기존 팝업 방식을 제거하고 직접 입력하게 수정
    lgStrSQL = lgStrSQL & " W1_1= " & FilterVar(Trim(UCase(arrColVal(12))),"''","S")   & ","  
    lgStrSQL = lgStrSQL & " W2= " & FilterVar(Trim(UCase(arrColVal(13))),"''","S")   & ","  
    
    lgStrSQL = lgStrSQL & " UPDT_USER_ID   = " & FilterVar(gUsrId,"''","S")   & ","  
    lgStrSQL = lgStrSQL & " UPDT_DT   = " &  FilterVar(GetSvrDateTime,"","S")   
	lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S") 
	lgStrSQL = lgStrSQL & "		  and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S") 
	lgStrSQL = lgStrSQL & "		  and rep_type =" &FilterVar(Trim(UCase(lgKeyStream(2))),"","S") 
    lgStrSQL = lgStrSQL & "		  and w1= " & FilterVar(Trim(UCase(arrColVal(1))),"''","S") 
    
'PrintLog lgStrSQL
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
    lgStrSQL = "DELETE  TB_8_2"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "  co_cd = " & FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S")  
    lgStrSQL = lgStrSQL & "  and fisc_year =" & FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S")
    lgStrSQL = lgStrSQL & "  and rep_type =" & FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S") 
    lgStrSQL = lgStrSQL & "  and w1 =" & UNIConvNum(arrColVal(1),0) 

'PrintLog lgStrSQL
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
                     
                Case "R"
                       lgStrSQL = "	 SELECT w1,    w1_nm, w1_1,  w2,"
                       lgStrSQL = lgStrSQL & "  w3       ,w4   ,w5	, w6  ,   w2_A,  w2_B,           w2_C               ,    w2_c_val ,   w2_D   "
                       lgStrSQL = lgStrSQL & " FROM  TB_8_2  "
                       lgStrSQL = lgStrSQL & " where "
                       lgStrSQL = lgStrSQL &   pComp & pCode 
                       'lgStrSQL = lgStrSQL & "Order by  w1 "           


 
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
                Parent.DbDeleteOk
            Else
            End If
    End Select
    
</Script>
