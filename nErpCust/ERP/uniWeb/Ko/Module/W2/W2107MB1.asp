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
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->

<%
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
    '---------------------------------------Common-----------------------------------------------------------
dim lgStrSQL1,lgLngMaxRow2,lgLngMaxRow3
Dim CloseStatus

    lgErrorStatus     = "NO"
    CloseStatus       = "N"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")   
    lgIntFlgMode = CInt(Request("txtFlgMode"))   
    lgLngMaxRow = CInt(Request("txtMaxRows"))   
    lgLngMaxRow2 = CInt(Request("txtMaxRows2"))    

    
                                  '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

   
	
	
	
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    Call CheckVersion( Request("txtFISC_YEAR"),  Request("cboRep_type"))	' 2005-03-11 버전관리기능 추가 
      
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001) 
                                                               '☜: Clear Error status                                                        '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)  
                                                      '☜: Save,Update
             Call SubBizSave()

        Case CStr(UID_M0003)   
             Call SubBizDelete()  

        Case CStr(UID_M0004)   
         
        Case CStr(UID_M0005)   
                                            
             Call SubBizAutoQuery()         
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear      
    
    Call SubBizQueryMulti()
    Call SubBizQueryMulti2()
    Call SubBizQueryMulti3
   

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)

    
    
    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                   '☜ : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE           
              Call SubBizSaveSingleUpdate()
    End Select
    
      
     Call SubBizSaveSingleUPDATE()
     Call SubBizSaveMultiUpdate()
     Call SubBizSaveMulti()                                                             '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next
    Err.Clear

    lgStrSQL = "DELETE From  TB_17_D2"
    lgStrSQL = lgStrSQL & " where "
    lgStrSQL = lgStrSQL &     " co_cd = " &  FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S") 
    lgStrSQL = lgStrSQL &     " and rep_type = " & FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S") 
    lgStrSQL = lgStrSQL &     " and FISC_YEAR = " & FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S") 
    lgStrSQL = "DELETE From  TB_17_D1"
    lgStrSQL = lgStrSQL & " where "
    lgStrSQL = lgStrSQL &     " co_cd = " &  FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S") 
    lgStrSQL = lgStrSQL &     " and rep_type = " & FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S") 
    lgStrSQL = lgStrSQL &     " and FISC_YEAR = " & FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S") 
    lgStrSQL = "DELETE From  TB_17H"
    lgStrSQL = lgStrSQL & " where "
    lgStrSQL = lgStrSQL &     " co_cd = " &  FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S") 
    lgStrSQL = lgStrSQL &     " and rep_type = " & FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S") 
    lgStrSQL = lgStrSQL &     " and FISC_YEAR = " & FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S") 

    
    

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere,IntRetCD
    Dim lgObjRs1,StrSQL
    dim i
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
  
    

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
   
  
    
    strWhere = " co_cd = " & FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S")  
    strWhere = strWhere & "  and fisc_year =" & FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S")
    strWhere = strWhere & "  and rep_type =" & FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S") 
    


     
    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQ)                                 '☆ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then

        iClsRs = 1
      
        'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write "Call Parent.FncNew"  &  vbCrLf
        Response.Write " </Script>"        
	else    

        lgstrData = ""
        iDx       = 1
       

         
        Do While Not lgObjRs.EOF
           
              
           
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("seq_no"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("w1"))
            	lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("w2"))
            	lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("code_no"))
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("w3"))
				lgstrData = lgstrData & Chr(11) & ""
				lgstrData = lgstrData & Chr(11) & RemoveZero(lgObjRs("w4"))
				lgstrData = lgstrData & Chr(11) & RemoveZero(lgObjRs("w5"))
				lgstrData = lgstrData & Chr(11) & RemoveZero(lgObjRs("w6"))
				lgstrData = lgstrData & Chr(11) & RemoveZero(lgObjRs("w7"))
		
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
				lgstrData = lgstrData & Chr(11) &  iDx
				lgstrData = lgstrData & Chr(11) & Chr(12)
			
		        lgObjRs.MoveNext
		  

            iDx =  iDx + 1
       
        Loop 
  
  end if  

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    

Function RemoveZero(Byval pNum)
	If CDbl(pNum) = 0 Then
		RemoveZero = ""
	Else
		RemoveZero = pNum
	End If
End Function

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti2()

    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere,IntRetCD
    Dim lgObjRs1,StrSQL
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Dim SumcurrQty
    Dim SumRcptQty 
    Dim SumGiQty1
    Dim SumGiQty2
    Dim SumStock
    Dim SumStock2
    Dim SumPURQTY
    

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
   
    strWhere = " co_cd = " & FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S")  
    strWhere = strWhere & "  and fisc_year =" & FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S")
    strWhere = strWhere & "  and rep_type =" & FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S") 
    


    
    
      Call SubMakeSQLStatements("MD",strWhere,"X",C_EQ)                                '☆ : Make sql statements
  
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
     
        'Call SetErrorStatus()
   ELSE


         Response.Write "<Script Language=vbscript>" & vbCr
	     Response.Write "With parent" & vbCr
	     Response.Write "	.frm1.txtW8.value       = """ & ConvSPChars(lgObjRs("w8"))       	& """" & vbCr 
	     Response.Write "	.frm1.txtW9.text       = """ & ConvSPChars(lgObjRs("w9"))       	& """" & vbCr
	     Response.Write "	.frm1.txtW10.text       = """ & ConvSPChars(lgObjRs("w10"))       	& """" & vbCr
	     Response.Write "	.frm1.txtW11.text       = """ & ConvSPChars(lgObjRs("w11"))       	& """" & vbCr
	     Response.Write "	.frm1.txtW12.value       = """ & ConvSPChars(lgObjRs("w12"))       	& """" & vbCr
	     Response.Write "	.frm1.txtW13.value       = """ & ConvSPChars(lgObjRs("w13"))       	& """" & vbCr	
	     Response.Write "	.frm1.txtW14.value       = """ & ConvSPChars(lgObjRs("w14"))       	& """" & vbCr	 		
	     Response.Write " End With "	& vbCr
         Response.Write "</Script>"  & vbCr
	     
       
End if
	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
	
    Call SubCloseRs(lgObjRs)    
                                                 '☜: Release RecordSSet

End Sub    


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti3()

    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere,IntRetCD
    Dim lgObjRs1,StrSQL
    dim etc_total
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
   
    

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
   
    
     strWhere = " co_cd = " & FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S")  
     strWhere = strWhere & "  and fisc_year =" & FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S")
     strWhere = strWhere & "  and rep_type =" & FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S") 

    
    
    Call SubMakeSQLStatements("MS",strWhere,"X",C_EQ)                                 '☆ : Make sql statements


    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
     
		'Call SetErrorStatus
    ELSE


        


       ' Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

        lgstrData2 = ""
        iDx       = 1
       
          
         
        Do While Not lgObjRs.EOF
           
           
                lgstrData2 = lgstrData2 & Chr(11) & ConvSPChars(lgObjRs("seq_no"))
				lgstrData2 = lgstrData2 & Chr(11) & ConvSPChars(lgObjRs("w8"))
            	lgstrData2 = lgstrData2 & Chr(11) & ConvSPChars(lgObjRs("w15"))
            	lgstrData2 = lgstrData2 & Chr(11) & RemoveZero(lgObjRs("w9"))
            	lgstrData2 = lgstrData2 & Chr(11) & ConvSPChars(lgObjRs("w_REMARK"))
		
			  
                
   '------ Developer Coding part (End   ) ------------------------------------------------------------------
				lgstrData2 = lgstrData2 & Chr(11) &  iDx
				lgstrData2 = lgstrData2 & Chr(11) & Chr(12)
			
		        etc_total= etc_total +  cdbl(lgObjRs("w9"))	
		       lgObjRs.MoveNext
		 
        Loop 
  
  
                
         Response.Write "<Script Language=vbscript>" & vbCr
	     Response.Write "With parent" & vbCr
	     'Response.Write "	.frm1.txtW99.text       = """ & etc_total  & """" & vbCr 
	     		
	     Response.Write " End With "	& vbCr
         Response.Write "</Script>"  & vbCr
			
End if
	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
	
    Call SubCloseRs(lgObjRs)    
                                                '☜: Release RecordSSet

End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim arrColVal2,arrRowVal2
    Dim arrColVal3,arrRowVal3
    
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing

    Err.Clear                                                                        '☜: Clear Error status

	
	
	
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

    
    if lgErrorStatus <> "YES" then
		arrRowVal2 = Split(Request("txtSpread2"), gRowSep)                                 '☜: Split Row    data
 
		For iDx = 1 To lgLngMaxRow2
		    arrColVal2 = Split(arrRowVal2(iDx-1), gColSep)                                 '☜: Split Column data
		 
		    Select Case arrColVal2(0)
		        Case "C"
		                Call SubBizSaveMultiCreate2(arrColVal2)                            '☜: Create
		        Case "U"
 
		                Call SubBizSaveMultiUpdate2(arrColVal2)                            '☜: Update
		        Case "D"
		                Call SubBizSaveMultiDelete2(arrColVal2)                            '☜: Delete
		    End Select
		    
		    If lgErrorStatus    = "YES" Then
		       lgErrorPos = lgErrorPos & arrColVal2(1) & gColSep
		       Exit For
		    End If
		    
		 Next
    End if
    

End Sub    


'============================================================================================================
' Name : SubBizSaveSingleCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
  
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = " insert into  TB_17H (co_cd,fisc_year,rep_type,w8, w9, w10, w11, w12, w13, w14  ,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) values  "
    lgStrSQL = lgStrSQL & "("
    lgStrSQL = lgStrSQL &  FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S")   & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S") & ","
    lgStrSQL = lgStrSQL &  FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S")   & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtw8"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtw9"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtw10"),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtw11"),0) & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtw12"),0) & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtw13"),0) & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtw14"),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ","  
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S")  & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ","  
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S")  & ")" 


    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
  

    
End Sub


'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
  
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = " insert into  TB_17_D1 (co_cd,fisc_year,rep_type, seq_no ,  w1, w2, code_no ,  w3 ,  w4, w5, w6, w7 ,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) values  "
    lgStrSQL = lgStrSQL & "("
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S")   & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S")  & ","
    lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(2)),"","S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(3)),"","S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(4)),"","S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(5)),"","S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(6)),"","S") & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(7),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(8),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(9),0) & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(10),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ","  
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S")  & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ","  
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S")  & ")" 


    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
    
  
    
End Sub



'============================================================================================================
' Name : SubBizSaveCreate2
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate2(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
  
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = " insert into  TB_17_D2 (co_cd,fisc_year,rep_type, seq_no , w8, w15, w9, w_remark, INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) values  "
    lgStrSQL = lgStrSQL & "("
    lgStrSQL = lgStrSQL &  FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S")   & ","
    lgStrSQL = lgStrSQL &  FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S") & ","
    lgStrSQL = lgStrSQL &  FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S")  & ","
    lgStrSQL = lgStrSQL &  UNIConvNum(arrColVal(2),0) & ","
    lgStrSQL = lgStrSQL &  FilterVar(Trim(arrColVal(3)),"","S") & ","
    lgStrSQL = lgStrSQL &  FilterVar(Trim(arrColVal(4)),"","S") & ","
    lgStrSQL = lgStrSQL &  UNIConvNum(arrColVal(5),0) & ","
    lgStrSQL = lgStrSQL &  FilterVar(Trim(arrColVal(6)),"","S") & ","
    lgStrSQL = lgStrSQL &  FilterVar(gUsrId,"''","S") & ","  
    lgStrSQL = lgStrSQL &  FilterVar(GetSvrDateTime,"''","S")  & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ","  
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S")  & ")" 
    

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
    
  
    
End Sub



'============================================================================================================
' Name : SubBizSaveSingleUPDATE
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUPDATE()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
  
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    
    lgStrSQL =  " if EXISTS (select * from  TB_17H "
    lgStrSQL = lgStrSQL & " where "
    lgStrSQL = lgStrSQL &     " co_cd = " &  FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S") 
    lgStrSQL = lgStrSQL &     " and rep_type = " & FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S") 
    lgStrSQL = lgStrSQL &     " and FISC_YEAR = " & FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S") 
    lgStrSQL = lgStrSQL &  " )  " 
    lgStrSQL = lgStrSQL &  " BEGIN  "
    lgStrSQL = lgStrSQL &"		update  TB_17H set  "
    lgStrSQL = lgStrSQL & "				w8 =	"	&  UNIConvNum(Request("txtw8"),0) & ","
    lgStrSQL = lgStrSQL & "				w9 =	"	&  UNIConvNum(Request("txtw9"),0) & ","
    lgStrSQL = lgStrSQL & "				w10 =	"	&  UNIConvNum(Request("txtw10"),0) & ","
    lgStrSQL = lgStrSQL & "				w11 =	"	&  UNIConvNum(Request("txtw11"),0) & "," 
    lgStrSQL = lgStrSQL & "				w12 =	"	&  UNIConvNum(Request("txtw12"),0) & "," 
    lgStrSQL = lgStrSQL & "				w13 =	"	&  UNIConvNum(Request("txtw13"),0) & ","  
    lgStrSQL = lgStrSQL & "				w14 =	"	&  UNIConvNum(Request("txtw14"),0) & ","      
    lgStrSQL = lgStrSQL & "				UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & ","  
    lgStrSQL = lgStrSQL & "				UPDT_DT = " & FilterVar(GetSvrDateTime,"''","S")  
    lgStrSQL = lgStrSQL & "		where "
    lgStrSQL = lgStrSQL &     " co_cd = " &  FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S") 
    lgStrSQL = lgStrSQL &     " and rep_type = " & FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S") 
    lgStrSQL = lgStrSQL &     " and FISC_YEAR = " & FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S") 
    lgStrSQL = lgStrSQL & " END " 
    lgStrSQL = lgStrSQL & " ELSE " 
    lgStrSQL = lgStrSQL &  " BEGIN  "
    lgStrSQL = lgStrSQL &  " insert into  TB_17H (co_cd,fisc_year,rep_type,w8, w9, w10, w11, w12, w13, w14  ,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) values  "
    lgStrSQL = lgStrSQL & "("
    lgStrSQL = lgStrSQL &	 FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S")   & ","
    lgStrSQL = lgStrSQL &	 FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S") & ","
    lgStrSQL = lgStrSQL &	 FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S")   & ","
    lgStrSQL = lgStrSQL &	 UNIConvNum(Request("txtw8"),0) & ","
    lgStrSQL = lgStrSQL &	 UNIConvNum(Request("txtw9"),0) & ","
    lgStrSQL = lgStrSQL &	 UNIConvNum(Request("txtw10"),0) & ","
    lgStrSQL = lgStrSQL &	 UNIConvNum(Request("txtw11"),0) & "," 
    lgStrSQL = lgStrSQL &	 UNIConvNum(Request("txtw12"),0) & "," 
    lgStrSQL = lgStrSQL &	 UNIConvNum(Request("txtw13"),0) & ","  
    lgStrSQL = lgStrSQL &	 UNIConvNum(Request("txtw14"),0) & ","
    lgStrSQL = lgStrSQL &	 FilterVar(gUsrId,"''","S") & ","  
    lgStrSQL = lgStrSQL &	 FilterVar(GetSvrDateTime,"''","S")  & "," 
    lgStrSQL = lgStrSQL &	 FilterVar(gUsrId,"''","S") & ","  
    lgStrSQL = lgStrSQL &	 FilterVar(GetSvrDateTime,"''","S")  & ")" 
    
    lgStrSQL = lgStrSQL &  " END  "

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
    

    
End Sub



Sub SubBizAutoQuery()

    Dim strWhere
   
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
  
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
   
    strWhere = " co_cd = " & FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S")  
    strWhere = strWhere & "  and fisc_year =" & FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S")
    strWhere = strWhere & "  and rep_type =" & FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S") 
    


    
    
    Call SubMakeSQLStatements("MK",strWhere,"X",C_EQ)                                 '☆ : Make sql statements
  
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
     
        iClsRs = 1
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
   ELSE

         Response.Write "<Script Language=vbscript>" & vbCr
	     Response.Write "With parent" & vbCr
	   
	     Response.Write "	.frm1.txtW9.text       = """ & UNIConvNum(lgObjRs("A_TYPE_AMT"),0)       	& """" & vbCr
	     Response.Write "	.frm1.txtW10.text       = """ & UNIConvNum(lgObjRs("C_TYPE_AMT"),0)       	& """" & vbCr
	     Response.Write "	.frm1.txtW11.text       = """ & UNIConvNum(lgObjRs("K_TYPE_AMT"),0)       	& """" & vbCr
	 
	     Response.Write " End With "	& vbCr
         Response.Write "</Script>"  & vbCr
	     

              
End if
	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
	
    Call SubCloseRs(lgObjRs)    
    Call SubCloseCommandObject(lgObjComm)                                                      '☜: Release RecordSSet

End Sub    





'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
 On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 

    

    lgStrSQL = " update TB_17_D1  set  "
    lgStrSQL = lgStrSQL & " w1 =	"	&  FilterVar(arrColVal(3),"''","S") & ","
    lgStrSQL = lgStrSQL & " w2 =	"	&  FilterVar(arrColVal(4),"''","S") & ","
    lgStrSQL = lgStrSQL & " CODE_NO  =	"	&  FilterVar(arrColVal(5),"''","S") & ","
    lgStrSQL = lgStrSQL & "	w3 =	"	&  FilterVar(arrColVal(6),"''","S") &  ","   
    lgStrSQL = lgStrSQL & " w4 =	"	&  UNIConvNum(arrColVal(7),0)  & "," 
    lgStrSQL = lgStrSQL & " w5 =	"	&  UNIConvNum(arrColVal(8),0)  & ","  
    lgStrSQL = lgStrSQL & " w6 =	"	&  UNIConvNum(arrColVal(9),0)  & ","
    lgStrSQL = lgStrSQL & " w7 =	"	&  UNIConvNum(arrColVal(10),0)  & ","    
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & ","  
    lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(GetSvrDateTime,"''","S")  
    lgStrSQL = lgStrSQL & " where "
    lgStrSQL = lgStrSQL &     " co_cd = " &  FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S") 
    lgStrSQL = lgStrSQL &     " and rep_type = " & FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S") 
    lgStrSQL = lgStrSQL &     " and FISC_YEAR = " & FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S") 
    lgStrSQL = lgStrSQL &     " AND SEQ_NO = " &  UNIConvNum(arrColVal(2),0)
 

  '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub


Sub SubBizSaveMultiUpdate2(arrColVal)
 On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
    

    
    lgStrSQL = " update TB_17_D2  set  "
    lgStrSQL = lgStrSQL & " w8 =	"	&  FilterVar(arrColVal(3),"''","S") & ","
    lgStrSQL = lgStrSQL & " w15 =	"	&  FilterVar(arrColVal(4),"''","S") & ","
    lgStrSQL = lgStrSQL & " w9 =	"	& UNIConvNum(arrColVal(5),0) & ","
    lgStrSQL = lgStrSQL & " w_remark =	"	&  FilterVar(arrColVal(6),"''","S") & ","
    lgStrSQL = lgStrSQL & " UPDT_USER_ID = " &  FilterVar(gUsrId,"''","S") & ","  
    lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(GetSvrDateTime,"''","S")  
    lgStrSQL = lgStrSQL & " where "
    lgStrSQL = lgStrSQL &     " co_cd = " &  FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S") 
    lgStrSQL = lgStrSQL &     " and rep_type = " & FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S") 
    lgStrSQL = lgStrSQL &     " and FISC_YEAR = " & FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S") 
    lgStrSQL = lgStrSQL &     " and SEQ_NO = " &  UNIConvNum(arrColVal(2),0)


                                                  '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

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

    lgStrSQL = " delete from  TB_17_D1    "
    lgStrSQL = lgStrSQL & " where "
    lgStrSQL = lgStrSQL &     " co_cd = " &  FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S") 
    lgStrSQL = lgStrSQL &     " and rep_type = " & FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S") 
    lgStrSQL = lgStrSQL &     " and FISC_YEAR = " & FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S") 
    lgStrSQL = lgStrSQL &     " and SEQ_NO = " &  UNIConvNum(arrColVal(2),0)


    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
    
    If Err.number <> 0 Then
    	Call svrmsgbox(Err.Description & vbCrLf & "Error Code : " & Err.Number , vbCritical, i_mkscript)
        ObjectContext.SetAbort
        Call SetErrorStatus
    End If
    
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete2
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete2(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = " delete from  TB_17_D2    "
    lgStrSQL = lgStrSQL & " where "
    lgStrSQL = lgStrSQL &     " co_cd = " &  FilterVar(Trim(UCase(Request("txtCo_Cd"))),"","S") 
    lgStrSQL = lgStrSQL &     " and rep_type = " & FilterVar(Trim(UCase(Request("cboREP_TYPE"))),"","S") 
    lgStrSQL = lgStrSQL &     " and FISC_YEAR = " & FilterVar(Trim(UCase(Request("txtFISC_YEAR"))),"","S") 
    lgStrSQL = lgStrSQL &     " and SEQ_NO = " &  UNIConvNum(arrColVal(2),0)


    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
    
    If Err.number <> 0 Then
    	Call svrmsgbox(Err.Description & vbCrLf & "Error Code : " & Err.Number , vbCritical, i_mkscript)
        ObjectContext.SetAbort
        Call SetErrorStatus
    End If
    
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
	    
        Case "M"
        
     
           
           Select Case Mid(pDataType,2,1)
               Case "R"
                       lgStrSQL = "SELECT seq_no ,code_no,  w1, w2, w3, w4, w5, w6 , w7   "
                       lgStrSQL = lgStrSQL & " FROM  TB_17_D1 "
                       lgStrSQL = lgStrSQL & " where "
                       lgStrSQL = lgStrSQL &   pCode                  

                      
              Case "D"
					   lgStrSQL = "SELECT  w8, w9, w10, w11, w12, w13  "
                       lgStrSQL = lgStrSQL & " FROM  TB_17H "
                       lgStrSQL = lgStrSQL & " where "
                       lgStrSQL = lgStrSQL &   pCode    
    
              Case "S"
                       lgStrSQL = "SELECT seq_no , w8, w15, w9, w_remark  "
                       lgStrSQL = lgStrSQL & " FROM  TB_17_D2 "
                       lgStrSQL = lgStrSQL & " where "
                       lgStrSQL = lgStrSQL &   pCode  
                    
                                          
                       
              Case "K"
                       lgStrSQL = "SELECT      A_TYPE_AMT , K_TYPE_AMT , C_TYPE_AMT  "
                       lgStrSQL = lgStrSQL & " FROM  TB_WORK_8 "
                       lgStrSQL = lgStrSQL & " where "
                       lgStrSQL = lgStrSQL &   pCode  
                   

              Case "X"
                    
                       
                      
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

<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
            
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .ggoSpread.Source     = .frm1.vspdData2
                .ggoSpread.SSShowData "<%=lgstrData2%>"    
 
                .DBQueryOk
    
                         
	          End With
          End If
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
		  Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             
              Parent.DBDeleteOk
          Else   
          End If
        Case "<%=UID_M0004%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             
              Parent.DBSaveOk
          Else   
          End If         
    End Select    
    
</Script>
