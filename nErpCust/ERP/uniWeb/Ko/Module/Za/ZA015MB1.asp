<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
	DIM lgUsrID
	Dim MaxTimeOutValue
	
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")
    
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgUsrID			  = Request("txtUsrId")

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)    
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)           
    

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query        
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
			 Call SubMaxTimeOutQuery()
             Call SubBizSaveMulti()
    End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next
    Err.Clear
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
   On Error Resume Next
   Err.Clear            
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubMaxTimeOutQuery()
    Dim MAXTIMEOUT
    
    On Error Resume Next    
    Err.Clear
    
    lgStrSQL = ""
	lgStrSQL = "SELECT REFERENCE FROM B_CONFIGURATION WHERE MAJOR_CD  = 'Z0050' AND MINOR_CD = 5 "
   
    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
       lgStrPrevKey = ""
    
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
       Call SetErrorStatus()
       
    Else

        IF NOT lgObjRs.EOF THEN
			MaxTimeOutValue = lgObjRs(0)
        ELSE
			MaxTimeOutValue = 60
        END IF
               
    End If
    
    Call SubCloseRs(lgObjRs)    

End Sub    

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere
    Dim MAXTIMEOUT
    
    On Error Resume Next    
    Err.Clear
    
    '조회시 Z_USR_MAST_REC 테이블에 신규사용자는 insert 없는사용자는 delete *************************
	lgStrSQL =			  "Declare @ReOpenIE	As Char(1) " 
	lgStrSQL = lgStrSQL & "Declare @FontName	As VarChar(13) " 
	lgStrSQL = lgStrSQL & "Declare @FontSize	As Int " 
	lgStrSQL = lgStrSQL & "Declare @TimeOut	As Int " 
	lgStrSQL = lgStrSQL & "IF EXISTS(SELECT * FROM B_CONFIGURATION WHERE MAJOR_CD  = 'Z0050' AND MINOR_CD = '1' AND REFERENCE = 'Y') " 
	lgStrSQL = lgStrSQL & "BEGIN " 
	lgStrSQL = lgStrSQL & "SELECT @ReOpenIE = REFERENCE FROM B_CONFIGURATION WHERE MAJOR_CD  = 'Z0052' AND MINOR_CD = 3  " 
	lgStrSQL = lgStrSQL & "SELECT @FontName = REFERENCE FROM B_CONFIGURATION WHERE MAJOR_CD  = 'Z0052' AND MINOR_CD = 1  " 
	lgStrSQL = lgStrSQL & "SELECT @FontSize = Convert(int,REFERENCE) FROM B_CONFIGURATION WHERE MAJOR_CD  = 'Z0052' AND MINOR_CD = 2  " 
	lgStrSQL = lgStrSQL & "SELECT @TimeOut = Convert(int,REFERENCE) FROM B_CONFIGURATION WHERE MAJOR_CD  = 'Z0050' AND MINOR_CD = 3 " 
	lgStrSQL = lgStrSQL & "DELETE  FROM Z_CONNECTOR_CONFIG WHERE USR_ID NOT IN  ( SELECT USR_ID FROM Z_USR_MAST_REC ) " 
	lgStrSQL = lgStrSQL & "INSERT INTO Z_CONNECTOR_CONFIG " 
	lgStrSQL = lgStrSQL & "SELECT  " 
	lgStrSQL = lgStrSQL & "USR_ID,@ReOpenIE,@FontName,@FontSize,@TimeOut,'unierp',getdate(),'unierp',getdate() " 
	lgStrSQL = lgStrSQL & "FROM Z_USR_MAST_REC " 
	lgStrSQL = lgStrSQL & "Where USR_ID NOT IN ( SELECT USR_ID FROM Z_CONNECTOR_CONFIG ) " 
	lgStrSQL = lgStrSQL & "END " 
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
    '*************************************************************************************************
    
   
	lgStrSQL = "SELECT A.USR_ID,B.USR_NM,B.USR_ENG_NM,REOPENIE,FONT_NAME,FONT_SIZE,TIMEOUT  "
	lgStrSQL = lgStrSQL & "FROM Z_CONNECTOR_CONFIG A,Z_USR_MAST_REC B "
	lgStrSQL = lgStrSQL & "WHERE A.USR_ID = B.USR_ID AND A.USR_ID LIKE '%" & lgUsrID & "%' "
    
'    Call SubMakeSQLStatements("MR",strWhere,"X",C_EQGT)                              '☜ : Make sql statements

    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
       lgStrPrevKey = ""
    
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
       Call SetErrorStatus()
    Else
       Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
       lgstrData = ""
       iDx       = 1
       Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("USR_ID"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("USR_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("USR_ENG_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FONT_NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FONT_SIZE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REOPENIE"))                  
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TIMEOUT"))
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
	   	    lgObjRs.MoveNext
          
            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If   
                       
        Loop 
        
    End If

    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey = ""
    End If     
                
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)    

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next
    Err.Clear 
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	
	For iDx = 1 To lgLngMaxRow
		arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
		
		IF Cint(arrColVal(5)) > Cint(MaxTimeOutValue) Then
			Response.Write "<script language=vbscript>Parent.LimitMsg(" & MaxTimeOutValue & ")</script>"
			Exit Sub
		END IF
	Next
		
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
		
        Select Case arrColVal(0)
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next
End Sub      

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status
    For i = 6 To 11
		IF arrColVal(i) = "YES" then
		    arrColVal(i) = "Y"
		ElseIF arrColVal(i) = "NO" then
		    arrColVal(i) = "N"
		End IF	       
	Next      
	
    lgStrSQL = "UPDATE  Z_CONNECTOR_CONFIG"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " REOPENIE	   = " & FilterVar(UCase(arrColVal(4)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " FONT_SIZE     = " & FilterVar(UCase(arrColVal(3)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " TIMEOUT       = " & FilterVar(UCase(arrColVal(5)), "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_DT      = GETDATE() "
    lgStrSQL = lgStrSQL & " WHERE            "
    lgStrSQL = lgStrSQL & " USR_ID       = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
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
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .DBQueryOk        
	         End with        
	         
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
    End Select       
       
</Script>	
