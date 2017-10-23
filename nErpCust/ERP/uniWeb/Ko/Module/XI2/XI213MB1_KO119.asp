<%@ LANGUAGE=VBSCript%>
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
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
       
    Call LoadBasisGlobalInf()  
	Call loadInfTB19029B( "I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    Call SubBizQueryMulti()
End Sub    

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    if  Trim(lgKeyStream(0)) <> "" then
         iKey1 = " AND a.plant_cd = " &	FilterVar(Trim(lgKeyStream(0)), "''", "S")
    end if

    If lgKeyStream(1) <> "" AND lgKeyStream(2) <> "" Then
        iKey1 = iKey1 & " AND Convert(varchar(10), a.send_dt, 121) >= " & FilterVar(UNIConvDate(lgKeyStream(1)), "''", "S")
        iKey1 = iKey1 & " AND Convert(varchar(10), a.send_dt, 121) <= " & FilterVar(UNIConvDate(lgKeyStream(2)), "''", "S")
    End If

    if  Trim(lgKeyStream(3)) <> "" then
         iKey1 = iKey1 & " AND a.item_cd LIKE " &	FilterVar(Trim(lgKeyStream(3)) & "%", "''", "S")
    end if

    if  Trim(lgKeyStream(4)) <> "" and  Trim(lgKeyStream(4)) <> "A"  then
         iKey1 =  iKey1 & " AND a.mes_receive_flag = " & FilterVar(Trim(lgKeyStream(4)), "''", "S")  
    end if

    if  Trim(lgKeyStream(5)) <> "" then
         iKey1 = iKey1 & " AND a.item_acct = " & FilterVar(Trim(lgKeyStream(5)), "''", "S")
    end if

    Call SubMakeSQLStatements("MR",iKey1,"X","")                                 '¡Ù : Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("plant_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("spec"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("minor_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_group_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_group_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("std_time"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("valid_from_dt"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("valid_to_dt"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("create_type"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("send_dt"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("mes_receive_flag"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("err_desc"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("mes_receive_dt"))
            
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
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"
                       lgStrSQL = "SELECT TOP " & iSelCount  
                       lgStrSQL = lgStrSQL & "        a.plant_cd plant_cd,					"
					   lgStrSQL = lgStrSQL & "        a.item_cd item_cd,					"
                       lgStrSQL = lgStrSQL & "        a.item_nm item_nm,					"
                       lgStrSQL = lgStrSQL & "        a.spec spec,							"
                       lgStrSQL = lgStrSQL & "        c.minor_nm minor_nm,					"
                       lgStrSQL = lgStrSQL & "        a.item_group_cd item_group_cd,		"
                       lgStrSQL = lgStrSQL & "        d.item_group_nm item_group_nm,		"
                       lgStrSQL = lgStrSQL & "        a.std_time std_time,					"
                       lgStrSQL = lgStrSQL & "        a.valid_from_dt valid_from_dt,		"
                       lgStrSQL = lgStrSQL & "        a.valid_to_dt valid_to_dt,			"
                       lgStrSQL = lgStrSQL & "        a.create_type create_type,			"
                       lgStrSQL = lgStrSQL & "        a.send_dt send_dt,					"
                       lgStrSQL = lgStrSQL & "        a.mes_receive_flag mes_receive_flag,	"
                       lgStrSQL = lgStrSQL & "        a.err_desc err_desc,					"
                       lgStrSQL = lgStrSQL & "        a.mes_receive_dt mes_receive_dt		"
                       lgStrSQL = lgStrSQL & "FROM    T_IF_SND_ITEM_KO119 a, B_ITEM b, B_MINOR c, B_ITEM_GROUP d "
                       lgStrSQL = lgStrSQL & "WHERE   a.item_cd *= b.item_cd "
                       lgStrSQL = lgStrSQL & "    AND c.major_cd = 'P1001' AND a.item_acct *= c.minor_cd "
                       lgStrSQL = lgStrSQL & "    AND a.item_group_cd *= d.item_group_cd " & pCode
'Response.Write lgStrSQL
'Response.End                       
           End Select             
    End Select
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
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .DBQueryOk        
	         End with
          End If   
    End Select    
    
       
</Script>	
