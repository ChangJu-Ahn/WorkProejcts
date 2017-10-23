<%@ LANGUAGE=VBSCript %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Basis Architect                                                            *
'*  2. Function Name        : User Management, Biz                                                    *
'*  3. Program ID             : ZA003PB2                                                                *
'*  4. Program Name        : Login History Popup 2                                                *
'*  5. Program Desc          : Lists and updates locking status.                                *
'*  7. Modified date(First)  : 2002/05/21                                                            *
'*  8. Modified date(Last)  : 2002/05/21                                                            *
'*  9. Modifier (First)         :    PARK, SANGHOON                                                    *
'* 10. Modifier (Last)        : PARK, SANGHOON                                                    *
'* 11. Comment              :                                                                                *
'********************************************************************************************************

-->

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%                            
On Error Resume Next
'Err.Clear

Call HideStatusWnd

'---------------------------------------Common-----------------------------------------------------------
Call LoadBasisGlobalInf()        
'------ Developer Coding part (Start ) ------------------------------------------------------------------

'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

Dim IntRetCD
Dim lgStrSQL2
Const C_SHEETMAXROWS_D = 30 'Sheet Max Rows

lgOpModeCRUD      = Request("txtMode")                                           

Select Case lgOpModeCRUD

Case CStr(UID_M0001)                                                        


    lgLngMaxRow       = Request("txtMaxRows")                                        
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)   
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           

    '------ Developer Coding part (Start ) ------------------------------------------------------------------
    Call SubOpenDB(lgObjConn)

    Call SubCreateCommandObject(lgObjComm)

    strFromDt = FilterVar(Request("txtFromDt"), "''", "S")

    If Request("txtUsrId") <> "" Then
        strUserId    = FilterVar(Request("txtUsrId"), "''", "S")
    Else
        strUserId = FilterVar("%", "''", "S")
    End If

    strToDt    = FilterVar(Request("txtToDt"), "''", "S")
    strS4  = FilterVar("4", "''", "S")

    If Request("txtUser") <> "" Then
        strClient  = FilterVar(Request("txtClient"), "''", "S")
    Else
        strClient = FilterVar("%", "''", "S")
    End If

    If strFromDt <> "''" then 
        Call SubBizQuery("R")
    End if
        
    Call SubCloseCommandObject(lgObjComm)    
    Call SubCloseDB(lgObjConn)      
%>
<Script Language="VBScript">
    With parent
        If "<%=lgErrorStatus%>" = "NO" And "<%=IntRetCd%>" <> -1 Then
            .ggoSpread.Source = .frm1.vspdData
            .lgStrPrevKeyIndex = "<%=lgStrPrevKeyIndex%>"
            .ggoSpread.SSShowData "<%=ConvSPChars(lgstrData)%>"

            If .frm1.vspdData.MaxRows < .VisibleRowCnt(.frm1.vspdData,0)  And .lgStrPrevKeyIndex <> "" Then     ' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 
                .DbQuery
            Else
                .DbQueryOk
            End If
            .frm1.vspdData.focus
        End If   

    End With           
</Script>    

<%
Case CStr(UID_M0002)        
    Err.Clear                                                                        

    Dim arrVal, arrTemp                                                                
    Dim strStatus                                                                    

    LngMaxRow = CInt(Request("txtMaxRows"))                                            
    arrTemp = Split(Request("txtSpread"), gRowSep)                                    
    
    lGrpCnt = 0
    
    For LngRow = 1 To LngMaxRow
        
        arrVal = Split(arrTemp(LngRow-1), gColSep)

        '------ Developer Coding part (Start ) ------------------------------------------------------------------
        Call SubOpenDB(lgObjConn)
        Call SubCreateCommandObject(lgObjComm)

        If CStr(arrVal(0)) <> "" Then
            strUserId    = FilterVar(CStr(arrVal(0)), "''", "S")
        Else
            strUserId = FilterVar("%", "''", "S")
        End If

        Call SubBizQuery("U")
            
        Call SubCloseCommandObject(lgObjComm)    
        Call SubCloseDB(lgObjConn)      

    Next

%>

<Script Language=vbscript>
    With parent                                                                        
        .DbSaveOk
    End With
</Script>

<%                    
End Select
%>


<%
'=========================================================================================================
Sub SubBizQuery(pType)            ' pType means that R is read mode and U is update mode.

    Dim iDx
    
    On Error Resume Next                                                             
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    If UCase(pType) = "R" Then
        Call SubMakeSQLStatements("R", strUserId, strFromDt, strToDt, strS4, strClient)           
    Else
        Call SubMakeSQLStatements("U", strUserId, strFromDt, strToDt, strS4, strClient)           
    End If
            
    '---------------------------
    ' Header Single 조회 
    '---------------------------    

    If     FncOpenRs(pType,lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    
        IntRetCD = -1
        lgStrPrevKeyIndex = ""        

        Call DisplayMsgBox("210301", vbInformation, "", "", I_MKSCRIPT)      '☜ : Login History Management : Cannot find the data.. 
        Call SetErrorStatus()
    Else
        IntRetCD = 1
         Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKeyIndex)

        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF
        
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(0))
            lgstrData = lgstrData & Chr(11) & SplitTime(lgObjRs(0))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(1))
            lgstrData = lgstrData & Chr(11) & SplitTime(lgObjRs(1))
            lgstrData = lgstrData & Chr(11) & lgObjRs(2)
            lgstrData = lgstrData & Chr(11) & lgObjRs(3)
            lgstrData = lgstrData & Chr(11) & lgObjRs(4)
            lgstrData = lgstrData & Chr(11) & lgObjRs(5)
            lgstrData = lgstrData & Chr(11) & lgObjRs(6)            
    '------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

            lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
               Exit Do
            End If   
               
        Loop 
    End If
    
    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKeyIndex = ""
    End If   

    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                             
        
    lgStrSQL = ""
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub    

'=========================================================================================================
Function SplitTime(Byval dtDateTime)
    SplitTime = Right("0" & Hour(dtDateTime), 2) & ":" _
            & Right("0" & Minute(dtDateTime), 2) & ":" _
            & Right("0" & Second(dtDateTime), 2)
End Function

'=========================================================================================================
Sub SubMakeSQLStatements(pDataType, pUserId, pFromDt, pToDt,  pS4, pClient)

    On Error Resume Next                                                             
    Err.Clear                                                                        
    
    Dim iSelCount

    '------ Developer Coding part (Start ) ------------------------------------------------------------------
	Select Case pDataType

        Case "R"

            lgStrSQL = "SELECT A.LOGIN_DT, A.LOGOUT_DT, B.USR_ID, B.USR_NM, A.STATUS, A.CLIENT_ID, A.CLIENT_IP " 
            lgStrSQL = lgStrSQL & " FROM Z_LOG_IN_HSTRY   A,  Z_USR_MAST_REC   B"
            lgStrSQL = lgStrSQL & " WHERE A.USR_ID = B.USR_ID "
            lgStrSQL = lgStrSQL & " AND A.LOGIN_DT >=  " & pFromDt
            lgStrSQL = lgStrSQL & " AND A.LOGIN_DT <=  " & pToDt
            lgStrSQL = lgStrSQL & " AND B.USR_ID LIKE " & pUserId
            lgStrSQL = lgStrSQL & " AND A.CLIENT_ID LIKE " & pClient
            lgStrSQL = lgStrSQL & " AND A.STATUS = " & pS4 
            lgStrSQL = lgStrSQL & " ORDER BY A.LOGIN_DT   DESC " 

        Case "U"

            lgStrSQL = "UPDATE Z_LOG_IN_HSTRY " 
            lgStrSQL = lgStrSQL & " SET STATUS = " & FilterVar("5", "''", "S") & " "
            lgStrSQL = lgStrSQL & " WHERE STATUS = " & FilterVar("4", "''", "S") & " "

            If pUserId <> "%" Then
                lgStrSQL = lgStrSQL & " AND USR_ID = " & pUserId
            Else
                lgStrSQL = lgStrSQL & " AND USR_ID LIKE " & pUserId
            End If

	End Select   

	%>
	<Script Language=vbscript>
            'msgbox "<%=lgStrSQL%>"
            'parent.frm1.txtUsrId.value = "<%=lgStrSQL%>"
	</Script>
<%            

    '------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub    

'=========================================================================================================
Sub CommonOnTransactionCommit()
    '------ Developer Coding part (Start ) ------------------------------------------------------------------
    '------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'=========================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
    '------ Developer Coding part (Start ) ------------------------------------------------------------------
    '------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'=========================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         
    '------ Developer Coding part (Start ) ------------------------------------------------------------------
    '------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'=========================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             
    Err.Clear                                                                        

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
        Case "MB"
            ObjectContext.SetAbort
            Call SetErrorStatus        
    End Select
End Sub

%>
