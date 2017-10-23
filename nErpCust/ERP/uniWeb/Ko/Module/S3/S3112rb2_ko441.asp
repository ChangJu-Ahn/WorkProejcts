<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<%
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd                                                               '☜: Hide Processing message
Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "*", "NOCOOKIE", "PB")

'---------------------------------------Common-----------------------------------------------------------
lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
lgStrPrevKeyIndex = CInt(Request("lgStrPrevKeyIndex"))   
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space
lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

'------ Developer Coding part (Start ) ------------------------------------------------------------------
Dim IntRetCD
	

Const C_SHEETMAXROWS_D = 100


	
'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

Call SubBizQueryMulti()

Call SubCloseDB(lgObjConn)

Response.End



'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim PrntKey
	Dim NodX
	Dim Node
	Dim iIntCnt
	Dim TmpBuffer
	Dim iTotalStr
	Dim strWhere
	
	

	   strWhere = strWhere & " AND  A.SHIP_TO_PARTY  LIKE  " &  FilterVar(lgKeyStream(0),"'%'","S")
	   strWhere = strWhere & " AND  ISNULL(A.packing_list,'') LIKE " &  FilterVar(lgKeyStream(2),"'%'","S")
	   strWhere = strWhere & " AND  ISNULL(B.ITEM_CD,'')  like  " &  FilterVar(lgKeyStream(1),"'%'","S")
       
       
	Call SubMakeSQLStatements("MR",strWhere,"X","")   
	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists
		lgStrPrevKeyIndex = ""    
		IntRetCD = -1
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()
		
		Response.End
	End If

	IntRetCD = 1
	iIntCnt = 1
ReDim TmpBuffer(0)
		Call SubSkipRs(lgObjRs, C_SHEETMAXROWS_D * lgStrPrevKeyIndex )
        Do While Not lgObjRs.EOF
		lgstrData = ""
		    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("plant_cd"))
		    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("packing_list"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("out_type"))
            lgstrData = lgstrData & Chr(11) &  ConvSPChars(lgObjRs("gi_qty")) 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("send_dt"))
			
			lgstrData = lgstrData & Chr(11) & (lgLngMaxRow + iIntCnt)
			lgstrData = lgstrData & Chr(11) & Chr(12)

		lgObjRs.MoveNext
		
		ReDim Preserve TmpBuffer(iIntCnt-1)
		
		TmpBuffer(iIntCnt-1) = lgstrData
		
		iIntCnt =  iIntCnt + 1

	    If iIntCnt > C_SHEETMAXROWS_D Then
			Exit Do
	    End If
	Loop
		
	If lgObjRs.EOF Then
		lgStrPrevKeyIndex = ""
	Else
		lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
	End If
 
	iTotalStr = Join(TmpBuffer, "")
			
	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "With parent" & vbCrLf
	
	If Trim(lgErrorStatus) = "NO" And IntRetCd <> -1 Then
	    Response.Write ".ggoSpread.Source = .vspdData" & vbCrLf
		Response.Write ".lgStrPrevKeyIndex = """ & lgStrPrevKeyIndex & """" & vbCrLf
	    Response.Write ".ggoSpread.SSShowDataByClip " & """" & ConvSPChars(iTotalStr) & """" & vbCrLf
	End If
			

			
	' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 
	Response.Write "If .vspdData.MaxRows < .VisibleRowCnt(.vspdData, 0)  And .lgStrPrevKeyIndex <> """" Then" & vbCrLf
	Response.Write "	.DbQuery" & vbCrLf
	Response.Write "Else" & vbCrLf
	Response.Write "	.DbQueryOk" & vbCrLf
	Response.Write "End If" & vbCrLf
	
	Response.Write ".vspdData.Focus" & vbCrLf
	
	Response.Write "End With" & vbCrLf
	Response.Write "</Script>" & vbCrLf
        
	Call SubCloseRs(lgObjRs)       
    
End Sub    

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)

  Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
        
       
           
           Select Case Mid(pDataType,2,1)
              
               Case "R"
                        
                        lgStrSQL = " select a.plant_cd, "
                        lgStrSQL = lgStrSQL & " a.packing_list, "
                        lgStrSQL = lgStrSQL & " b.item_cd, "
                        lgStrSQL = lgStrSQL & " b.item_nm, "
                        lgStrSQL = lgStrSQL & " a.out_type, "
                        lgStrSQL = lgStrSQL & " a.gi_qty, "
                        lgStrSQL = lgStrSQL & " a.send_dt  "
                        lgStrSQL = lgStrSQL & " from T_IF_RCV_VIRTURE_OUT_KO441 a "
                        lgStrSQL = lgStrSQL & " inner join b_item b on a.mes_item_cd = b.CBM_DESCRIPTION and b.item_acct in('10','20') "
                        lgStrSQL = lgStrSQL & " WHERE 1=1 "
                        lgStrSQL = lgStrSQL &  pComp & pCode
                        lgStrSQL = lgStrSQL & " ORDER BY   A.plant_cd ,A.packing_list,B.item_cd"
                        
 
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
