<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!--
======================================================================================================
'********************************************************************************************************
'*  1. Module Name          : Inventory																*
'*  2. Function Name        : Popup Item By Plant														*	
'*  3. Program ID           : i2512pb1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Item by Plant Popup														*
'*  7. Modified date(First) : 2000/03/29																*
'*  8. Modified date(Last)  : 2002/04/17																*
'*  9. Modifier (First)     : Im Hyun Soo																*
'* 10. Modifier (Last)      : Ahn Jung Je																*
'* 11. Comment              :																			*
=======================================================================================================-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
Call LoadBasisGlobalInf()
Call HideStatusWnd                                                               '☜: Hide Processing message

On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status
    '---------------------------------------Common-----------------------------------------------------------
    lgLngMaxRow       = Request("txtMaxRows")                                    '☜: Read Operation Mode (CRUD)
    lgMaxCount        = 100                              '☜: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = Trim(Request("lgStrPrevKey"))
    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                      '☜: Set to space
'    lgOpModeCRUD      = Request("txtMode")                                      '☜: Read Operation Mode (CRUD)
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
'Dim IntRetCD
	
Dim strPlantCd
Dim strItemCd
Dim strItemNm
Dim strFromItemAcct
Dim strToItemAcct
Dim strSpec
Dim strLotFlg

	strPlantCd = FilterVar(Trim(Request("PlantCd"))	, "''", "S")
	
	strItemCd  = FilterVar(Request("txtItemCd"), "''", "S")
	if lgStrPrevKeyIndex <> "" then strItemCd  = FilterVar(lgStrPrevKeyIndex, "''", "S")
	
	strItemNm  = FilterVar("%" & Trim(Request("txtItemNm")) & "%", "''", "S")
	
	If Request("cboItemAccount") <> "" Then		'품목계정 
		strFromItemAcct = FilterVar(Trim(Request("cboItemAccount"))	, "''", "S")
		strToItemAcct   = FilterVar(Request("cboItemAccount"), "''", "S")
	Else
		If Request("ToItemAcct") <> "" Then
			strFromItemAcct = FilterVar(Request("FromItemAcct"), "''", "S")
			strToItemAcct   = FilterVar(Request("ToItemAcct"), "''", "S")
		End If
	End If
	
	strSpec   = FilterVar("%" & Trim(Request("txtItemSpec")) & "%", "''", "S")
	strLotFlg = FilterVar(Request("rdoLotItem"), "''", "S")
	
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	
	IF strItemNm = "" & FilterVar("%%", "''", "S") & "" Or (strItemCd <> "''" And strItemNm <> "" & FilterVar("%%", "''", "S") & "" ) Then
		Call SubBizQueryMulti("ITEM_CD")
	Else		
        Call SubBizQueryMulti("ITEM_NM")
    End If
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti(pType)
	
	'Call ServerMesgBox("시작 >", vbCritical, I_MKSCRIPT)
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim PrntKey
	Dim NodX
	Dim Node		    
	Dim iDx
	Dim PvArr

	If pType = "ITEM_CD" Then
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
		Call SubMakeSQLStatements("CD",strPlantCd,strItemCd,strItemNm,strFromItemAcct,strToItemAcct,strSpec,strLotFlg)           '☜ : Make sql statements
	Else	
		Call SubMakeSQLStatements("NM",strPlantCd,strItemCd,strItemNm,strFromItemAcct,strToItemAcct,strSpec,strLotFlg)           '☜ : Make sql statements
	End If

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		lgStrPrevKey = ""    
'		IntRetCD = -1
		Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()
				 
		Response.End
     
	Else
'		IntRetCD = 1
		iDx = 0
'		Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex )
		ReDim PvArr(0)
        Do While Not lgObjRs.EOF
			
			ReDim Preserve PvArr(iDx)
				
			lgstrData =		Chr(11) & Trim(lgObjRs("ITEM_CD")) & _
							Chr(11)	& lgObjRs("ITEM_NM") & _
							Chr(11) & lgObjRs("SPEC") & _
							Chr(11) & Trim(lgObjRs("BASIC_UNIT")) & _
							Chr(11) & Trim(lgObjRs(2)) & _
							Chr(11)	& Trim(lgObjRs("ITEM_GROUP_CD")) & _
							Chr(11) & lgObjRs("LOT_FLG") & _
							Chr(11)	& lgObjRs("TRACKING_FLG") & _
							Chr(11) & Trim(lgObjRs("MAJOR_SL_CD")) & _
							Chr(11)	& Trim(lgObjRs("ISSUED_SL_CD")) & _
							Chr(11)	& lgObjRs(57) & _
							Chr(11)	& UNIDateClientFormat(lgObjRs(58)) & _
							Chr(11) & UNIDateClientFormat(lgObjRs(59)) & _
							Chr(11) & lgLngMaxRow + iDx & Chr(11) & Chr(12)

			PvArr(iDx) = lgstrData
		    lgObjRs.MoveNext

	        iDx =  iDx + 1
            If iDx = lgMaxCount Then									
               Exit Do
            End If   
        Loop 
			lgstrData = Join(PvArr, "")
			
		If iDx >= lgMaxCount  Then
			lgStrPrevKeyIndex = Trim(lgObjRs("ITEM_CD"))
		else
			lgStrPrevKeyIndex = ""
		End If   

		Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
		Call SubCloseRs(lgObjRs)       
    End If
 
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords 
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub    

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2,pCode3,pCode4,pCode5,pCode6)
    Dim iSelCount
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case pDataType

		Case "CD"
			lgStrSQL = "SELECT Top " & CStr(lgMaxCount + 1) & " A.*, B.* "
			lgStrSQL = lgStrSQL & " FROM B_ITEM_BY_PLANT A, B_ITEM B "
			lgStrSQL = lgStrSQL & " WHERE A.ITEM_CD = B.ITEM_CD AND B.PHANTOM_FLG = " & FilterVar("N", "''", "S") & " "
			lgStrSQL = lgStrSQL & " AND  A.PLANT_CD = " & pCode
			lgStrSQL = lgStrSQL & " AND A.ITEM_CD >= " & pCode1
			lgStrSQL = lgStrSQL & " AND B.ITEM_NM Like  " & pCode2
			lgStrSQL = lgStrSQL & " AND A.ITEM_ACCT >= " & pCode3
			lgStrSQL = lgStrSQL & " AND A.ITEM_ACCT <= " & pCode4
			lgStrSQL = lgStrSQL & " AND B.SPEC Like " & pCode5
			lgStrSQL = lgStrSQL & " AND A.LOT_FLG Like " & pCode6

			lgStrSQL = lgStrSQL & " ORDER BY A.ITEM_CD, B.ITEM_NM " 
			
		Case "NM"
			lgStrSQL = "SELECT Top " & CStr(lgMaxCount + 1) & "  A.*, B.* "
			lgStrSQL = lgStrSQL & " FROM B_ITEM_BY_PLANT A, B_ITEM B "
			lgStrSQL = lgStrSQL & " WHERE A.ITEM_CD = B.ITEM_CD AND B.PHANTOM_FLG = " & FilterVar("N", "''", "S") & " "
			lgStrSQL = lgStrSQL & " AND  A.PLANT_CD = " & pCode
			lgStrSQL = lgStrSQL & " AND A.ITEM_CD >= " & pCode1
			lgStrSQL = lgStrSQL & " AND B.ITEM_NM Like  " & pCode2
			lgStrSQL = lgStrSQL & " AND A.ITEM_ACCT >= " & pCode3
			lgStrSQL = lgStrSQL & " AND A.ITEM_ACCT <= " & pCode4
			
			lgStrSQL = lgStrSQL & " AND B.SPEC Like " & pCode5 
			lgStrSQL = lgStrSQL & " AND A.LOT_FLG Like " & pCode6
			
			
			lgStrSQL = lgStrSQL & " ORDER BY B.ITEM_NM, A.ITEM_CD " 
			
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
<Script Language="VBScript">
	With parent
		
			
			.lgStrPrevKey	    = "<%=lgStrPrevKeyIndex%>"	 
			
	        .ggoSpread.Source   = .vspdData
	        .ggoSpread.SSShowData "<%=ConvSPChars(lgstrData)%>"
		If .vspdData.MaxRows < .VisibleRowCnt(.vspdData,0)  And .lgStrPrevKey <> "" Then	
			.DbQuery
		Else
			.DbQueryOk
		End If
		.vspdData.focus
	End With	
       
</Script>
