<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../inc/IncSvrMain.asp" -->
<!-- #Include file="../inc/IncSvrDate.inc" -->
<!-- #Include file="../inc/AdoVbs.inc" -->
<!-- #Include file="../inc/lgSvrVariables.inc" -->
<!-- #Include file="../inc/incServerAdoDB.asp" -->
<%
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd                                                               '☜: Hide Processing message
Call LoadBasisGlobalInf

'---------------------------------------Common-----------------------------------------------------------
lgLngMaxRow       = Request("txtMaxRows")						                                     '☜: Fetch count at a time for VspdData
lgStrPrevKeyIndex = CInt(Request("lgStrPrevKeyIndex"))   
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space
lgMaxCount		  = 30
lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

'------ Developer Coding part (Start ) ------------------------------------------------------------------
Dim IntRetCD
	
'Dim strPlantCd
Dim strItemCd
Dim strItemNm
Dim strClassCd
Dim strItemGroup
Dim strItemAcct
Dim strBaseDt
Dim strValidFlg
Dim strCharValueCd1
Dim strCharValueCd2
Dim strCharCd1
Dim strCharCd2

Dim TmpBuffer
Dim iTotalStr

strItemCd		= FilterVar(Trim(Request("txtItemCd")) ,"''", "S")
strItemNm		= FilterVar(Trim(Request("txtItemNm")) ,"''", "S")
strClassCd		= FilterVar(Trim(Request("txtClassCd"))	,"''", "S")
strItemGroup	= FilterVar(Trim(Request("txtItemGroup"))	,"''", "S")
strItemAcct		= FilterVar(Trim(Request("cboItemAccount"))	,"''", "S")
strValidFlg		= FilterVar(Trim(Request("rdoValidFlg"))	,"''", "S")
strCharValueCd1 = FilterVar(Trim(Request("txtCharValueCd1")) ,"''", "S")
strCharValueCd2 = FilterVar(Trim(Request("txtCharValueCd2")) ,"''", "S")
If Request("lgCurDate") <> "" Then
	strBaseDt	= FilterVar(UniConvDate(Request("lgCurDate"))	,"''", "S")
Else
	strBaseDt	= "''"
End If

lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
	
'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	
Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '☜: Query
		%>
		<Script Language=vbscript>
			parent.txtClassNm.Value = ""
			parent.txtCharValueNm1.Value = ""
			parent.txtCharValueNm2.Value = ""
		</Script>
		<%
		If strClassCd <> "''" Then Call SubBizClassNm()
		If strCharValueCd1 <> "''" Then Call SubBizCharValueDesc1()
		If strCharValueCd2 <> "''" Then Call SubBizCharValueDesc2()
		If strItemGroup <> "''" Then Call SubBizItemGroupNm()

		IF strItemNm = "''" Or (strItemCd <> "''" And strItemNm <> "''" ) Then
			Call SubBizQueryMulti("ITEM_CD")
		Else		
		    Call SubBizQueryMulti("ITEM_NM")
		End If
    Case ELSE
         Call SubBizLookup()

End Select
    
Call SubCloseDB(lgObjConn)

Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "With parent" & vbCrLf
		If Trim(lgErrorStatus) = "NO" And IntRetCd <> -1 Then
	        Response.Write ".ggoSpread.Source = .vspdData" & vbCrLf
			Response.Write ".lgStrPrevKeyIndex = """ & lgStrPrevKeyIndex & """" & vbCrLf
	        Response.Write ".ggoSpread.SSShowDataByClip " & """" & ConvSPChars(iTotalStr) & """" & vbCrLf
        End If
		
		' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 
		Response.Write "If .vspdData.MaxRows < .VisibleRowCnt(.vspdData,0)  And .lgStrPrevKeyIndex <> """" Then" & vbCrLf
			Response.Write ".DbQuery" & vbCrLf
		Response.Write "Else" & vbCrLf
			Response.Write ".DbQueryOk" & vbCrLf
		Response.Write "End If" & vbCrLf
		Response.Write ".vspdData.Focus" & vbCrLf
	Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
    
'============================================================================================================
' Name : SubBizLookup
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizLookup()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------   
    Call SubMakeSQLStatements("LK",strItemCd,strItemNm,strClassCd,strCharValueCd1,strCharValueCd2,strItemGroup,strItemAcct,strBaseDt,strValidFlg)           '☜ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
    Else
		%>
		<Script Language=vbscript>
			Parent.lgCharCd1 = "<%=ConvSPChars(lgObjRs("CHAR_CD1"))%>"
			Parent.lgCharCd2 = "<%=ConvSPChars(lgObjRs("CHAR_CD2"))%>"
		</Script>
		<%
		lgStrSQL = ""		
	    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
	End If

End Sub
    
'============================================================================================================
' Name : SubBizClassNm
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizClassNm()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------   
    Call SubMakeSQLStatements("LK",strItemCd,strItemNm,strClassCd,strCharValueCd1,strCharValueCd2,strItemGroup,strItemAcct,strBaseDt,strValidFlg)           '☜ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("122650", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
        Response.End
    Else
		strCharCd1 = FilterVar(lgObjRs("CHAR_CD1") ,"''", "S")
		strCharCd2 = FilterVar(lgObjRs("CHAR_CD2") ,"''", "S")
		%>
		<Script Language=vbscript>
			parent.txtClassNm.Value = "<%=ConvSPChars(lgObjRs("CLASS_NM"))%>"
			Parent.lgCharCd1 = "<%=ConvSPChars(lgObjRs("CHAR_CD1"))%>"
			Parent.lgCharCd2 = "<%=ConvSPChars(lgObjRs("CHAR_CD2"))%>"
		</Script>
		<%
		lgStrSQL = ""		
	    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
	End If

End Sub

'============================================================================================================
' Name : SubBizCharValueDesc1
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizCharValueDesc1()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	lgPrevNext = "1"

    '---------- Developer Coding part (Start) ---------------------------------------------------------------   
    Call SubMakeSQLStatements("L1",strItemCd,strItemNm,strClassCd,strCharValueCd1,strCharValueCd2,strItemGroup,strItemAcct,strBaseDt,strValidFlg)           '☜ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("122640", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
        Response.End
    Else
		%>
		<Script Language=vbscript>
			parent.txtCharValueNm1.Value = "<%=ConvSPChars(lgObjRs("CHAR_VALUE_NM"))%>"
		</Script>
		<%
		lgStrSQL = ""		
	    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
	End If

End Sub

'============================================================================================================
' Name : SubBizCharValueDesc2
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizCharValueDesc2()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------   
    Call SubMakeSQLStatements("L2",strItemCd,strItemNm,strClassCd,strCharValueCd1,strCharValueCd2,strItemGroup,strItemAcct,strBaseDt,strValidFlg)           '☜ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("122640", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
        Response.End
    Else
		%>
		<Script Language=vbscript>
			parent.txtCharValueNm2.Value = "<%=ConvSPChars(lgObjRs("CHAR_VALUE_NM"))%>"
		</Script>
		<%
		lgStrSQL = ""		
	    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
	End If

End Sub

'============================================================================================================
' Name : SubBizItemNm
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizItemCd()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------   
    Call SubMakeSQLStatements("LM",strItemCd,strItemNm,strClassCd,strCharValueCd1,strCharValueCd2,strItemGroup,strItemAcct,strBaseDt,strValidFlg)           '☜ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
    Else
		%>
		<Script Language=vbscript>
			parent.txtItemCd.Value = "<%=ConvSPChars(lgObjRs("ITEM_CD"))%>"
			parent.txtItemNM.Value = "<%=ConvSPChars(lgObjRs("ITEM_NM"))%>"
		</Script>
		<%
		lgStrSQL = ""		
	    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
	End If

End Sub

'============================================================================================================
' Name : SubBizItemNm
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizItemNm()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------   
    Call SubMakeSQLStatements("LI",strItemCd,strItemNm,strClassCd,strCharValueCd1,strCharValueCd2,strItemGroup,strItemAcct,strBaseDt,strValidFlg)           '☜ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
    Else
		%>
		<Script Language=vbscript>
			parent.txtItemNm.Value = "<%=ConvSPChars(lgObjRs("ITEM_NM"))%>"
		</Script>
		<%
		lgStrSQL = ""		
	    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
	End If

End Sub

'============================================================================================================
' Name : SubBizItemGroupNm
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizItemGroupNm()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------   
    Call SubMakeSQLStatements("LG",strItemCd,strItemNm,strClassCd,strCharValueCd1,strCharValueCd2,strItemGroup,strItemAcct,strBaseDt,strValidFlg)           '☜ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("127400", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		lgStrSQL = ""		
	    Call SubCloseRs(lgObjRs)
	    Response.End
	End If

End Sub
    
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti(pType)
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim PrntKey
	Dim NodX
	Dim Node
	Dim iDx	
	
	iDx = 0	    

	If pType = "ITEM_CD" Then
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
		Call SubMakeSQLStatements("MC",strItemCd,strItemNm,strClassCd,strCharValueCd1,strCharValueCd2,strItemGroup,strItemAcct,strBaseDt,strValidFlg)           '☜ : Make sql statements
	Else
		Call SubMakeSQLStatements("MN",strItemCd,strItemNm,strClassCd,strCharValueCd1,strCharValueCd2,strItemGroup,strItemAcct,strBaseDt,strValidFlg)           '☜ : Make sql statements
	End If
	
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		lgStrPrevKeyIndex = ""    
		IntRetCD = -1
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()
				 
		Response.End
    Else
		IntRetCD = 1
		ReDim TmpBuffer(0)
		Call SubSkipRs(lgObjRs, lgMaxCount * lgStrPrevKeyIndex )
        Do While Not lgObjRs.EOF
			
			lgstrData = ""
			lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("ITEM_CD"))			'품목코드 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("ITEM_NM")				'품목명 
	        lgstrData = lgstrData & Chr(11) & lgObjRs("SPEC")					'규격 
	        lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("CLASS_CD"))			'클래스 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("CLASS_NM")				'클래스명 
	        lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("CHAR_VALUE_CD1"))	'사양값1
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("CHAR_VALUE_NM1")			'사양값명1
	        lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("CHAR_VALUE_CD2"))	'사양값2
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("CHAR_VALUE_NM2")			'사양값명2	        
	        lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("BASIC_UNIT"))		'단위 
			lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("ITEM_ACCT"))		'계정 
			lgstrData = lgstrData & Chr(11) & Trim(lgObjRs("MINOR_NM_ITEM_ACCT"))		'계정명 
			lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("ITEM_GROUP_CD"))	'품목그룹 
	        lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("ITEM_GROUP_NM"))	'품목그룹명	
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("BASE_ITEM_CD")			'기준품목 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("BASE_ITEM_NM")			'기준품목명 
	        lgstrData = lgstrData & Chr(11)	& UNIDateClientFormat(lgObjRs("VALID_FROM_DT"))	'시작일 
	        lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("VALID_TO_DT"))	'종료일 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("UNIT_WEIGHT")			'단위중량 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("UNIT_OF_WEIGHT")			'중량단위 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("DRAW_NO")				'도면번호 
	        lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("HS_CD"))			'HS코드 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("HS_UNIT")				'HS단위 
	        lgstrData = lgstrData & Chr(11)	& Trim(lgObjRs("ITEM_IMAGE_FLG"))	'품목사진유무 
	        lgstrData = lgstrData & Chr(11)	& lgObjRs("FORMAL_NM")				'품목정식명칭		
	        lgstrData = lgstrData & Chr(11) & lgObjRs("VALID_FLG")				'유효구분 
		
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
	        lgstrData = lgstrData & Chr(11) & (lgMaxCount * lgStrPrevKeyIndex) + iDx
	        lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext
	        
	        ReDim Preserve TmpBuffer(iDx)
	        
	        TmpBuffer(iDx) = lgstrData
	         
	        iDx =  iDx + 1
	            
	        If iDx >= lgMaxCount Then			'처음에 최상위품목row를 한줄 써주었으므로 
	           lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
	           Exit Do
	        End If   
	                  
        Loop 
        
        iTotalStr = Join(TmpBuffer, "")
        
		If iDx < lgMaxCount Then
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
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2,pCode3,pCode4,pCode5,pCode6,pCode7,pCode8)
    Dim iSelCount
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case Mid(pDataType,1,1)
        Case "L"
			Select Case Mid(pDataType,2,1)
				Case "K"
				     lgStrSQL = "SELECT *   " 
				     lgStrSQL = lgStrSQL & " FROM  B_CLASS "
				     lgStrSQL = lgStrSQL & " WHERE CLASS_CD = " & pCode2
				Case "1"
				      If strCharCd1 <> "" Then
				  		lgStrSQL = "SELECT *   " 
				  		lgStrSQL = lgStrSQL & " FROM  B_CHAR_VALUE "
				  		lgStrSQL = lgStrSQL & " WHERE CHAR_CD = " & strCharCd1
				  		lgStrSQL = lgStrSQL & " AND	 CHAR_VALUE_CD = " & pCode3
				  	Else
				  		lgStrSQL = "SELECT Top 1 *   " 
				  		lgStrSQL = lgStrSQL & " FROM  B_CHAR_VALUE "
				  		lgStrSQL = lgStrSQL & " WHERE CHAR_VALUE_CD = " & pCode3
				  	End If
				Case "2"
				      If strCharCd2 <> "" Then
				  		lgStrSQL = "SELECT *   " 
				  		lgStrSQL = lgStrSQL & " FROM  B_CHAR_VALUE "
				  		lgStrSQL = lgStrSQL & " WHERE CHAR_CD = " & strCharCd2
				  		lgStrSQL = lgStrSQL & " AND	 CHAR_VALUE_CD = " & pCode4
				  	Else
				  		lgStrSQL = "SELECT Top 1 *   " 
				  		lgStrSQL = lgStrSQL & " FROM  B_CHAR_VALUE "
				  		lgStrSQL = lgStrSQL & " WHERE CHAR_VALUE_CD = " & pCode4
				  	End If
				Case "I"
				     lgStrSQL = "SELECT *   " 
				     lgStrSQL = lgStrSQL & " FROM  B_ITEM "
				     lgStrSQL = lgStrSQL & " WHERE ITEM_CD = " & pCode
				Case "M"
				     lgStrSQL = "SELECT *   " 
				     lgStrSQL = lgStrSQL & " FROM  B_ITEM "
				     lgStrSQL = lgStrSQL & " WHERE ITEM_NM = " & pCode1
                  Case "G"
                       lgStrSQL = "SELECT *   " 
                       lgStrSQL = lgStrSQL & " FROM  B_ITEM_GROUP "
                       lgStrSQL = lgStrSQL & " WHERE ITEM_GROUP_CD = " & pCode5
				     
			End Select
        Case "M"
			lgStrSQL = "SELECT A.*, D.CLASS_NM, E.CHAR_VALUE_NM CHAR_VALUE_NM1, F.CHAR_VALUE_NM CHAR_VALUE_NM2, B.ITEM_NM BASE_ITEM_NM, A.ITEM_ACCT, dbo.ufn_GetCodeName('P1001', A.ITEM_ACCT) MINOR_NM_ITEM_ACCT, C.ITEM_GROUP_NM"
			lgStrSQL = lgStrSQL & " FROM B_ITEM A, B_ITEM B, B_ITEM_GROUP C, B_CLASS D, B_CHAR_VALUE E, B_CHAR_VALUE F"
			lgStrSQL = lgStrSQL & " WHERE A.BASE_ITEM_CD *= B.ITEM_CD AND A.ITEM_GROUP_CD *= C.ITEM_GROUP_CD "
			lgStrSQL = lgStrSQL & " AND A.CLASS_CD = D.CLASS_CD AND D.CHAR_CD1 = E.CHAR_CD AND D.CHAR_CD2 *= F.CHAR_CD"
			lgStrSQL = lgStrSQL & " AND A.CHAR_VALUE_CD1 = E.CHAR_VALUE_CD AND A.CHAR_VALUE_CD2 *= F.CHAR_VALUE_CD "
			If pCode <> "''" Then
				pCode = "'%" & Trim(Request("txtItemCd")) & "%'"
				lgStrSQL = lgStrSQL & " AND A.ITEM_CD like " & pCode
			End If
			If pCode1 <> "''" Then
				pCode1 = "'%" & Trim(Request("txtItemNm")) & "%'"
				lgStrSQL = lgStrSQL & " AND A.ITEM_NM like " & pCode1
			End If
			If pCode2 <> "''" Then	lgStrSQL = lgStrSQL & " AND A.CLASS_CD = " & pCode2
			If pCode3 <> "''" Then	lgStrSQL = lgStrSQL & " AND A.CHAR_VALUE_CD1 >= " & pCode3
			If pCode4 <> "''" Then	lgStrSQL = lgStrSQL & " AND A.CHAR_VALUE_CD2 >= " & pCode4
			If pCode5 <> "''" Then	lgStrSQL = lgStrSQL & " AND A.ITEM_GROUP_CD = " & pCode5
			If pCode6 <> "''" Then	lgStrSQL = lgStrSQL & " AND A.ITEM_ACCT = " & pCode6
'			If pCode7 <> "''" Then	lgStrSQL = lgStrSQL & " AND A.VALID_FROM_DT <= " & pCode7
			If pCode7 <> "''" Then	lgStrSQL = lgStrSQL & " AND A.VALID_TO_DT >= " & pCode7
			If pCode8 <> "''" Then	lgStrSQL = lgStrSQL & " AND A.VALID_FLG = " & pCode8
        
			Select Case Mid(pDataType,2,1)        
			Case "C"
				lgStrSQL = lgStrSQL & " ORDER BY A.ITEM_CD, A.ITEM_NM, A.CLASS_CD " 
			Case "N"
				lgStrSQL = lgStrSQL & " ORDER BY A.ITEM_NM, A.ITEM_CD, A.CLASS_CD " 
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
