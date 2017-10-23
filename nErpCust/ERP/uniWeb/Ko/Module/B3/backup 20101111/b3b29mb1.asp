<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call LoadBasisGlobalInf
Call loadInfTB19029B("Q", "P", "NOCOOKIE","MB")

    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    '---------------------------------------Common-----------------------------------------------------------
    Const C_SHEETMAXROWS_D = 50
    
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgMaxCount        = C_SHEETMAXROWS_D				                                  '☜: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Dim TmpBuffer
	Dim iTotalStr
	
	Dim strItemCd
	Dim strClassCd
	Dim strCharValueCd1
	Dim strCharValueCd2
	Dim	strCharCd1
	Dim	strCharCd2
	Dim	strItemAcct
	Dim	strItemGroup
	Dim	strItemValid
	Dim	strBaseDt
	
	strItemCd		= FilterVar(Request("txtItemCd"), "''", "S")
	strClassCd		= FilterVar(Request("txtClassCd"), "''", "S")
	strCharValueCd1 = FilterVar(Request("txtCharValueCd1"), "''", "S")
	strCharValueCd2 = FilterVar(Request("txtCharValueCd2"), "''", "S")
	strItemAcct		= FilterVar(Request("cboItemAcct"), "''", "S")
	strItemGroup	= FilterVar(Request("txtItemGroupCd"), "''", "S")
	strItemValid	= FilterVar(Request("rdoDefaultFlg"), "''", "S")
	If Request("txtBaseDt") <> "" Then
		strBaseDt	= FilterVar(Trim(UNIConvDate(Request("txtBaseDt"))),"" & FilterVar("1900-01-01", "''", "S") & " ", "S")
	Else
		strBaseDt	= "''"
	End If
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
             Call SubBizSaveMulti()
        Case ELSE
             Call SubBizLookup()

    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizLookup
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizLookup()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	lgPrevNext = "L"

    '---------- Developer Coding part (Start) ---------------------------------------------------------------   
    Call SubMakeSQLStatements("S",strItemCd,strClassCd,strCharValueCd1,strCharValueCd2,strItemAcct,strItemGroup,strItemValid,strBaseDt)
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("122650", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else
		%>
		<Script Language=vbscript>
			parent.lgCharCd1 = "<%=ConvSPChars(lgObjRs("CHAR_CD1"))%>"
			parent.lgCharCd2 = "<%=ConvSPChars(lgObjRs("CHAR_CD2"))%>"
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

	lgPrevNext = "L"
    '---------- Developer Coding part (Start) ---------------------------------------------------------------   
    Call SubMakeSQLStatements("S",strItemCd,strClassCd,strCharValueCd1,strCharValueCd2,strItemAcct,strItemGroup,strItemValid,strBaseDt)
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
   		%>
		<Script Language=vbscript>
			parent.lgLocalModeFlag = FALSE		
			parent.frm1.txtClassNm.Value = ""
		</Script>
		<%
        Call DisplayMsgBox("122650", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
		Response.End
    Else
		strCharCd1 = FilterVar(lgObjRs("CHAR_CD1"), "''", "S")
		strCharCd2 = FilterVar(lgObjRs("CHAR_CD2"), "''", "S")
		%>
		<Script Language=vbscript>
			parent.frm1.txtClassNm.Value = "<%=ConvSPChars(lgObjRs("CLASS_NM"))%>"
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
    Call SubMakeSQLStatements("S",strItemCd,strClassCd,strCharValueCd1,strCharValueCd2,strItemAcct,strItemGroup,strItemValid,strBaseDt)
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
       	%>
		<Script Language=vbscript>
			parent.lgLocalModeFlag = FALSE		
			parent.frm1.txtCharValueNm1.Value = ""
		</Script>
		<%
    Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtCharValueNm1.Value = "<%=ConvSPChars(lgObjRs("CHAR_VALUE_NM"))%>"
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

	lgPrevNext = "2"

    '---------- Developer Coding part (Start) ---------------------------------------------------------------   
    Call SubMakeSQLStatements("S",strItemCd,strClassCd,strCharValueCd1,strCharValueCd2,strItemAcct,strItemGroup,strItemValid,strBaseDt)
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
       	%>
		<Script Language=vbscript>
			parent.lgLocalModeFlag = FALSE		
			parent.frm1.txtCharValueNm2.Value = ""
		</Script>
		<%
    Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtCharValueNm2.Value = "<%=ConvSPChars(lgObjRs("CHAR_VALUE_NM"))%>"
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

	lgPrevNext = "I"

    '---------- Developer Coding part (Start) ---------------------------------------------------------------   
    Call SubMakeSQLStatements("S",strItemCd,strClassCd,strCharValueCd1,strCharValueCd2,strItemAcct,strItemGroup,strItemValid,strBaseDt)
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		%>
		<Script Language=vbscript>
			parent.frm1.txtItemNm.Value = ""
		</Script>
		<%
    Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtItemNm.Value = "<%=ConvSPChars(lgObjRs("ITEM_NM"))%>"
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

	lgPrevNext = "G"

    '---------- Developer Coding part (Start) ---------------------------------------------------------------   
    Call SubMakeSQLStatements("S",strItemCd,strClassCd,strCharValueCd1,strCharValueCd2,strItemAcct,strItemGroup,strItemValid,strBaseDt)
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("127400", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		%>
		<Script Language=vbscript>
			parent.frm1.txtItemGroupNm.Value = ""
		</Script>
		<%
		lgStrSQL = ""		
	    Call SubCloseRs(lgObjRs)
	    Response.End
    Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtItemGroupNm.Value = "<%=ConvSPChars(lgObjRs("ITEM_GROUP_NM"))%>"
		</Script>
		<%
		lgStrSQL = ""		
	    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
	End If

End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    If strClassCd <> "''" Then
		Call SubBizClassNm()
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtClassNm.Value = ""
		</Script>
		<%
	End If
    If strCharValueCd1 <> "''" Then
		Call SubBizCharValueDesc1()
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtCharValueNm1.Value = ""
		</Script>
		<%
	End If
    If strCharValueCd2 <> "''" Then
		Call SubBizCharValueDesc2()
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtCharValueNm2.Value = ""
		</Script>
		<%
	End If
    If strItemCd <> "''" Then
		Call SubBizItemNm()
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtItemNm.Value = ""
		</Script>
		<%
	End If
    If strItemGroup <> "''" Then
		Call SubBizItemGroupNm()
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtItemGroupNm.Value = ""
		</Script>
		<%
	End If	
    Call SubBizQueryMulti()

	%>
	<Script Language=vbscript>
		parent.frm1.hItemCd.Value			= "<%=Request("txtItemCd")%>"
		parent.frm1.hItemAcct.Value			= "<%=Request("cboItemAcct")%>"
		parent.frm1.hItemGroupCd.Value		= "<%=Request("txtItemGroupCd")%>"
		parent.frm1.hAvailableItem.Value	= "<%=Request("rdoDefaultFlg")%>"
		parent.frm1.hBaseDt.Value			= "<%=Request("txtBaseDt")%>"
		parent.frm1.hClassCd.Value			= "<%=Request("txtClassCd")%>"
		parent.frm1.hCharValueCd1.Value		= "<%=Request("txtCharValueCd1")%>"
		parent.frm1.hCharValueCd2.Value		= "<%=Request("txtCharValueCd2")%>"
	</Script>
	<%

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------   
    Call SubMakeSQLStatements("MR",strItemCd,strClassCd,strCharValueCd1,strCharValueCd2,strItemAcct,strItemGroup,strItemValid,strBaseDt)
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

        ReDim TmpBuffer(0)
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF
            
            lgstrData = ""     
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_cd"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_nm"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("spec"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("class_cd"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("class_nm"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("char_value_cd1"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("char_value_nm1"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("char_value_cd2"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("char_value_nm2"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("basic_unit"))			
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_acct"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_acct"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_group_cd"))						
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_group_nm"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("base_item_cd"))						
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("base_item_nm"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("valid_from_dt"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("valid_to_dt"))			
			lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("unit_weight"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("unit_of_weight"))
			lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("GROSS_WEIGHT"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GROSS_UNIT"))
			lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("CBM"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CBM_DESCRIPTION"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("draw_no"))			
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("hs_cd"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("hs_unit"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_image_flg"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("formal_nm"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("valid_flg"))

	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext
			
			ReDim Preserve TmpBuffer(iDx-1)
			
			TmpBuffer(iDx-1) = lgstrData
            iDx =  iDx + 1
            If iDx > lgMaxCount Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
               Exit Do
            End If   
               
        Loop 
    End If
    
    iTotalStr = Join(TmpBuffer, "")
    
    If iDx <= lgMaxCount Then
       lgStrPrevKeyIndex = ""
    End If   

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
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

    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2,pCode3,pCode4,pCode5,pCode6,pCode7)
    Dim iSelCount

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case Mid(pDataType,1,1)
        Case "S"
	       Select Case  lgPrevNext 
                  Case " "
                  Case "L"
                       lgStrSQL = "SELECT *   " 
                       lgStrSQL = lgStrSQL & " FROM  B_CLASS "
                       lgStrSQL = lgStrSQL & " WHERE CLASS_CD = " & pCode1
                  Case "1"
					    If strCharCd1 <> "" Then
							lgStrSQL = "SELECT *   " 
							lgStrSQL = lgStrSQL & " FROM  B_CHAR_VALUE "
							lgStrSQL = lgStrSQL & " WHERE CHAR_CD = " & strCharCd1
							lgStrSQL = lgStrSQL & " AND	 CHAR_VALUE_CD = " & pCode2
						Else
							lgStrSQL = "SELECT Top 1 *   " 
							lgStrSQL = lgStrSQL & " FROM  B_CHAR_VALUE "
							lgStrSQL = lgStrSQL & " WHERE CHAR_VALUE_CD = " & pCode2
						End If
                  Case "2"
					    If strCharCd2 <> "" Then
							lgStrSQL = "SELECT *   " 
							lgStrSQL = lgStrSQL & " FROM  B_CHAR_VALUE "
							lgStrSQL = lgStrSQL & " WHERE CHAR_CD = " & strCharCd2
							lgStrSQL = lgStrSQL & " AND	 CHAR_VALUE_CD = " & pCode3
						Else
							lgStrSQL = "SELECT Top 1 *   " 
							lgStrSQL = lgStrSQL & " FROM  B_CHAR_VALUE "
							lgStrSQL = lgStrSQL & " WHERE CHAR_VALUE_CD = " & pCode3
						End If

                  Case "I"
                       lgStrSQL = "SELECT *   " 
                       lgStrSQL = lgStrSQL & " FROM  B_ITEM "
                       lgStrSQL = lgStrSQL & " WHERE ITEM_CD = " & pCode
                  Case "G"
                       lgStrSQL = "SELECT *   " 
                       lgStrSQL = lgStrSQL & " FROM  B_ITEM_GROUP "
                       lgStrSQL = lgStrSQL & " WHERE ITEM_GROUP_CD = " & pCode5
                       
           End Select
        Case "M"
        
           iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1
           
           Select Case Mid(pDataType,2,1)
               Case "C"
                       lgStrSQL = "SELECT *   " 
                       lgStrSQL = lgStrSQL & " FROM  B_MAJOR "
                       lgStrSQL = lgStrSQL & " WHERE MAJOR_CD " & pComp & pCode
               Case "D"
                       lgStrSQL = "SELECT *   " 
                       lgStrSQL = lgStrSQL & " FROM  B_MAJOR "
                       lgStrSQL = lgStrSQL & " WHERE MAJOR_CD " & pComp & pCode
               Case "R"
					lgStrSQL = " SELECT		a.item_cd,a.item_nm,a.formal_nm,a.spec,dbo.ufn_GetCodeName(" & FilterVar("P1001", "''", "S") & " ,a.item_acct) item_acct, "
					lgStrSQL = lgStrSQL & "	a.item_class, a.hs_cd,a.hs_unit, a.unit_weight, a.unit_of_weight, a.gross_weight, a.gross_unit, "
					lgStrSQL = lgStrSQL & "	a.cbm, a.cbm_description, a.basic_unit,a.phantom_flg,a.draw_no,a.item_image_flg, "
					lgStrSQL = lgStrSQL & "	a.blanket_pur_flg,a.base_item_cd,a.item_group_cd,a.proportion_rate,a.valid_flg, "
					lgStrSQL = lgStrSQL & "	a.valid_from_dt,a.valid_to_dt,a.vat_type,a.vat_rate,a.class_flg,a.class_cd, "
					lgStrSQL = lgStrSQL & "	a.char_value_cd1,a.char_value_cd2,b.item_nm base_item_nm, "
					lgStrSQL = lgStrSQL & "	c.item_group_nm,d.class_nm,d.char_value_nm char_value_nm1,e.char_value_nm char_value_nm2 "
					lgStrSQL = lgStrSQL & "	from	b_item a,	b_item b,	b_item_group c, "
					lgStrSQL = lgStrSQL & "		(select	a.class_cd, a.class_nm, b.char_value_cd, b.char_value_nm "
					lgStrSQL = lgStrSQL & "		from	b_class a, b_char_value b "
					lgStrSQL = lgStrSQL & "		where	a.char_cd1 *= b.char_cd "
					lgStrSQL = lgStrSQL & "		union "
					lgStrSQL = lgStrSQL & "		select	a.class_cd, a.class_nm, b.char_value_cd, b.char_value_nm "
					lgStrSQL = lgStrSQL & "		from	b_class a, b_char_value b "
					lgStrSQL = lgStrSQL & "		where	a.char_cd2 *= b.char_cd ) d, "
					lgStrSQL = lgStrSQL & "		(select	a.class_cd, a.class_nm, b.char_value_cd, b.char_value_nm "
					lgStrSQL = lgStrSQL & "		from	b_class a, b_char_value b "
					lgStrSQL = lgStrSQL & "		where	a.char_cd1 *= b.char_cd "
					lgStrSQL = lgStrSQL & "		union "
					lgStrSQL = lgStrSQL & "		select	a.class_cd, a.class_nm, b.char_value_cd, b.char_value_nm "
					lgStrSQL = lgStrSQL & "		from	b_class a, b_char_value b "
					lgStrSQL = lgStrSQL & "		where	a.char_cd2 *= b.char_cd ) e "					
					lgStrSQL = lgStrSQL & "	where	a.base_item_cd *= b.item_cd "
					lgStrSQL = lgStrSQL & "	and	a.item_group_cd *= c.item_group_cd "
					lgStrSQL = lgStrSQL & "	and	a.class_cd *= d.class_cd "
					lgStrSQL = lgStrSQL & "	and	a.char_value_cd1 *= d.char_value_cd "
					lgStrSQL = lgStrSQL & "	and	a.class_cd *= e.class_cd "
					lgStrSQL = lgStrSQL & "	and	a.char_value_cd2 *= e.char_value_cd "
					
					If pCode <> "''" Then
						pCode = " " & FilterVar("%" & Trim(Request("txtItemCd")) & "%", "''", "S") & ""
						lgStrSQL = lgStrSQL & "	and	a.item_cd LIKE " & pCode
					End If
					If pCode1 <> "''" Then
						lgStrSQL = lgStrSQL & "	and	a.class_cd = " & pCode1
					End If
					If pCode2 <> "''" Then
						lgStrSQL = lgStrSQL & "	and	a.char_value_cd1 >= " & pCode2
					End If
					If pCode3 <> "''" Then
						lgStrSQL = lgStrSQL & "	and	a.char_value_cd2 >= " & pCode3
					End If
					If pCode4 <> "''" Then
						lgStrSQL = lgStrSQL & "	and	a.item_acct = " & pCode4
					End If
					If pCode5 <> "''" Then
						lgStrSQL = lgStrSQL & "	and a.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & pCode5 & " ))"
					End If
					If pCode6 <> "''" Then
						lgStrSQL = lgStrSQL & "	and	a.valid_flg = " & pCode6
					End If
					If pCode7 <> "''" Then
'						lgStrSQL = lgStrSQL & "	and	a.valid_from_dt <= " & pCode7
						lgStrSQL = lgStrSQL & "	and	a.valid_to_dt >= " & pCode7
					End If

               Case "U"
                       lgStrSQL = "SELECT *   " 
                       lgStrSQL = lgStrSQL & " FROM  B_MAJOR "
                       lgStrSQL = lgStrSQL & " WHERE MAJOR_CD " & pComp & pCode
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
                .ggoSpread.SSShowDataByClip "<%=iTotalStr%>"
                .DBQueryOk()
                
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
    
       
</Script>	

