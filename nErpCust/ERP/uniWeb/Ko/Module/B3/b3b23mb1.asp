<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    '---------------------------------------Common-----------------------------------------------------------
    
    Const C_SHEETMAXROWS_D    = 50
    
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgMaxCount        = C_SHEETMAXROWS_D		                                  '☜: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = CInt(Request("lgStrPrevKeyIndex"))                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Dim strClassCd
	Dim strCharValueCd1
	Dim strCharValueCd2
	Dim	strCharCd1
	Dim	strCharCd2
	
	Dim TmpBuffer
	Dim iTotalStr

	strClassCd		= FilterVar(Request("txtClassCd"), "''", "S")
	strCharValueCd1 = FilterVar(Request("txtCharValueCd1"), "''", "S")
	strCharValueCd2 = FilterVar(Request("txtCharValueCd2"), "''", "S")

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
    Call SubMakeSQLStatements("S",strClassCd,strCharValueCd1,strCharValueCd2)                                 '☆ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("122650", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
        Response.End
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
    Call SubMakeSQLStatements("S",strClassCd,strCharValueCd1,strCharValueCd2)                                 '☆ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("122650", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		%>
		<Script Language=vbscript>
			parent.lgLocalModeFlag = FALSE
		</Script>
		<%
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
    Call SubMakeSQLStatements("S",strClassCd,strCharValueCd1,strCharValueCd2)                                 '☆ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        'Call SetErrorStatus()
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
    Call SubMakeSQLStatements("S",strClassCd,strCharValueCd1,strCharValueCd2)                                 '☆ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        'Call SetErrorStatus()
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
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
	%>
	<Script Language=vbscript>
		parent.frm1.txtClassNm.Value = ""
		parent.frm1.txtCharValueNm1.Value = ""
		parent.frm1.txtCharValueNm2.Value = ""
	</Script>
	<%
    
    Call SubBizClassNm()
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
    Call SubBizQueryMulti()
	%>
	<Script Language=vbscript>
		parent.frm1.hClassCd.Value = "<%=Request("txtClassCd")%>"
		parent.frm1.hCharValueCd1.Value = "<%=Request("txtCharValueCd1")%>"
		parent.frm1.hCharValueCd2.Value = "<%=Request("txtCharValueCd2")%>"
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
    Call SubMakeSQLStatements("MR",strClassCd,strCharValueCd1,strCharValueCd2)                                 '☆ : Make sql statements
    
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
			
			lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_cd"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CHAR_VALUE_CD11"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CHAR_VALUE_NM11"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CHAR_VALUE_CD22"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CHAR_VALUE_NM22"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_formal_nm"))
            lgstrData = lgstrData & Chr(11) & "N"
            lgstrData = lgstrData & Chr(11) & "N"
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & 0
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
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & 0
            
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

			ReDim Preserve TmpBuffer(iDx-1)
			TmpBuffer(iDx-1) = lgstrData
			
		    lgObjRs.MoveNext

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
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2)
    Dim iSelCount

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case Mid(pDataType,1,1)
        Case "S"
	       Select Case  lgPrevNext 
                  Case " "
                  Case "P"
                  Case "N"
                  Case "L"
                       lgStrSQL = "SELECT *   " 
                       lgStrSQL = lgStrSQL & " FROM  B_CLASS "
                       lgStrSQL = lgStrSQL & " WHERE CLASS_CD = " & pCode
                  Case "1"
                       lgStrSQL = "SELECT *   " 
                       lgStrSQL = lgStrSQL & " FROM  B_CHAR_VALUE "
                       lgStrSQL = lgStrSQL & " WHERE CHAR_CD = " & strCharCd1
                       lgStrSQL = lgStrSQL & " AND	 CHAR_VALUE_CD = " & pCode1
                       
'Call ServerMesgBox(lgStrSQL, vbInformation, I_MKSCRIPT)	'⊙: 에러내용, 메세지타입, 스크립트유형   
                       
                  Case "2"
                       lgStrSQL = "SELECT *   " 
                       lgStrSQL = lgStrSQL & " FROM  B_CHAR_VALUE "
                       lgStrSQL = lgStrSQL & " WHERE CHAR_CD = " & strCharCd2
                       lgStrSQL = lgStrSQL & " AND	 CHAR_VALUE_CD = " & pCode2
                       
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
'					lgStrSQL = "select	aaa.CHAR_VALUE_CD11, aaa.CHAR_VALUE_CD22, aaa.CHAR_VALUE_NM11, aaa.CHAR_VALUE_NM22, "
'					lgStrSQL = lgStrSQL & " case when isNULL(bbb.db_item_cd,'') = '' then aaa.item_cd else bbb.db_item_cd end item_cd, "
'					lgStrSQL = lgStrSQL & " case when isNULL(bbb.db_item_nm,'') = '' then aaa.item_nm else bbb.db_item_nm end item_nm, "
'					lgStrSQL = lgStrSQL & " case when isNULL(bbb.db_item_cd,'') = '' then 'N' else 'Y'end item_flag, "
'					lgStrSQL = lgStrSQL & " bbb.item_by_plant_cd,bbb.formal_nm,bbb.spec,bbb.item_acct,bbb.item_class,bbb.hs_cd,bbb.hs_unit, "
'					lgStrSQL = lgStrSQL & " bbb.unit_weight,bbb.unit_of_weight,bbb.basic_unit,bbb.phantom_flg,bbb.draw_no,bbb.item_image_flg, "
'					lgStrSQL = lgStrSQL & " bbb.blanket_pur_flg,bbb.base_item_cd,bbb.item_group_cd,bbb.proportion_rate,bbb.valid_flg, "
'					lgStrSQL = lgStrSQL & " bbb.valid_from_dt,bbb.valid_to_dt,bbb.vat_type,bbb.vat_rate,bbb.class_cd,bbb.class_flg "
'					lgStrSQL = lgStrSQL & " from "
'					lgStrSQL = lgStrSQL & "			(select	aa.CHAR_VALUE_CD11,	bb.CHAR_VALUE_CD22, "
'					lgStrSQL = lgStrSQL & 			pCode & " + '-' + aa.CHAR_VALUE_CD11 + '-' + bb.CHAR_VALUE_CD22 item_cd, "
'					lgStrSQL = lgStrSQL & "			aa.classdesc + '-' + aa.CHAR_VALUE_NM11 + '-' + bb.CHAR_VALUE_NM22 item_nm, "
'					lgStrSQL = lgStrSQL & "			aa.CHAR_VALUE_NM11, bb.CHAR_VALUE_NM22 "				
'					lgStrSQL = lgStrSQL & "			from "
'					lgStrSQL = lgStrSQL & "					(select 	a.CHAR_VALUE_CD CHAR_VALUE_CD11, a.CHAR_VALUE_NM CHAR_VALUE_NM11, b.CLASS_NM classdesc "
'					lgStrSQL = lgStrSQL & "			 		from	B_CHAR_VALUE a, B_CLASS b "
'					lgStrSQL = lgStrSQL & "			 		where	a.CHAR_CD = b.CHAR_CD1 "
'					lgStrSQL = lgStrSQL & "			 		and	b.class_cd = " & pCode & " ) aa, "
'					lgStrSQL = lgStrSQL & "					(select 	a.CHAR_VALUE_CD CHAR_VALUE_CD22,  a.CHAR_VALUE_NM CHAR_VALUE_NM22 "	
'					lgStrSQL = lgStrSQL & "					from	B_CHAR_VALUE a,	B_CLASS b "	
'					lgStrSQL = lgStrSQL & "					where	a.CHAR_CD = b.CHAR_CD2 "	
'					lgStrSQL = lgStrSQL & "					and	b.class_cd = " & pCode & " ) bb) aaa, "
'					lgStrSQL = lgStrSQL & "			(select	a.CHAR_VALUE_CD1,a.CHAR_VALUE_CD2,a.item_cd db_item_cd,a.item_nm db_item_nm, "
'					lgStrSQL = lgStrSQL & "			b.item_cd item_by_plant_cd,a.formal_nm,a.spec,a.item_acct,a.item_class,a.hs_cd, "
'					lgStrSQL = lgStrSQL & "			a.hs_unit,a.unit_weight,a.unit_of_weight,a.basic_unit,a.phantom_flg,a.draw_no,a.item_image_flg, "
'					lgStrSQL = lgStrSQL & "			a.blanket_pur_flg,a.base_item_cd,a.item_group_cd,a.proportion_rate,a.valid_flg,a.valid_from_dt, "
'					lgStrSQL = lgStrSQL & "			a.valid_to_dt,a.vat_type,a.vat_rate,a.class_cd,a.class_flg "
'					lgStrSQL = lgStrSQL & "			from	b_item a, b_item_by_plant b "
'					lgStrSQL = lgStrSQL & "			where	a.class_cd = " & pCode
'					lgStrSQL = lgStrSQL & "			and	isNULL(a.CHAR_VALUE_CD1,'') <> '' and isNULL(a.CHAR_VALUE_CD2,'') <> '' "
'					lgStrSQL = lgStrSQL & "			and	a.item_cd *= b.item_cd "
'					lgStrSQL = lgStrSQL & "			group by a.CHAR_VALUE_CD1, a.CHAR_VALUE_CD2, a.item_cd, a.item_nm, b.item_cd, a.formal_nm, a.spec, "
'					lgStrSQL = lgStrSQL & "			a.item_acct, a.item_class, a.hs_cd, a.hs_unit, a.unit_weight, a.unit_of_weight, "
'					lgStrSQL = lgStrSQL & "			a.basic_unit, a.phantom_flg, a.draw_no, a.item_image_flg, a.blanket_pur_flg,  "
'					lgStrSQL = lgStrSQL & "			a.base_item_cd, a.item_group_cd, a.proportion_rate, a.valid_flg, a.valid_from_dt, a.valid_to_dt,  "
'					lgStrSQL = lgStrSQL & "			a.vat_type, a.vat_rate, a.class_cd, a.class_flg) bbb "
'					lgStrSQL = lgStrSQL & "	where	aaa.CHAR_VALUE_CD11 *= bbb.CHAR_VALUE_CD1 "
'					lgStrSQL = lgStrSQL & "	and	aaa.CHAR_VALUE_CD22 *= bbb.CHAR_VALUE_CD2 "
'					lgStrSQL = lgStrSQL & "	order by 1, 2 "

					If strCharCd2 <> "''" Then
						lgStrSQL = "			select	aaa.CHAR_VALUE_CD11,aaa.CHAR_VALUE_CD22, "
						lgStrSQL = lgStrSQL & " aaa.CHAR_VALUE_NM11,aaa.CHAR_VALUE_NM22,aaa.item_cd,aaa.item_nm,aaa.item_formal_nm "
						lgStrSQL = lgStrSQL & " from	(select	aa.CHAR_VALUE_CD11,	bb.CHAR_VALUE_CD22,aa.CHAR_VALUE_NM11,bb.CHAR_VALUE_NM22, "
						lgStrSQL = lgStrSQL & "	Left( " & pCode & " + " & FilterVar("-", "''", "S") & " + aa.CHAR_VALUE_CD11 + " & FilterVar("-", "''", "S") & " + bb.CHAR_VALUE_CD22,18) item_cd, "
						lgStrSQL = lgStrSQL & "	Left( 	RTrim(aa.classdesc) + " & FilterVar("-", "''", "S") & " + RTrim(aa.CHAR_VALUE_NM11) + " & FilterVar("-", "''", "S") & " + RTrim(bb.CHAR_VALUE_NM22),40) item_nm, "
						lgStrSQL = lgStrSQL & "	Left( 	RTrim(aa.classdesc) + " & FilterVar("-", "''", "S") & " + RTrim(aa.CHAR_VALUE_NM11) + " & FilterVar("-", "''", "S") & " + RTrim(bb.CHAR_VALUE_NM22),50) item_formal_nm "
						lgStrSQL = lgStrSQL & "		from "
						lgStrSQL = lgStrSQL & "		(select a.CHAR_VALUE_CD CHAR_VALUE_CD11, a.CHAR_VALUE_NM CHAR_VALUE_NM11, b.CLASS_NM classdesc "
						lgStrSQL = lgStrSQL & "		from	B_CHAR_VALUE a,	B_CLASS b "
						lgStrSQL = lgStrSQL & "		where	a.CHAR_CD = b.CHAR_CD1 "

						If pCode1 = "" Then
							lgStrSQL = lgStrSQL & "		and	b.class_cd = " & pCode & " ) aa, "	
						Else
							lgStrSQL = lgStrSQL & "		and	b.class_cd = " & pCode
							lgStrSQL = lgStrSQL & "		and	a.char_value_cd >= " & pCode1 & " ) aa, "	
						End If
		
						lgStrSQL = lgStrSQL & "		(select a.CHAR_VALUE_CD CHAR_VALUE_CD22,  a.CHAR_VALUE_NM CHAR_VALUE_NM22 "
						lgStrSQL = lgStrSQL & "		from	B_CHAR_VALUE a,	B_CLASS b "
						lgStrSQL = lgStrSQL & "		where	a.CHAR_CD = b.CHAR_CD2 "

						If pCode2 = "" Then
							lgStrSQL = lgStrSQL & "		and	b.class_cd = " & pCode & ") bb) aaa "
						Else
							lgStrSQL = lgStrSQL & "		and	b.class_cd = " & pCode
							lgStrSQL = lgStrSQL & "		and	a.char_value_cd >= " & pCode2 & ") bb) aaa "
						End If

'						lgStrSQL = lgStrSQL & " where	aaa.CHAR_VALUE_CD11 not in "
'						lgStrSQL = lgStrSQL & " 		(select	a.CHAR_VALUE_CD1 "
'						lgStrSQL = lgStrSQL & " 		from	b_item a, b_item_by_plant b "
'						lgStrSQL = lgStrSQL & " 		where	a.class_cd = " & pCode
'						lgStrSQL = lgStrSQL & " 		and	isNULL(a.CHAR_VALUE_CD1,'') <> '' and isNULL(a.CHAR_VALUE_CD2,'') <> '' "
'						lgStrSQL = lgStrSQL & " 		and	a.item_cd *= b.item_cd "
'						lgStrSQL = lgStrSQL & " 		group by a.CHAR_VALUE_CD1) "
'						lgStrSQL = lgStrSQL & " or	aaa.CHAR_VALUE_CD22 not in "
'						lgStrSQL = lgStrSQL & " 		(select	a.CHAR_VALUE_CD2 "
'						lgStrSQL = lgStrSQL & " 		from	b_item a, b_item_by_plant b "
'						lgStrSQL = lgStrSQL & " 		where	a.class_cd = " & pCode
'						lgStrSQL = lgStrSQL & " 		and	isNULL(a.CHAR_VALUE_CD1,'') <> '' and isNULL(a.CHAR_VALUE_CD2,'') <> '' "
'						lgStrSQL = lgStrSQL & " 		and	a.item_cd *= b.item_cd "
'						lgStrSQL = lgStrSQL & " 		group by a.CHAR_VALUE_CD2) "
						lgStrSQL = lgStrSQL & " where	convert(varchar(18),aaa.item_cd) not in (select item_cd from b_item where class_cd = " & pCode & " ) "
						lgStrSQL = lgStrSQL & " order by 1, 2 "
					Else
						lgStrSQL = "			select	aaa.CHAR_VALUE_CD11,aaa.CHAR_VALUE_CD22, "
						lgStrSQL = lgStrSQL & " aaa.CHAR_VALUE_NM11,aaa.CHAR_VALUE_NM22,aaa.item_cd,aaa.item_nm,aaa.item_formal_nm "
						lgStrSQL = lgStrSQL & " from	(select	aa.CHAR_VALUE_CD11,	aa.CHAR_VALUE_CD22,aa.CHAR_VALUE_NM11,aa.CHAR_VALUE_NM22, "
						lgStrSQL = lgStrSQL & "	Left( " & pCode & " + " & FilterVar("-", "''", "S") & " + aa.CHAR_VALUE_CD11,18) item_cd, "
						lgStrSQL = lgStrSQL & "	Left( 	RTrim(aa.classdesc) + " & FilterVar("-", "''", "S") & " + RTrim(aa.CHAR_VALUE_NM11),40) item_nm, "
						lgStrSQL = lgStrSQL & "	Left( 	RTrim(aa.classdesc) + " & FilterVar("-", "''", "S") & " + RTrim(aa.CHAR_VALUE_NM11),50) item_formal_nm "
						lgStrSQL = lgStrSQL & "		from "
						lgStrSQL = lgStrSQL & "		(select a.CHAR_VALUE_CD CHAR_VALUE_CD11, a.CHAR_VALUE_NM CHAR_VALUE_NM11, b.CLASS_NM classdesc, "
						lgStrSQL = lgStrSQL & "				'' CHAR_VALUE_CD22, '' CHAR_VALUE_NM22 "
						lgStrSQL = lgStrSQL & "		from	B_CHAR_VALUE a,	B_CLASS b "
						lgStrSQL = lgStrSQL & "		where	a.CHAR_CD = b.CHAR_CD1 "

						If pCode1 = "" Then
							lgStrSQL = lgStrSQL & "		and	b.class_cd = " & pCode & " ) aa) aaa "	
						Else
							lgStrSQL = lgStrSQL & "		and	b.class_cd = " & pCode
							lgStrSQL = lgStrSQL & "		and	a.char_value_cd >= " & pCode1 & " ) aa) aaa "	
						End If

'						lgStrSQL = lgStrSQL & " where	aaa.CHAR_VALUE_CD11 not in "
'						lgStrSQL = lgStrSQL & " 		(select	a.CHAR_VALUE_CD1 "
'						lgStrSQL = lgStrSQL & " 		from	b_item a, b_item_by_plant b "
'						lgStrSQL = lgStrSQL & " 		where	a.class_cd = " & pCode
'						lgStrSQL = lgStrSQL & " 		and	isNULL(a.CHAR_VALUE_CD1,'') <> '' "
'						lgStrSQL = lgStrSQL & " 		and	a.item_cd *= b.item_cd "
'						lgStrSQL = lgStrSQL & " 		group by a.CHAR_VALUE_CD1) "
						lgStrSQL = lgStrSQL & " where	convert(varchar(18),aaa.item_cd) not in (select item_cd from b_item where class_cd = " & pCode & " ) "
						lgStrSQL = lgStrSQL & " order by 1, 2 "
					End If

                       
                    'Response.Write lgStrSQL
					'Response.End 

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
                .DBQueryOk        
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
