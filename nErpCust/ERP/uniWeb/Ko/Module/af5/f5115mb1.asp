<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMAin.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc"  -->

<%
Dim txtBpNm

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call LoadBasisGlobalInf()    
    Call LoadInfTB19029B("I","*","NOCOOKIE","MB")   
    Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB") 

    Call HideStatusWnd  
    Dim sChangeOrgId                                                           '��: Hide Processing message
    Dim adCmdText
    Dim adExcuteNoRecords

        sChangeOrgId = GetGlobalInf("gChangeOrgId")                                                              '��: Hide Processing message
    
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    Const C_SHEETMAXROWS_D = 100
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '��: Save,Update
'             Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
    End Select

    If lgErrorStatus  = "YES" Then
		ObjectContext.SetAbort
	Else
		ObjectContext.SetComplete
	End if				

    Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    'Dim strWhere
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6    
	Dim strWhere
	
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

	' ���Ѱ��� �߰� 
	Dim lgAuthBizAreaCd	' ����� 
	Dim lgInternalCd	' ���κμ� 
	Dim lgSubInternalCd	' ���κμ�(��������)
	Dim lgAuthUsrID		' ���� 

	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))


    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	'�ŷ�ó�� 
	If Trim(lgKeyStream(1)) <> "" Then				
		Call CommonQueryRs(" BP_NM "," B_BIZ_PARTNER "," BP_CD =  " & FilterVar(lgKeyStream(1), "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		If Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  txtBpNm = ""
		    Call DisplayMsgBox("126100", vbInformation, "", "", I_MKSCRIPT)                  '��: No data is found. 
			Call SetErrorStatus()
			exit sub
		Else   
		  txtBpNm = ConvSPChars(Trim(Replace(lgF0,Chr(11),"")))
		End if    	    
	Else 
		txtBpNm = ""
	End If

    strWhere = ""   
    If lgKeyStream(0) <> "" then  strWhere = strWhere & " And f_note.NOTE_NO =  " & FilterVar(lgKeyStream(0), "''", "S") & "  "       
    If lgKeyStream(1) <> "" then  strWhere = strWhere & " And f_note.BP_CD =  " & FilterVar(lgKeyStream(1), "''", "S") & "  "           
    If lgKeyStream(2) <> "" then  strWhere = strWhere & " And convert(char(10), f_note.ISSUE_DT, 126) >=  " & FilterVar(UniConvDate(lgKeyStream(2)), "''", "S") & " " 
    If lgKeyStream(3) <> "" then  strWhere = strWhere & " And convert(char(10), f_note.ISSUE_DT, 126) <=  " & FilterVar(UniConvDate(lgKeyStream(3)), "''", "S") & " " 
    If lgKeyStream(4) <> "" then  strWhere = strWhere & " And convert(char(10), f_note.DUE_DT, 126) <=  " & FilterVar(UniConvDate(lgKeyStream(4)), "''", "S") & " " 
	If lgKeyStream(5) <> "" then  strWhere = strWhere & " And f_note.NOTE_FG  =  " & FilterVar(lgKeyStream(5), "''", "S") & " " 
 
 	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		strWhere		= strWhere & " AND f_note.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		strWhere		= strWhere & " AND f_note.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		strWhere		= strWhere & " AND f_note.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		strWhere		= strWhere & " AND f_note.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If
 
    Call SubMakeSQLStatements("MR", strWhere, "X", C_LIKE)                                 '�� : Make sql statements
 
    If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKeyIndex)

        lgstrData = ""        
        iDx       = 1        
        
        Do While Not lgObjRs.EOF
                      
            lgstrData = lgstrData & Chr(11) & ""								'�������� 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0))						'������ȣ 
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ""								'�������� 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgObjRs(3)))					'�μ��ڵ�			
            lgstrData = lgstrData & Chr(11) & ""								'�μ��ڵ�popup
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(4))						'�μ��� 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgObjRs(5)))					'�ŷ�ó�ڵ� 
            lgstrData = lgstrData & Chr(11) & ""								'�ŷ�ó�ڵ�popup                       
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(6))						'�ŷ�ó�� 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(7))						'�����ڵ� 
            lgstrData = lgstrData & Chr(11) & ""								'�����ڵ�popup
            lgstrData = lgstrData & Chr(11) & ConvSPChars(Trim(lgObjRs(8)))					'�����ڵ��            
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(9))	'������                      
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(10))  '������           
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(11),	ggAmtOfMoney.DecPoint		,0)						'�����ݾ�                        
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(20),	ggAmtOfMoney.DecPoint		,0)			'������ 
            lgstrData = lgstrData & Chr(11) & ""								'�������CD
            lgstrData = lgstrData & Chr(11) & ""								'��Ÿ������CD
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(14))						'������ 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(15))						'��� 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1))						'��������CD
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(2))						'��������CD
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(12))						'�������CD
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(13))						'��Ÿ������CD            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(16))						'COSTCD
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(17))						'BIZAREACD
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(18))						'INTERNALCD
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(19))   		            'ORG_CHANGE_ID
            
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
	
	Call SubCloseRs(lgObjRs)                                                          '��: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '��: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '��: Split Column data

		If ChkAuth(arrColVal) = True Then
			Select Case arrColVal(0)
			    Case "C"
					Call SubBizSaveMultiCreate(arrColVal)                            '��: Create
			    Case "U"
			        Call SubBizSaveMultiUpdate(arrColVal)                            '��: Update
			    Case "D"
			        Call SubBizSaveMultiDelete(arrColVal)                            '��: Delete
			End Select
		End If
		        
        If lgErrorStatus  = "YES" Then
            lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
			Exit For
        End If
    Next
End Sub    

Function ChkAuth(arrColVal)

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

	Const I1_a_pgm_value_note_no = 2
    Const I1_a_pgm_value_dept_cd = 5
    Const I1_a_pgm_value_org_change_id = 18
    Const I1_a_pgm_value_internal_cd = 17
    Const I1_a_pgm_value_biz_area_cd = 16
    Const I1_a_pgm_value_updt_user_id = 0

    ' ���Ѱ����� ���� define
    Dim iStrSQL
    Dim objAChkDataAuth 

    Dim L1_a_data_auth_cud_char 
    
    Dim L2_a_pgm_value
    Const L2_a_pgm_value_dept_cd = 0
    Const L2_a_pgm_value_internal_cd = 1
    Const L2_a_pgm_value_biz_area_cd = 2
    Const L2_a_pgm_value_updt_user_id = 3

	' -- ���Ѱ����߰� 
	Dim I1_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 

  	Redim I1_a_data_auth(3)
	I1_a_data_auth(0)	= Trim(Request("txthAuthBizAreaCd"))
	I1_a_data_auth(1)	= Trim(Request("txthhInternalCd"))
	I1_a_data_auth(2)	= Trim(Request("txthSubInternalCd"))
	I1_a_data_auth(3)   = Trim(Request("txthAuthUsrID"))

	Redim L2_a_pgm_value(3)

	ChkAuth = False
	
    ' -- ���Ѱ��� �߰� 2006-08-01 JYK start .. CHOE0TAE Ŀ���͸���¡�Ѱ� 
    ' -- ���� DLL �� �ѱ�� ��ɾ� 
    L1_a_data_auth_cud_char = cstr(UCase(Trim(arrColVal(0))))
    
    ' -- ��ȸ SQL
    iStrSQL = ""
    iStrSQL = iStrSQL & "DECLARE @DEPT_CD    CHAR(10) " & vbCrLf & _
                        ",   @INTERNAL_CD    CHAR(30) " & vbCrLf & _
                        ",   @BIZ_AREA_CD    VARCHAR(10) " & vbCrLf & _
                        ",   @UPDT_USER_ID   VARCHAR(13) " & vbCrLf
        
    If UCase(Trim(L1_a_data_auth_cud_char)) = "C" Or UCase(Trim(L1_a_data_auth_cud_char)) = "U" Then
        'asp���� �Է¹��� �μ����� 
        'asp���� �Է¹��� ���κμ��ڵ�/�Է¹��� ����������̵�(org_change_id)�� �μ��ڵ�(dept_cd)�� ��ġ�� ���κμ��ڵ�(internal_cd)
        'asp���� �Է¹��� ������ڵ� / �Է¹��� ����������̵�� �μ��ڵ�� cost_cd�� ã�� b_cost_center���� cost_cd�� ��ġ�� ������ڵ�(biz_area_cd)

        iStrSQL = iStrSQL & "SELECT  @INTERNAL_CD = A.INTERNAL_CD " & vbCrLf & _
                            ",   @BIZ_AREA_CD = B.BIZ_AREA_CD " & vbCrLf & _
                            ",   @DEPT_CD = " & FilterVar(arrColVal(I1_a_pgm_value_dept_cd), "''", "S") & vbCrLf & _
                            "FROM    B_ACCT_DEPT  A " & vbCrLf & _
                            "    INNER JOIN B_COST_CENTER B ON A.COST_CD = B.COST_CD " & vbCrLf & _
                            "WHERE   A.ORG_CHANGE_ID = " & FilterVar(arrColVal(I1_a_pgm_value_org_change_id), "''", "S") & vbCrLf & _
                            "AND A.DEPT_CD = " & FilterVar(arrColVal(I1_a_pgm_value_dept_cd), "''", "S") & vbCrLf
    End If
        
    If UCase(Trim(L1_a_data_auth_cud_char)) = "U" Then
         '������� �������� ������ID
             
        iStrSQL = iStrSQL & "SELECT  @UPDT_USER_ID = UPDT_USER_ID " & vbCrLf & _
                            "FROM    F_NOTE " & vbCrLf & _
                            "WHERE   NOTE_NO = " & FilterVar(arrColVal(I1_a_pgm_value_note_no), "''", "S") & vbCrLf

    ElseIf UCase(Trim(L1_a_data_auth_cud_char)) = "D" Then
        '������� �������� �μ��ڵ� 
        '������� �������� ���κμ��ڵ� 
        '������� �������� ������ڵ� 
        '������� �������� ������ID
        iStrSQL = iStrSQL & "SELECT  @DEPT_CD = DEPT_CD " & vbCrLf & _
                            ",   @INTERNAL_CD = INTERNAL_CD " & vbCrLf & _
                            ",   @BIZ_AREA_CD = BIZ_AREA_CD " & vbCrLf & _
                            ",   @UPDT_USER_ID = UPDT_USER_ID " & vbCrLf & _
                            "FROM    F_NOTE " & vbCrLf & _
                            "WHERE   NOTE_NO = " & FilterVar(arrColVal(I1_a_pgm_value_note_no), "''", "S") & vbCrLf
    End If
        
    ' -- ����Ÿ ���� 
    iStrSQL = iStrSQL & "SELECT @DEPT_CD DEPT_CD, @INTERNAL_CD INTERNAL_CD, @BIZ_AREA_CD BIZ_AREA_CD, @UPDT_USER_ID UPDT_USER_ID"
        
    If 	FncOpenRs("R", lgObjConn, lgObjRs, iStrSQL, "X", "X") = True Then
        L2_a_pgm_value(L2_a_pgm_value_dept_cd) = lgObjRs(0)
        L2_a_pgm_value(L2_a_pgm_value_internal_cd) = lgObjRs(1)
        L2_a_pgm_value(L2_a_pgm_value_biz_area_cd) = lgObjRs(2)
        L2_a_pgm_value(L2_a_pgm_value_updt_user_id) = lgObjRs(3)
    End If

    ' -- ���Ѱ��� ȣ�� 
    Set objAChkDataAuth = Server.CreateObject("PA0CG07.cAChkDataAuthSvr")

    If CheckSYSTEMError(Err,True) = True Then
		Exit Function	
    End If

    Call objAChkDataAuth.A_CHECK_DATA_AUTH_SVR(gStrGlobalCollection, L1_a_data_auth_cud_char, I1_a_data_auth, L2_a_pgm_value)

    If CheckSYSTEMError(Err,True) = True Then
		lgErrorStatus = "YES"
		ObjectContext.SetAbort
		Exit Function
    End If

    Set objAChkDataAuth = Nothing

    If lgErrorStatus = "YES" Then Exit Function
    ' -- ���Ѱ��� �߰� 2006-08-01 JYK end
    
    ChkAuth = True
End Function

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    '--------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
                        
    lgStrSQL = "INSERT INTO F_NOTE("
    lgStrSQL = lgStrSQL & " NOTE_NO,		NOTE_FG,		NOTE_AMT,		STTL_AMT, "
    lgStrSQL = lgStrSQL & " ISSUE_DT,		DUE_DT,			PLACE,			RCPT_FG, " 
    lgStrSQL = lgStrSQL & " PUBLISHER,		NOTE_STS,		NOTE_DESC,		BP_CD, "
	lgStrSQL = lgStrSQL & " BANK_CD,		BIZ_AREA_CD,	ORG_CHANGE_ID,	DEPT_CD, "     
    lgStrSQL = lgStrSQL & " INTERNAL_CD,	COST_CD,		ENDORSE_FG,		BP_ENDORSE_CD, "
    lgStrSQL = lgStrSQL & " BP_ORG_CD,		USED_FG,		INSRT_USER_ID,	INSRT_DT, "
    lgStrSQL = lgStrSQL & " UPDT_USER_ID,	UPDT_DT,		CASH_RATE,		CASH_AMT ) " 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & " " & FilterVar(arrColVal(2), "''", "S") & " , "			'������ȣ 
    lgStrSQL = lgStrSQL & " " & FilterVar(arrColVal(3), "''", "S") & " , "			'�������� 
    lgStrSQL = lgStrSQL & "" & arrColVal(10) & ", "								'�����ݾ� 
    lgStrSQL = lgStrSQL & "0, "														'
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDate(arrColVal(8))	,null,"S") & ", "	'������   
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDate(arrColVal(9))	,null,"S") & ", "   '������ 
    lgStrSQL = lgStrSQL & " " & FilterVar(arrColVal(11), "''", "S") & ", "								'�������   
    lgStrSQL = lgStrSQL & " " & FilterVar(arrColVal(12), "''", "S") & ", "								'�ڼ�Ÿ������ 
    lgStrSQL = lgStrSQL & " " & FilterVar(arrColVal(13)	, "''", "S") & " , "			'������ 
    lgStrSQL = lgStrSQL & " " & FilterVar(arrColVal(4), "''", "S") & ", "								'�������� 
    lgStrSQL = lgStrSQL & " " & FilterVar(arrColVal(14)	, "''", "S") & " , "			'��� 
    lgStrSQL = lgStrSQL & " " & FilterVar(arrColVal(6), "''", "S") & " , "			'�ŷ�ó 
    lgStrSQL = lgStrSQL & " " & FilterVar(arrColVal(7), "''", "S") & " , "			'�����ڵ� 
    lgStrSQL = lgStrSQL & " " & FilterVar(arrColVal(16), "''", "S") & ", "								'������ڵ� 
    lgStrSQL = lgStrSQL & " " & FilterVar(arrColVal(18), "''", "S") & ", "								'����������̵� 
    lgStrSQL = lgStrSQL & " " & FilterVar(arrColVal(5), "''", "S") & " , "			'�μ��ڵ� 
    lgStrSQL = lgStrSQL & " " & FilterVar(arrColVal(17), "''", "S") & ", "								'�����ڵ� 
    lgStrSQL = lgStrSQL & " " & FilterVar(arrColVal(15), "''", "S") & ", "								'�ڽ�Ʈ���� 
    lgStrSQL = lgStrSQL & " " & FilterVar(arrColVal(3), "''", "S") & ", "								'
    lgStrSQL = lgStrSQL & "'', "													'
    lgStrSQL = lgStrSQL & "'', "													'
    lgStrSQL = lgStrSQL & "" & FilterVar("Y", "''", "S") & " , "    
    lgStrSQL = lgStrSQL & FilterVar(gUsrId				, "''", "S") & ", "		
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime		,NULL,"S") & ", "
    lgStrSQL = lgStrSQL & FilterVar(gUsrId				, "''", "S") & ", "		
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime		,NULL,"S") & ", "
    lgStrSQL = lgStrSQL & "  " & FilterVar(arrColVal(19), "''", "S") & ", "    
    lgStrSQL = lgStrSQL & "0 )"
    
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
    Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
    lgStrSQL = " UPDATE F_NOTE_NO SET STS = " & FilterVar("PR", "''", "S") & "  WHERE NOTE_NO = " & FilterVar(arrColVal(2), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
    Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
'    Response.Write                
    lgStrSQL = "UPDATE  F_NOTE"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " NOTE_FG		 =  " & FilterVar(UCase(arrColVal(3))   , "''", "S") & ", "
    lgStrSQL = lgStrSQL & " NOTE_AMT     =  " & FilterVar(UCase(arrColVal(10)), "''", "S") & ", "       
    lgStrSQL = lgStrSQL & " ISSUE_DT     =  " & FilterVar(UNIConvDate(arrColVal(8))	,null,"S") & ", " 
    lgStrSQL = lgStrSQL & " DUE_DT       =  " & FilterVar(UNIConvDate(arrColVal(9))	,null,"S") & ", " 
    lgStrSQL = lgStrSQL & " PLACE        =  " & FilterVar(UCase(arrColVal(11)), "''", "S") & ", "   
    lgStrSQL = lgStrSQL & " RCPT_FG      =  " & FilterVar(UCase(arrColVal(12)), "''", "S") & ", "  
    lgStrSQL = lgStrSQL & " PUBLISHER    =  " & FilterVar(UCase(arrColVal(13)), "''", "S") & " , " 
    lgStrSQL = lgStrSQL & " NOTE_DESC    =  " & FilterVar(UCase(arrColVal(14)), "''", "S") & " , "   
    lgStrSQL = lgStrSQL & " BP_CD        =  " & FilterVar(UCase(arrColVal(6)), "''", "S") & " , "  
    lgStrSQL = lgStrSQL & " BANK_CD      =  " & FilterVar(UCase(arrColVal(7)), "''", "S") & " , "   
    lgStrSQL = lgStrSQL & " BIZ_AREA_CD  =  " & FilterVar(UCase(arrColVal(16)), "''", "S") & ", "  
    lgStrSQL = lgStrSQL & " ORG_CHANGE_ID=  " & FilterVar(UCase(arrColVal(18)), "''", "S") & ", "      
    lgStrSQL = lgStrSQL & " CASH_RATE	 =  " & FilterVar(arrColVal(19), "''", "S") & ", "    
    lgStrSQL = lgStrSQL & " DEPT_CD      =  " & FilterVar(UCase(arrColVal(5)), "''", "S") & " , "   
    lgStrSQL = lgStrSQL & " INTERNAL_CD  =  " & FilterVar(UCase(arrColVal(17)), "''", "S") & ", "  
    lgStrSQL = lgStrSQL & " COST_CD      =  " & FilterVar(UCase(arrColVal(15)), "''", "S") & ", "            
    lgStrSQL = lgStrSQL & " UPDT_USER_ID =  " & FilterVar(gUsrId				, "''", "S") & ", "		
    lgStrSQL = lgStrSQL & " UPDT_DT      =  " & FilterVar(GetSvrDateTime		,NULL,"S") & "  "  
    lgStrSQL = lgStrSQL & " WHERE	NOTE_NO =  " & FilterVar(UCase(arrColVal(2)), "''", "S") & "  "
    'lgStrSQL = lgStrSQL & " AND		NOTE_FG = '" & UCase(arrColVal(3))   & "' "

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    lgStrSQL = "DELETE  F_NOTE "    
    lgStrSQL = lgStrSQL & " WHERE NOTE_NO    =  " & FilterVar(UCase(arrColVal(2)), "''", "S") & "  "    
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

	lgStrSQL = " UPDATE F_NOTE_NO SET STS = " & FilterVar("NP", "''", "S") & "  WHERE NOTE_NO = " & FilterVar(UCase(arrColVal(2)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
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
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKeyIndex + 1
           
           Select Case Mid(pDataType,2,1)
  
               Case "R"
               
                       lgStrSQL = "Select distinct Top " & iSelCount 
                       lgStrSQL = lgStrSQL & " note_no,		note_fg,		note_sts,	f_note.dept_cd, "
                       lgStrSQL = lgStrSQL & " dept_nm,		f_note.bp_cd,	bp_nm,		f_note.bank_cd, "
                       lgStrSQL = lgStrSQL & " bank_nm,		issue_dt,		due_dt,		note_amt, "
                       lgStrSQL = lgStrSQL & " place,		rcpt_fg,		publisher,	note_desc, "
                       lgStrSQL = lgStrSQL & " b_acct_dept.cost_cd, b_cost_center.biz_area_cd, b_acct_dept.internal_cd, f_note.org_change_id, "
                       lgStrSQL = lgStrSQL & " f_note.cash_rate "
                       lgStrSQL = lgStrSQL & " from f_note, b_acct_dept, b_bank, b_biz_partner, b_cost_center "
                       lgStrSQL = lgStrSQL & " where b_acct_dept.dept_cd = f_note.dept_cd "
                       lgStrSQL = lgStrSQL & " and b_acct_dept.org_change_id = f_note.org_change_id "                  
                       lgStrSQL = lgStrSQL & " and b_bank.bank_cd = f_note.bank_cd "
                       lgStrSQL = lgStrSQL & " and b_biz_partner.bp_cd = f_note.bp_cd " 
                       lgStrSQL = lgStrSQL & " and b_acct_dept.cost_cd = b_cost_center.cost_cd " 
                       lgStrSQL = lgStrSQL & " and f_note.note_fg <>" & FilterVar("CR", "''", "S") & "  " 
                       lgStrSQL = lgStrSQL & " and f_note.note_fg <>" & FilterVar("CP", "''", "S") & "  " & pCode 
                       lgStrSQL = lgStrSQL & " order by issue_dt, note_no "                       
                       'Response.write lgStrSQL
                       
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
    lgErrorStatus     = "YES"                                                         '��: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Select Case pOpCode
        Case "MC"
'				Response.Write pErr
                 If CheckSYSTEMError(pErr,True) = True Then

                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                     '  Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
                 
        Case "MD"
				If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,true) = True Then
                       'Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub


%>

<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '�� : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"
                .frm1.txtBpNm.Value	  = "<%=txtBpNm%>"               
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '�� : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '�� : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
    
       
</Script>	
