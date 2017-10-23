<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%

    On Error Resume Next
    Err.Clear
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*","NOCOOKIE", "MB")
    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	Dim lgStrPrevKey
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim txtpayCD , txtpayNM
	Dim txtFactoryCD ,txtFactoryNm


    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
	Const C_SHEETMAXROWS_D  = 100

    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)


    'Multi Multi SpreadSheet
    lgCurrentSpd      = Request("lgCurrentSpd")                                      '☜: "M"(Spread #1) "S"(Spread #2)
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------


	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update            
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Call SubBizQueryMulti()
End Sub	
'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
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
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iKey1

    On Error Resume Next
    Err.Clear
        '---------- Developer Coding part (Start) ---------------------------------------------------------------
    iKey1 = iKey1 & FilterVar(lgKeyStream(0), "''","S")
    iKey1 = iKey1 & " AND B.PAY_TYPE = " & FilterVar(lgKeyStream(1), "''", "S")   
    iKey1 = iKey1 & " AND B.BIZ_AREA_CD = " & FilterVar(lgKeyStream(2), "''", "S")    

   Call CommonQueryRs("minor_NM "," B_minor"," major_cd = " & FilterVar("H0040", "''", "S") & "  AND (minor_cd >=" & FilterVar("2", "''", "S") & "  and minor_cd <= " & FilterVar("9", "''", "S") & " ) and minor_CD = " & FilterVar(lgKeyStream(1), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        
        if Trim(Replace(lgF0,Chr(11),"")) = "X" then
			txtpayCD = ""
            txtpayNM = ""
            Call DisplayMsgBox("800142", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
            Response.End 
        else
			txtpayCD = Trim(lgKeyStream(1))
            txtpayNM = Trim(Replace(lgF0,Chr(11),""))
		End If
		
   Call CommonQueryRs("BIZ_AREA_NM "," B_BIZ_AREA","BIZ_AREA_CD = " & FilterVar(lgKeyStream(2), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        
        if Trim(Replace(lgF0,Chr(11),"")) = "X" then
			txtFactoryCD = ""
            txtFactoryNm = ""
            Call DisplayMsgBox("800142", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
            Response.End 
            'EXIT SUB
        else
			txtFactoryCD = Trim(lgKeyStream(2))
            txtFactoryNm = Trim(Replace(lgF0,Chr(11),""))
		End If

    
    Call SubMakeSQLStatements("MR",iKey1,"X",C_EQGT)   
    
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        If lgCurrentSpd = "M" Then
           Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
            Response.End 
        End If   
        Call SetErrorStatus()
    Else
    
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKeyIndex)

        lgstrData = ""
        
        iDx = 1
        
        Do While Not lgObjRs.EOF
         Select Case lgCurrentSpd
           Case "M","M1"
			
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))

            lgstrData = lgstrData & Chr(11) & ""		 'button  

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TYPE"))
            lgstrData = lgstrData & Chr(11) & ""		 'button
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("AMT"),  ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("A1"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("A2"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("A3"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("A4"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("A5"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("A6"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("A7"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("A8"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("A9"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("A10"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("A11"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("A12"), ggAmtOfMoney.DecPoint, 0)  
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CD"))      
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ORG"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("INTERNAL_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_AREA_CD"))     
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TYPECD"))           
                      
           Case Else 
            
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0)) 

            lgstrData = lgstrData & Chr(11) & ""		 'button  
            lgstrData = lgstrData & Chr(11) & ""

            lgstrData = lgstrData & Chr(11) & ""		 'button
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("B"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("B1"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("B2"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("B3"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("B4"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("B5"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("B6"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("B7"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("B8"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("B9"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("B10"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("B11"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("B12"), ggAmtOfMoney.DecPoint, 0)  
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ""            
            lgstrData = lgstrData & Chr(11) & ""            
                      
      End Select      
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

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    Call SubCloseRs(lgObjRs)         
    %>
<SCRIPT LANGUAGE=vbscript>
	With Parent
		.frm1.txtpayCD.value = "<%=ConvSPChars(txtpayCD)%>"
		.frm1.txtpayNM.value = "<%=ConvSPChars(txtpayNM)%>"		
		.frm1.txtFactoryCD.value = "<%=ConvSPChars(txtFactoryCD)%>"
		.frm1.txtFactoryNM.value = "<%=ConvSPChars(txtFactoryNM)%>"
	END With
</SCRIPT>
<%                                                 '☜: Release RecordSSet

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
' Name : SubBizSaveMultiCreate
' Desc : Save Multi Data
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)

    On Error Resume Next
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

	'데이터 중복 체크 
	lgStrSQL =			  " select  count(PAY_TYPE) as Cnt "
	lgStrSQL = lgStrSQL & " from	A_MONTHLY_BONUS			  "
	lgStrSQL = lgStrSQL & " where	YYYY	=	"	& FilterVar(arrColVal(2), "''", "S")
	lgStrSQL = lgStrSQL & " and		PAY_TYPE =	"	& FilterVar(arrColVal(3), "''", "S")
	lgStrSQL = lgStrSQL & " and		DEPT_CD =	"	& FilterVar(arrColVal(4), "''", "S")
	lgStrSQL = lgStrSQL & " and		ACCT_TYPE =	"	& FilterVar(arrColVal(5), "''", "S")
	lgStrSQL = lgStrSQL & " and		BIZ_AREA_CD =	"	& FilterVar(arrColVal(8), "''", "S")
	lgStrSQL = lgStrSQL & " and		ORG_CHANGE_ID =	"	& FilterVar(arrColVal(6), "''", "S")
	lgStrSQL = lgStrSQL & " and		INTERNAL_CD =	"	& FilterVar(arrColVal(7), "''", "S")
													
			
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then                   'R(Read) X(CursorType) X(LockType) 
		If cdbl(lgObjRs("Cnt")) <> 0  Then
			Call DisplayMsgBox("800446", vbInformation, "", "", I_MKSCRIPT)
			Call SetErrorStatus
			Exit Sub
	   End If
	End if
    lgStrSQL = "INSERT INTO A_MONTHLY_BONUS("
    lgStrSQL = lgStrSQL & " YYYY        ," 
    lgStrSQL = lgStrSQL & " PAY_TYPE    ," 
    lgStrSQL = lgStrSQL & " DEPT_CD     ," 
    lgStrSQL = lgStrSQL & " ACCT_TYPE   ," 
    lgStrSQL = lgStrSQL & " AMT_01       ," 
    lgStrSQL = lgStrSQL & " AMT_02       ," 
    lgStrSQL = lgStrSQL & " AMT_03       ," 
    lgStrSQL = lgStrSQL & " AMT_04       ,"    
    lgStrSQL = lgStrSQL & " AMT_05       ," 
    lgStrSQL = lgStrSQL & " AMT_06       ," 
    lgStrSQL = lgStrSQL & " AMT_07       ," 
    lgStrSQL = lgStrSQL & " AMT_08       ,"    
    lgStrSQL = lgStrSQL & " AMT_09       ," 
    lgStrSQL = lgStrSQL & " AMT_10       ," 
    lgStrSQL = lgStrSQL & " AMT_11       ," 
    lgStrSQL = lgStrSQL & " AMT_12       ,"    
    lgStrSQL = lgStrSQL & " BIZ_AREA_CD  ," 
    lgStrSQL = lgStrSQL & " ORG_CHANGE_ID,"
    lgStrSQL = lgStrSQL & " INTERNAL_CD,"
    lgStrSQL = lgStrSQL & " INSRT_USER_ID," 
    lgStrSQL = lgStrSQL & " INSRT_DT     ," 
    lgStrSQL = lgStrSQL & " UPDT_USER_ID ," 
    lgStrSQL = lgStrSQL & " UPDT_DT      )" 
    lgStrSQL = lgStrSQL & " VALUES("     
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(2), "''", "S")      & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(3), "''", "S")      & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(4), "''", "S")      & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(5), "''", "S")      & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(9),0)          & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(10),0)          & ","
	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(11),0)         & ","
	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(12),0)         & ","
	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(13),0)         & ","
	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(14),0)         & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(15),0)         & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(16),0)         & ","
	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(17),0)         & ","
	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(18),0)         & ","
	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(19),0)         & ","
	lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(20),0)         & ","    
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(8), "''", "S")      & ","
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(6), "''", "S")      & ","    
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(7), "''", "S")      & ","    
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")            & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")  & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")            & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & ")"



    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = " UPDATE  A_MONTHLY_BONUS"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " AMT_01   = " & UNIConvNum(arrColVal(9),0)         & ","
    lgStrSQL = lgStrSQL & " AMT_02   = " & UNIConvNum(arrColVal(10),0)         & ","
    lgStrSQL = lgStrSQL & " AMT_03   = " & UNIConvNum(arrColVal(11),0)         & ","
    lgStrSQL = lgStrSQL & " AMT_04   = " & UNIConvNum(arrColVal(12),0)         & ","
    lgStrSQL = lgStrSQL & " AMT_05   = " & UNIConvNum(arrColVal(13),0)         & ","
    lgStrSQL = lgStrSQL & " AMT_06   = " & UNIConvNum(arrColVal(14),0)         & ","
    lgStrSQL = lgStrSQL & " AMT_07   = " & UNIConvNum(arrColVal(15),0)         & ","
    lgStrSQL = lgStrSQL & " AMT_08   = " & UNIConvNum(arrColVal(16),0)         & ","
    lgStrSQL = lgStrSQL & " AMT_09   = " & UNIConvNum(arrColVal(17),0)         & ","
    lgStrSQL = lgStrSQL & " AMT_10   = " & UNIConvNum(arrColVal(18),0)         & ","
    lgStrSQL = lgStrSQL & " AMT_11   = " & UNIConvNum(arrColVal(19),0)         & ","
    lgStrSQL = lgStrSQL & " AMT_12   = " & UNIConvNum(arrColVal(20),0)        
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " YYYY           = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "	AND PAY_TYPE   = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & "	AND DEPT_CD    = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    lgStrSQL = lgStrSQL & "	AND ACCT_TYPE  = " & FilterVar(UCase(arrColVal(5)), "''", "S")
    lgStrSQL = lgStrSQL & "	AND BIZ_AREA_CD    = " & FilterVar(UCase(arrColVal(8)), "''", "S")
    lgStrSQL = lgStrSQL & "	AND ORG_CHANGE_ID    = " & FilterVar(UCase(arrColVal(6)), "''", "S")
    lgStrSQL = lgStrSQL & "	AND INTERNAL_CD    = " & FilterVar(UCase(arrColVal(7)), "''", "S")
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
   
End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "DELETE  A_MONTHLY_BONUS"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "     YYYY       = " & FilterVar(arrColVal(2),"''"  ,"S")
    lgStrSQL = lgStrSQL & " AND PAY_TYPE   = " & FilterVar(arrColVal(3),"''"  ,"S")
    lgStrSQL = lgStrSQL & " AND DEPT_CD    = " & FilterVar(arrColVal(4),"''"  ,"S")
    lgStrSQL = lgStrSQL & " AND ACCT_TYPE  = " & FilterVar(arrColVal(5),"''"  ,"S")

    
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    On Error Resume Next
    Err.Clear
    
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
               Case "C"
                       lgStrSQL = "Select *   " 
                       lgStrSQL = lgStrSQL & " From  B_MINOR "
                       lgStrSQL = lgStrSQL & " Where MAJOR_CD " & pComp & pCode
                       lgStrSQL = lgStrSQL & " And   MINOR_CD " & pComp & pCode1
               Case "D"
                       lgStrSQL = "Select *   " 
                       lgStrSQL = lgStrSQL & " From  B_MINOR "
                       lgStrSQL = lgStrSQL & " WHERE MAJOR_CD " & pComp & pCode
                       lgStrSQL = lgStrSQL & " AND   MINOR_CD " & pComp & pCode1
               Case "R"
                       If lgCurrentSpd = "M" or  lgCurrentSpd = "M1" Then
                       lgStrSQL = " Select TOP " & iSelCount  & " DEPT_NM, B.DEPT_CD AS CD, B.ORG_CHANGE_ID AS ORG,B.INTERNAL_CD AS INTERNAL_CD, B.BIZ_AREA_CD AS BIZ_AREA_CD, B.ACCT_TYPE AS TYPECD,  A.MINOR_NM AS TYPE,  "
                       lgStrSQL = lgStrSQL     &			   " (ISNULL(AMT_01,0) + ISNULL(AMT_02,0) + ISNULL(AMT_03,0) + ISNULL(AMT_04,0)+ ISNULL(AMT_05,0) + ISNULL(AMT_06,0) +   " 
                       lgStrSQL = lgStrSQL     & " ISNULL(AMT_07,0) +ISNULL(AMT_08,0) + ISNULL(AMT_09,0)+ ISNULL(AMT_10,0) + ISNULL(AMT_11,0) + ISNULL(AMT_12,0) ) AS AMT, "                        
                       lgStrSQL = lgStrSQL     &			   " ISNULL(AMT_01,0) A1,ISNULL(AMT_02,0) A2, ISNULL(AMT_03,0) A3, ISNULL(AMT_04,0) A4, ISNULL(AMT_05,0) A5, ISNULL(AMT_06,0) A6,   " 
                       lgStrSQL = lgStrSQL     & " ISNULL(AMT_07,0) A7,ISNULL(AMT_08,0) A8, ISNULL(AMT_09,0) A9, ISNULL(AMT_10,0) A10, ISNULL(AMT_11,0) A11, ISNULL(AMT_12,0) A12 " 
                       lgStrSQL = lgStrSQL		& " From   B_MINOR A, A_MONTHLY_BONUS B, B_BIZ_AREA C, B_ACCT_DEPT D"
                       lgStrSQL = lgStrSQL		& " Where  A.MINOR_CD = B.ACCT_TYPE AND B.DEPT_CD = D.DEPT_CD AND  B.BIZ_AREA_CD= C.BIZ_AREA_CD AND A.MAJOR_CD = " & FilterVar("H0071", "''", "S") & "  "
                       lgStrSQL = lgStrSQL		& "      and d.org_change_id = (select max(org_change_id) from b_acct_dept) AND B.YYYY = " & pCode                     
                       lgStrSQL = lgStrSQL		& " Order by Dept_nm "
                  
                       Else
                       lgStrSQL = " Select TOP " & iSelCount & " C.BIZ_AREA_NM,sum(ISNULL(AMT_01,0) + ISNULL(AMT_02,0) + ISNULL(AMT_03,0) + ISNULL(AMT_04,0)+ ISNULL(AMT_05,0) + ISNULL(AMT_06,0) +   "                        
                       lgStrSQL = lgStrSQL       & " ISNULL(AMT_07,0) +ISNULL(AMT_08,0) + ISNULL(AMT_09,0)+ ISNULL(AMT_10,0) + ISNULL(AMT_11,0) + ISNULL(AMT_12,0) ) AS B, "                        
                       lgStrSQL = lgStrSQL       & " SUM(ISNULL(AMT_01,0)) B1,SUM(ISNULL(AMT_02,0)) B2, SUM(ISNULL(AMT_03,0)) B3, SUM(ISNULL(AMT_04,0)) B4, SUM(ISNULL(AMT_05,0)) B5, SUM(ISNULL(AMT_06,0)) B6,   "                        
                       lgStrSQL = lgStrSQL       & " SUM(ISNULL(AMT_07,0)) B7,SUM(ISNULL(AMT_08,0)) B8,  SUM(ISNULL(AMT_09,0)) B9, SUM(ISNULL(AMT_10,0)) B10,SUM(ISNULL(AMT_11,0)) B11, SUM(ISNULL(AMT_12,0)) B12 "                        
                       lgStrSQL = lgStrSQL		& " From   A_MONTHLY_BONUS B, B_BIZ_AREA C, B_ACCT_DEPT D"
                       lgStrSQL = lgStrSQL		& " Where  B.DEPT_CD = D.DEPT_CD AND B.BIZ_AREA_CD= C.BIZ_AREA_CD  "
                       lgStrSQL = lgStrSQL		& "   and d.org_change_id = (select max(org_change_id) from b_acct_dept)    AND B.YYYY = " & pCode
					   lgStrSQL = lgStrSQL		& " group by C.BIZ_AREA_NM"                    
   
                       End If             
               Case "U"
                       lgStrSQL = "Select *   " 
                       lgStrSQL = lgStrSQL & " From  B_MINOR "
                       lgStrSQL = lgStrSQL & " Where MAJOR_CD " & pComp & pCode
                       lgStrSQL = lgStrSQL & " And   MINOR_CD " & pComp & pCode1
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
                If Trim("<%=lgCurrentSpd%>") = "M" or Trim("<%=lgCurrentSpd%>") = "M1" Then
                   .ggoSpread.Source     = .frm1.vspdData
                   .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
                Else
                   .ggoSpread.Source     = .frm1.vspdData1
                   .lgStrPrevKeyIndex1    = "<%=lgStrPrevKeyIndex%>"
                End If  
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             'Parent.SubSetErrPos(Trim("<%'=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
       
</Script>	
