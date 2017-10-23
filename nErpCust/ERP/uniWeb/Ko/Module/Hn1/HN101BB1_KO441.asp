<%@ LANGUAGE=VBSCript TRANSACTION=Required%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100

    Dim lgSvrDateTime
    Dim LngRecs
    Dim StrDt, StrYYMM, StrProvCD, StrFileGubun

    call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgSvrDateTime = GetSvrDateTime
	LngRecs = 0
    Call HideStatusWnd                                                               '☜: Hide Processing message
'Call ServerMesgBox("bbbbbbb" , vbInformation, I_MKSCRIPT)
	
    
	lgErrorStatus    = "NO"
    lgErrorPos       = ""                                                           '☜: Set to space
    lgOpModeCRUD     = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream      = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow      = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey	 = UNICInt(Trim(Request("lgStrPrevKey")),0)						'☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    StrDt			= Request("htxtDt")  
	StrYYMM			= Request("htxtYYMM")
	StrProvCD		= Request("htxtProvCD")
	StrFileGubun	= Request("htxtFileGubun")

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
 
    Select Case lgOpModeCRUD
        'Case CStr(UID_M0001)                                                         '☜: Query
        '     Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select
'Call ServerMesgBox("End" , vbInformation, I_MKSCRIPT)    
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
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	Call SubBizSaveMultiDelete()
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    
	Dim iDx
    Dim iLoopMax
    Dim iKey1,iKey2,iKey3
   
    On Error Resume Next                                                             '☜: Protect system from crashing
    
	Err.Clear                                                                        '☜: Clear Error status
	
	iKey1 = FilterVar(lgKeyStream(0),"''", "S")										'급/상여일자
	iKey2 = FilterVar(Replace(lgKeyStream(1),"-",""),"''", "S")						'급/상여년월

	If lgKeyStream(2)<>"" then														'지급구분
		iKey3 = FilterVar(lgKeyStream(2) & "%","''", "S")									
	Else
		iKey3 = FilterVar("%","''", "S")
	End If

	Select Case StrFileGubun
		   Case "A"
				lgStrSQL = "SELECT  DEPT_CD, EMP_NO, PAY_YYMM,  "
				lgStrSQL = lgStrSQL & "	dbo.ufn_GetCodeName(" & FilterVar("H0040", "''", "S") & ", PROV_TYPE) PROV_TYPE_NM, " 
				lgStrSQL = lgStrSQL & " PROV_TYPE, PROV_DT, PAY_TOT_AMT, BONUS_TOT_AMT, NONTAX_TOT_AMT, "
				lgStrSQL = lgStrSQL & " CASE WHEN PROV_TYPE = 1 THEN TAX_AMT ELSE BONUS_TAX END PAY_TAX, PROV_TOT_AMT, "
				lgStrSQL = lgStrSQL & "  SUB_TOT_AMT, REAL_PROV_AMT, INCOME_TAX, RES_TAX, ANUT, MED_INSURE, EMP_INSURE "
				lgStrSQL = lgStrSQL & " FROM H_IF_HDF070T "
				lgStrSQL = lgStrSQL & " WHERE PROV_DT =  " & iKey1 & " "  
				lgStrSQL = lgStrSQL & "	AND PAY_YYMM =   " & iKey2 & " " 
				lgStrSQL = lgStrSQL & "	AND PROV_TYPE LIKE  " & iKey3 & " "
				lgStrSQL = lgStrSQL & " ORDER BY PROV_TYPE, PAY_YYMM, EMP_NO "
				
			Case "B"
				lgStrSQL = "SELECT  PAY_YYMM, EMP_NO,   "
				lgStrSQL = lgStrSQL & "	dbo.ufn_GetCodeName(" & FilterVar("H0040", "''", "S") & ", PROV_TYPE) PROV_TYPE_NM, PROV_TYPE, ALLOW_CD, " 
				lgStrSQL = lgStrSQL & "	dbo.ufn_H_GetCodeName(" & FilterVar("HDA010T", "''", "S") & ",ALLOW_CD,'')ALLOW_CD_NM, ALLOW " 				
				lgStrSQL = lgStrSQL & " FROM H_IF_HDF040T "
				lgStrSQL = lgStrSQL & " WHERE PAY_YYMM =   " & iKey2 & " " 				
				lgStrSQL = lgStrSQL & "	AND PROV_TYPE LIKE  " & iKey3 & " "
				lgStrSQL = lgStrSQL & " ORDER BY PROV_TYPE, PAY_YYMM, EMP_NO "
		
		   Case "C"
				lgStrSQL = "SELECT  SUB_YYMM, EMP_NO,   "
				lgStrSQL = lgStrSQL & "	dbo.ufn_GetCodeName(" & FilterVar("H0040", "''", "S") & ", SUB_TYPE) SUB_TYPE_NM, SUB_TYPE, SUB_CD, " 
				lgStrSQL = lgStrSQL & "	dbo.ufn_GetCodeName(" & FilterVar("H0040", "''", "S") & ", SUB_CD) SUB_CD_NM, SUB_AMT " 				
				lgStrSQL = lgStrSQL & " FROM H_IF_HDF060T "
				lgStrSQL = lgStrSQL & " WHERE SUB_YYMM =   " & iKey2 & " " 				
				lgStrSQL = lgStrSQL & "	AND SUB_TYPE LIKE  " & iKey3 & " "
				lgStrSQL = lgStrSQL & " ORDER BY SUB_TYPE, SUB_YYMM, EMP_NO "
	End Select
	
	'response.write lgStrSQL
    'Call SubMakeSQLStatements("MR",iKey1,iKey2,iKey3)                                 '☆ : Make sql statements
   
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then			
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)					 '☜ : No data is found. 
        Call SetErrorStatus()
		
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        iDx       = 1
       
		Select Case StrFileGubun
			   Case "A"
					Do While Not lgObjRs.EOF
						
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_CD"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO"))			
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_YYMM"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROV_TYPE_NM"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROV_TYPE"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROV_DT"))
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PAY_TOT_AMT"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("BONUS_TOT_AMT"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("NONTAX_TOT_AMT"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PAY_TAX"), ggAmtOfMoney.DecPoint,0)						
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PROV_TOT_AMT"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("SUB_TOT_AMT"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("REAL_PROV_AMT"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("INCOME_TAX"), ggAmtOfMoney.DecPoint,0)			
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("RES_TAX"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ANUT"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("MED_INSURE"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("EMP_INSURE"), ggAmtOfMoney.DecPoint,0)
					
						lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
						lgstrData = lgstrData & Chr(11) & Chr(12)

						lgObjRs.MoveNext

						iDx =  iDx + 1
						'If iDx > C_SHEETMAXROWS_D Then
						'   lgStrPrevKey = lgStrPrevKey + 1
						'   Exit Do
						'End If   
						   
					Loop 

			   Case "B"
					
					Do While Not lgObjRs.EOF
						
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_YYMM"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO"))			
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROV_TYPE_NM"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROV_TYPE"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_CD"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ALLOW_CD_NM"))						
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ALLOW"), ggAmtOfMoney.DecPoint,0)						
						lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
						lgstrData = lgstrData & Chr(11) & Chr(12)

						lgObjRs.MoveNext

						iDx =  iDx + 1
						'If iDx > C_SHEETMAXROWS_D Then
						'   lgStrPrevKey = lgStrPrevKey + 1
						'   Exit Do
						'End If   						   
					Loop 

			   Case "C"
					Do While Not lgObjRs.EOF
						
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUB_YYMM"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO"))			
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUB_TYPE_NM"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUB_TYPE"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUB_CD"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUB_CD_NM"))						
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("SUB_AMT"), ggAmtOfMoney.DecPoint,0)						
						lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
						lgstrData = lgstrData & Chr(11) & Chr(12)

						lgObjRs.MoveNext

						iDx =  iDx + 1
						'If iDx > C_SHEETMAXROWS_D Then
						'   lgStrPrevKey = lgStrPrevKey + 1
						'   Exit Do
						'End If   						   
					Loop 
		End Select
        
    End If
    
    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey = ""
    End If   

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
								                                                       '☜: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx
	Dim iKey1

    On Error Resume Next                                                             '☜: Protect system from crashing

    Err.Clear                                                                        '☜: Clear Error status    
	
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data

	If strProvCd <>"" Then
		iKey1 = FilterVar(strProvCd & "%","''", "S")
	Else
		iKey1 = FilterVar("%","''", "S")
	End if
'Call ServerMesgBox("SubBizSaveMulti(StrFileGubun) = " & StrFileGubun , vbInformation, I_MKSCRIPT)  
'	Select Case StrFileGubun
'		   Case "A"				
'				lgStrSQL = " DELETE HDF070T WHERE PROV_DT = " & FilterVar(StrDT,"''", "S") & " "
'				lgStrSQL = lgStrSQL & " AND PAY_YYMM = " & FilterVar(Replace(StrYYMM,"-",""),"''", "S") & "  "
'				lgStrSQL = lgStrSQL & " AND PROV_TYPE LIKE " & iKey1 & " "
'		   Case "B"
'				lgStrSQL = " DELETE HDF040T WHERE PAY_YYMM = " & FilterVar(Replace(StrYYMM,"-",""),"''", "S") & " "				
'				lgStrSQL = lgStrSQL & " AND PROV_TYPE LIKE " & iKey1 & " "
'		   Case "C"
'				lgStrSQL = " DELETE HDF060T WHERE SUB_YYMM = " & FilterVar(Replace(StrYYMM,"-",""),"''", "S") & " "				
'				lgStrSQL = lgStrSQL & " AND SUB_TYPE LIKE " & iKey1 & " "
'	End Select
'
'    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
'	
'    Call SubHandleError("SD",lgObjConn,lgObjRs,Err)
    
'Call ServerMesgBox("SubBizSaveMulti(StrDT) = " & StrDT , vbInformation, I_MKSCRIPT)    
    '*************************************************************
    'Delete from Interface Table 
    '*************************************************************
    Select Case StrFileGubun
		   Case "A"
				lgStrSQL = "DELETE  H_IF_HDF070T"
				lgStrSQL = lgStrSQL & " WHERE PROV_DT = " & FilterVar(StrDT,"''", "S") & " "
				lgStrSQL = lgStrSQL & " AND PAY_YYMM = " & FilterVar(Replace(StrYYMM,"-",""),"''", "S") & "  "
				lgStrSQL = lgStrSQL & " AND PROV_TYPE LIKE " & iKey1 & " "
		   Case "B"
				lgStrSQL = "DELETE  H_IF_HDF040T "
				lgStrSQL = lgStrSQL & " WHERE PAY_YYMM = " & FilterVar(Replace(StrYYMM,"-",""),"''", "S") & " "				
				lgStrSQL = lgStrSQL & " AND PROV_TYPE LIKE " & iKey1 & " "
		   Case "C"
				lgStrSQL = "DELETE  H_IF_HDF060T "
				lgStrSQL = lgStrSQL & " WHERE SUB_YYMM = " & FilterVar(Replace(StrYYMM,"-",""),"''", "S") & " "					
				lgStrSQL = lgStrSQL & " AND SUB_TYPE LIKE " & iKey1 & " "
	End Select

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
    
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
   
    
'Call ServerMesgBox("Ubound(arrRowVal,1) = " & cstr(Ubound(arrRowVal,1)) , vbInformation, I_MKSCRIPT)  
'Call ServerMesgBox("lgLngMaxRow = " & cstr(lgLngMaxRow) , vbInformation, I_MKSCRIPT)    	   
    For iDx = 1 To lgLngMaxRow
'Call ServerMesgBox("arrRowVal(iDx-1) = " & arrRowVal(iDx-1) , vbInformation, I_MKSCRIPT)     
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
'Call ServerMesgBox("arrColVal(0) = " & arrColVal(0) , vbInformation, I_MKSCRIPT)        
        Select Case arrColVal(0)
            Case "C"
	                Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
'            Case "U"
'                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
            Case "D"				
                    Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next
    
	'Call SuSaveCreate()
	
End Sub    


'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SuSaveCreate()

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
   
		
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
'Call ServerMesgBox("StrFileGubun = " & StrFileGubun , vbInformation, I_MKSCRIPT)  
	'지급구분이 급여이면 급여과세총액/상여이면 상여과세총액
	Select Case StrFileGubun
		   Case "A"
				Dim iValue1, iValue2 
				If Trim(UCase(arrColVal(5))) = 1 Then
					iValue1 = arrColVal(10)
					iValue2 = 0
				Else
					iValue1 = 0
					iValue2 = arrColVal(10)
				End if

				lgStrSQL = "INSERT INTO H_IF_HDF070T(DEPT_CD, EMP_NO, PAY_YYMM, PROV_TYPE, PROV_DT, PAY_TOT_AMT, BONUS_TOT_AMT, NONTAX_TOT_AMT, TAX_AMT, "
				lgStrSQL = lgStrSQL & "				 BONUS_TAX, PROV_TOT_AMT, SUB_TOT_AMT, REAL_PROV_AMT, INCOME_TAX, RES_TAX, ANUT, MED_INSURE, EMP_INSURE )" 
				lgStrSQL = lgStrSQL & " VALUES(" 
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(2))),"","S")		& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(3))),"","S")		& ","
				lgStrSQL = lgStrSQL & FilterVar(Replace(StrYYMM,"-",""),"''", "S")		& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(5))),"","S")		& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(6))),"","S")		& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(7)),0)					& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(8)),0)					& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(9)),0)					& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(iValue1),0)						& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(iValue2),0)						& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(11)),0)					& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(12)),0)					& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(13)),0)					& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(14)),0)					& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(15)),0)					& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(16)),0)					& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(17)),0)					& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(18)),0)					
				lgStrSQL = lgStrSQL & ")"
				
				'response.write lgStrSQL
				'Call ServerMesgBox(lgStrSQL , vbInformation, I_MKSCRIPT)
		   Case "B"
				
				lgStrSQL = "INSERT INTO H_IF_HDF040T(PAY_YYMM, EMP_NO, PROV_TYPE, ALLOW_CD, ALLOW )"
				lgStrSQL = lgStrSQL & " VALUES(" 
				lgStrSQL = lgStrSQL & FilterVar(Replace(StrYYMM,"-",""),"''", "S")		& ","
				'lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(2)),"''", "S")			& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(3))),"","S")		& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(4))),"","S")		& ","				
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(5))),"","S")		& ","				
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(6)),0)				
				lgStrSQL = lgStrSQL & ")"
				'response.write lgStrSQL

		   Case "C"
				lgStrSQL = "INSERT INTO H_IF_HDF060T(SUB_YYMM, EMP_NO, SUB_TYPE, SUB_CD, SUB_AMT )"
				lgStrSQL = lgStrSQL & " VALUES(" 
				lgStrSQL = lgStrSQL & FilterVar(Replace(StrYYMM,"-",""),"''", "S")		& ","
				'lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(2)),"''", "S")			& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(3))),"","S")		& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(4))),"","S")		& ","				
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(5))),"","S")		& ","				
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(6)),0)				
				lgStrSQL = lgStrSQL & ")"

	End Select
	
'Call ServerMesgBox(lgStrSQL , vbInformation, I_MKSCRIPT)
'response.write lgStrSQL
'Response.end
	
    'lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    lgObjConn.Execute lgStrSQL, LngRecs, adCmdText
'Call ServerMesgBox(LngRecs , vbInformation, I_MKSCRIPT)      
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete()	
	
    Dim iKey1,iKey2,iKey3
   
    On Error Resume Next																'☜: Protect system from crashing
    
	Err.Clear																			'☜: Clear Error status
	
	iKey1 = FilterVar(lgKeyStream(0),"''", "S")											'급/상여일자
	iKey2 = FilterVar(Replace(lgKeyStream(1),"-",""),"''", "S")							'급/상여년월

	If lgKeyStream(2)<>"" then															'지급구분
		iKey3 = FilterVar(lgKeyStream(2) & "%","''", "S")									
	Else
		iKey3 = FilterVar("%","''", "S")
	End If


	Select Case StrFileGubun
		   Case "A"
				lgStrSQL = "DELETE  H_IF_HDF070T"
				lgStrSQL = lgStrSQL & " WHERE PROV_DT = " & iKey1 & " "
				lgStrSQL = lgStrSQL & " AND PAY_YYMM = " & iKey2 & " "
				lgStrSQL = lgStrSQL & " AND PROV_TYPE LIKE " & iKey3 & " "
		   Case "B"
				lgStrSQL = "DELETE  H_IF_HDF040T "
				lgStrSQL = lgStrSQL & " WHERE PAY_YYMM = " & iKey2 & " "				
				lgStrSQL = lgStrSQL & " AND PROV_TYPE LIKE " & iKey3 & " "
		   Case "C"
				lgStrSQL = "DELETE  H_IF_HDF060T "
				lgStrSQL = lgStrSQL & " WHERE SUB_YYMM = " & iKey2 & " "				
				lgStrSQL = lgStrSQL & " AND SUB_TYPE LIKE " & iKey3 & " "
	End Select
   
	'response.write lgstrsql
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
    
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2)
    
	Dim iSelCount

	Select Case Mid(pDataType,1,1)
		   Case "M"
		
				iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
		   
				Select Case Mid(pDataType,2,1)
					   Case "R"
							lgStrSQL = "SELECT  DEPT_CD, EMP_NO, PAY_YYMM,  "
							lgStrSQL = lgStrSQL & "	dbo.ufn_GetCodeName(" & FilterVar("H0040", "''", "S") & ", PROV_TYPE) PROV_TYPE_NM, " 
							lgStrSQL = lgStrSQL & " PROV_TYPE, PROV_DT, PAY_TOT_AMT, BONUS_TOT_AMT, NONTAX_TOT_AMT, "
							lgStrSQL = lgStrSQL & " CASE WHEN PROV_TYPE = 1 THEN TAX_AMT ELSE BONUS_TAX END, PROV_TOT_AMT, "
							lgStrSQL = lgStrSQL & "  SUB_TOT_AMT, REAL_PROV_AMT, INCOME_TAX, RES_TAX, ANUT, MED_INSURE, EMP_INSURE "
							lgStrSQL = lgStrSQL & " FROM H_IF_HDF070T "
							lgStrSQL = lgStrSQL & " WHERE PROV_DT =  " & pCode & " "  
							lgStrSQL = lgStrSQL & "	AND PAY_YYMM =   " & pCode1 & " " 
							lgStrSQL = lgStrSQL & "	AND PROV_TYPE LIKE  " & pCode2 & " "
							lgStrSQL = lgStrSQL & " ORDER BY PROV_TYPE, PAY_YYMM, EMP_NO "	
					
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
                .lgStrPrevKey         = "<%=lgStrPrevKey%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
          	 Parent.lgChgStatus = <%=LngRecs%>
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
    
       
</Script>	
