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
    Dim StrDt, StrYYMM, StrProvCD, StrFileGubun

    call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgSvrDateTime = GetSvrDateTime

    Call HideStatusWnd                                                               '☜: Hide Processing message
	'Call ServerMesgBox("bbbbbbb" , vbInformation, I_MKSCRIPT)
	
    
	lgErrorStatus    = "NO"
    lgErrorPos       = ""                                                           '☜: Set to space
    lgOpModeCRUD     = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream      = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow      = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey	 = UNICInt(Trim(Request("lgStrPrevKey")),0)						'☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
         
	StrYYMM			= Request("htxtYYMM")
	StrProvCD		= Request("htxtProvCD")
	StrFileGubun	= Request("htxtFileGubun")

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
 
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update	
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete			
             Call SubBizDelete()
		Case CStr(UID_M0004)                                                         '☜: IF Query			
             Call SubBizIFQuery()
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
' Name : SubBizIFQuery
' Desc : Query Data from Db
'============================================================================================================
'Sub SubBizIFQuery()
'    On Error Resume Next                                                             '☜: Protect system from crashing
'    Err.Clear                                                                        '☜: Clear Error status
'    Call SubBizIFQueryMulti()
'End Sub 

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
    
	Dim txtKey   
    Dim iDx
    Dim iSelCount
    
    Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수   
    
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    	
    txtKey = ""
    txtKey = txtKey & " AND B.PAY_YYMM = " & FilterVar(Replace(lgKeyStream(0),"-",""),"''", "S") 
    txtKey = txtKey & " AND B.PROV_TYPE LIKE" & FilterVar(lgKeyStream(1) & "%", "''", "S")
	txtKey = txtKey & " AND A.EMP_NO LIKE" & FilterVar(lgKeyStream(2) & "%", "''", "S")


	iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
	
		
	
	Select Case StrFileGubun
		   Case "A"
				'lgStrSQL = "SELECT   A.NAME, A.EMP_NO, dbo.ufn_GetCodeName(B.DEPT_CD,'') DEPT_NM, B.DEPT_CD, "

				'lgStrSQL = "SELECT   A.NAME, A.EMP_NO, B.DEPT_CD DEPT_NM, B.DEPT_CD, "
				lgStrSQL = "SELECT  TOP " & iSelCount & " A.NAME, A.EMP_NO, A.DEPT_CD, dbo.ufn_getDeptName(a.DEPT_CD, " & FilterVar(Replace(lgKeyStream(0),"-","") & "01", "''", "S") & ") DEPT_NM, "
				lgStrSQL = lgStrSQL & "	dbo.ufn_GetCodeName(" & FilterVar("H0040", "''", "S") & ", B.PROV_TYPE) PROV_TYPE_NM, " 
				lgStrSQL = lgStrSQL & " B.PROV_TYPE, B.PROV_DT, "
				lgStrSQL = lgStrSQL & " CASE WHEN B.PROV_TYPE = 1 THEN B.PAY_TOT_AMT ELSE B.BONUS_TOT_AMT END PAY_TOT,  " 				
				lgStrSQL = lgStrSQL & " CASE WHEN B.PROV_TYPE = 1 THEN B.TAX_AMT     ELSE B.BONUS_TAX     END PAY_TAX, B.NON_TAX3, "
				lgStrSQL = lgStrSQL & " B.PROV_TOT_AMT, B.SUB_TOT_AMT, B.REAL_PROV_AMT, B.INCOME_TAX, B.RES_TAX, B.ANUT, B.MED_INSUR, B.EMP_INSUR, "
				lgStrSQL = lgStrSQL & " A.OCPT_TYPE, A.PAY_GRD1, A.PAY_GRD2, A.PAY_CD, A.TAX_CD, A.INTERNAL_CD "
				lgStrSQL = lgStrSQL & " FROM HDF020T A(NOLOCK) ,HDF070T B(NOLOCK) "
				lgStrSQL = lgStrSQL & " WHERE A.EMP_NO = B.EMP_NO " & txtKey & " "				
				lgStrSQL = lgStrSQL & " ORDER BY B.PROV_TYPE, A.EMP_NO, B.PROV_DT  "
				
			Case "B"
				
				lgStrSQL = "SELECT TOP " & iSelCount & " PAY_YYMM, EMP_NO,   "
				lgStrSQL = lgStrSQL & "	dbo.ufn_GetCodeName(" & FilterVar("H0040", "''", "S") & ", PROV_TYPE) PROV_TYPE_NM, PROV_TYPE, ALLOW_CD, " 
				lgStrSQL = lgStrSQL & "	dbo.ufn_H_GetCodeName(" & FilterVar("HDA010T", "''", "S") & ",ALLOW_CD,'')ALLOW_CD_NM, ALLOW " 				
				lgStrSQL = lgStrSQL & " FROM HDF040T(NOLOCK) "
				lgStrSQL = lgStrSQL & " WHERE PAY_YYMM =   " & FilterVar(Replace(lgKeyStream(0),"-",""),"''", "S") & " "			
				lgStrSQL = lgStrSQL & "	AND PROV_TYPE LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & " "				
				lgStrSQL = lgStrSQL & "	AND EMP_NO LIKE " & FilterVar(lgKeyStream(2) & "%", "''", "S") & " "						
				lgStrSQL = lgStrSQL & " ORDER BY PROV_TYPE, PAY_YYMM, EMP_NO "
		
		   Case "C"
				lgStrSQL = "SELECT TOP " & iSelCount & " SUB_YYMM, EMP_NO,   "
				lgStrSQL = lgStrSQL & "	dbo.ufn_GetCodeName(" & FilterVar("H0040", "''", "S") & ", SUB_TYPE) SUB_TYPE_NM, SUB_TYPE, SUB_CD, "
				lgStrSQL = lgStrSQL & "	dbo.ufn_H_GetCodeName(" & FilterVar("HDA010T", "''", "S") & ", SUB_CD,'') SUB_CD_NM, SUB_AMT " 				
				lgStrSQL = lgStrSQL & " FROM HDF060T(NOLOCK) "
				lgStrSQL = lgStrSQL & " WHERE SUB_YYMM =   " & FilterVar(Replace(lgKeyStream(0),"-",""),"''", "S") & " "					
				lgStrSQL = lgStrSQL & "	AND SUB_TYPE LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & " "				
				lgStrSQL = lgStrSQL & "	AND EMP_NO LIKE " & FilterVar(lgKeyStream(2) & "%", "''", "S") & " "
				lgStrSQL = lgStrSQL & " ORDER BY SUB_TYPE, SUB_YYMM, EMP_NO "
	End Select
	
	'response.write lgStrSQL
    
   
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs, CDbl(C_SHEETMAXROWS_D) * CDbl(lgStrPrevKey))

        lgstrData = ""
        iDx       = 1
       
		Select Case StrFileGubun
			   Case "A"
					Do While Not lgObjRs.EOF
						
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAME"))	
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_CD"))								
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROV_TYPE_NM"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROV_TYPE"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROV_DT"))
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PAY_TOT"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PAY_TAX"), ggAmtOfMoney.DecPoint,0)
						'lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("NONTAX_TOT_AMT"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("NON_TAX3"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PROV_TOT_AMT"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("SUB_TOT_AMT"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("REAL_PROV_AMT"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("INCOME_TAX"), ggAmtOfMoney.DecPoint,0)			
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("RES_TAX"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ANUT"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("MED_INSUR"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("EMP_INSUR"), ggAmtOfMoney.DecPoint,0)

						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("OCPT_TYPE"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_GRD1"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_GRD2"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_CD"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TAX_CD"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("INTERNAL_CD"))						

						lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
						lgstrData = lgstrData & Chr(11) & Chr(12)

						lgObjRs.MoveNext

						iDx =  iDx + 1
						
						If iDx > C_SHEETMAXROWS_D Then
						   lgStrPrevKey = lgStrPrevKey + 1
						   Exit Do
						End If   
						   
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
						
						If iDx > C_SHEETMAXROWS_D Then
						   lgStrPrevKey = lgStrPrevKey + 1
						   Exit Do
						End If   						   
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
						
						If iDx > C_SHEETMAXROWS_D Then
						   lgStrPrevKey = lgStrPrevKey + 1
						   Exit Do
						End If   						   
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
' Name : SubBizIFQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizIFQuery()
    
	Dim txtKey   

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    	
    txtKey = ""
    txtKey = txtKey & " AND B.PAY_YYMM = " & FilterVar(Replace(lgKeyStream(0),"-",""),"''", "S") 
    txtKey = txtKey & " AND B.PROV_TYPE LIKE " & FilterVar(lgKeyStream(1) & "%", "''", "S")
	txtKey = txtKey & " AND A.EMP_NO LIKE " & FilterVar(lgKeyStream(2) & "%", "''", "S")
	
	
	Select Case StrFileGubun
		   Case "A"
				'lgStrSQL = "SELECT   A.NAME, A.EMP_NO, dbo.ufn_GetCodeName(B.DEPT_CD,'') DEPT_NM, B.DEPT_CD, "

				lgStrSQL = "SELECT   A.NAME, A.EMP_NO, A.DEPT_CD, dbo.ufn_getDeptName(a.DEPT_CD, " & FilterVar(Replace(lgKeyStream(0),"-","") & "01", "''", "S") & ") DEPT_NM, "
				lgStrSQL = lgStrSQL & "	dbo.ufn_GetCodeName(" & FilterVar("H0040", "''", "S") & ", B.PROV_TYPE) PROV_TYPE_NM, " 
				lgStrSQL = lgStrSQL & " B.PROV_TYPE, B.PROV_DT, "
				lgStrSQL = lgStrSQL & " CASE WHEN B.PROV_TYPE = 1 THEN B.PAY_TOT_AMT ELSE B.BONUS_TOT_AMT END PAY_TOT,  " 				
				lgStrSQL = lgStrSQL & " CASE WHEN B.PROV_TYPE = 1 THEN B.TAX_AMT     ELSE B.BONUS_TAX     END PAY_TAX, B.NONTAX_TOT_AMT, "
				lgStrSQL = lgStrSQL & " B.PROV_TOT_AMT, B.SUB_TOT_AMT, B.REAL_PROV_AMT, B.INCOME_TAX, B.RES_TAX, B.ANUT, B.MED_INSURE, B.EMP_INSURE, "
				lgStrSQL = lgStrSQL & " A.OCPT_TYPE, A.PAY_GRD1, A.PAY_GRD2, A.PAY_CD, A.TAX_CD, A.INTERNAL_CD "
				lgStrSQL = lgStrSQL & " FROM HDF020T A(NOLOCK) , H_IF_HDF070T B(NOLOCK) "
				lgStrSQL = lgStrSQL & " WHERE A.EMP_NO = B.EMP_NO " & txtKey & " "				
				lgStrSQL = lgStrSQL & " ORDER BY B.PROV_TYPE, A.EMP_NO, B.PROV_DT  "
				
			Case "B"
				
				lgStrSQL = "SELECT  A.PAY_YYMM, A.EMP_NO,   "
				lgStrSQL = lgStrSQL & "	dbo.ufn_GetCodeName(" & FilterVar("H0040", "''", "S") & ", A.PROV_TYPE) PROV_TYPE_NM, A.PROV_TYPE, A.ALLOW_CD, " 
				lgStrSQL = lgStrSQL & "	dbo.ufn_H_GetCodeName(" & FilterVar("HDA010T", "''", "S") & ", A.ALLOW_CD,'') ALLOW_CD_NM, A.ALLOW " 				
				lgStrSQL = lgStrSQL & " FROM H_IF_HDF040T A(NOLOCK),"
				lgStrSQL = lgStrSQL & "      HDF020T      B(NOLOCK) "
				lgStrSQL = lgStrSQL & " WHERE A.EMP_NO = B.EMP_NO "
				lgStrSQL = lgStrSQL & " AND A.PAY_YYMM =   " & FilterVar(Replace(lgKeyStream(0),"-",""),"''", "S") & " "			
				lgStrSQL = lgStrSQL & "	AND A.PROV_TYPE LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & " "				
				lgStrSQL = lgStrSQL & "	AND A.EMP_NO LIKE " & FilterVar(lgKeyStream(2) & "%", "''", "S") & " "							
				lgStrSQL = lgStrSQL & " ORDER BY A.PROV_TYPE, A.PAY_YYMM, A.EMP_NO "
		
		   Case "C"
				lgStrSQL = "SELECT  A.SUB_YYMM, A.EMP_NO,   "
				lgStrSQL = lgStrSQL & "	dbo.ufn_GetCodeName(" & FilterVar("H0040", "''", "S") & ", A.SUB_TYPE) SUB_TYPE_NM, A.SUB_TYPE, A.SUB_CD, " 
				lgStrSQL = lgStrSQL & "	dbo.ufn_H_GetCodeName(" & FilterVar("HDA010T", "''", "S") & ", A.SUB_CD,'') SUB_CD_NM, A.SUB_AMT " 				
				lgStrSQL = lgStrSQL & " FROM H_IF_HDF060T A(NOLOCK),"
				lgStrSQL = lgStrSQL & "      HDF020T      B(NOLOCK) "
				lgStrSQL = lgStrSQL & " WHERE A.EMP_NO = B.EMP_NO "				
				lgStrSQL = lgStrSQL & " AND A.SUB_YYMM =   " & FilterVar(Replace(lgKeyStream(0),"-",""),"''", "S") & " "					
				lgStrSQL = lgStrSQL & "	AND A.SUB_TYPE LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & " "				
				lgStrSQL = lgStrSQL & "	AND A.EMP_NO LIKE " & FilterVar(lgKeyStream(2) & "%", "''", "S") & " "
				lgStrSQL = lgStrSQL & " ORDER BY A.SUB_TYPE, A.SUB_YYMM, A.EMP_NO "
	End Select
	
	'response.write lgStrSQL
    
   
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        iDx       = 1
       
		Select Case StrFileGubun
			   Case "A"
					Do While Not lgObjRs.EOF
						
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NAME"))			
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_CD"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROV_TYPE_NM"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROV_TYPE"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROV_DT"))
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PAY_TOT"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PAY_TAX"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("NONTAX_TOT_AMT"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PROV_TOT_AMT"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("SUB_TOT_AMT"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("REAL_PROV_AMT"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("INCOME_TAX"), ggAmtOfMoney.DecPoint,0)			
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("RES_TAX"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("ANUT"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("MED_INSURE"), ggAmtOfMoney.DecPoint,0)
						lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("EMP_INSURE"), ggAmtOfMoney.DecPoint,0)

						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("OCPT_TYPE"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_GRD1"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_GRD2"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAY_CD"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TAX_CD"))
						lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("INTERNAL_CD"))						

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
    
	Call SuSaveCreate()
	
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
  
	'지급구분이 급여이면 급여과세총액/상여이면 상여과세총액
	Select Case StrFileGubun
		   Case "A"
				Dim iValue1, iValue2 , iValue4, iValue5 
				If Trim(UCase(arrColVal(4))) = 1 Then
					iValue1 = arrColVal(6)				'급여총액
					iValue2 = 0							'상여총액
					iValue3 = arrColVal(7)				'급여과세총액
					iValue4 = 0							'상여과세총액
				Else
					iValue1 = 0
					iValue2 = arrColVal(6)
					iValue3 = 0
					iValue4 = arrColVal(7)
				End if
'Call ServerMesgBox("Trim(UCase(arrColVal(5))) = " & Trim(UCase(arrColVal(5))) , vbInformation, I_MKSCRIPT)
'Call ServerMesgBox("arrColVal(6) = " & arrColVal(6) , vbInformation, I_MKSCRIPT)
'Call ServerMesgBox("arrColVal(7) = " & arrColVal(7) , vbInformation, I_MKSCRIPT)
				 
				'NONTAX_TOT_AMT
				'lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(8)),0)					& ","
				'NON_TAX3 <= 비과세총액(NONTAX_TOT_AMT)
				lgStrSQL = "INSERT INTO HDF070T( PAY_YYMM, EMP_NO, DEPT_CD, PROV_TYPE, PROV_DT, PAY_TOT_AMT, BONUS_TOT_AMT, TAX_AMT, BONUS_TAX, NON_TAX3, "
				lgStrSQL = lgStrSQL & "         PROV_TOT_AMT, SUB_TOT_AMT, REAL_PROV_AMT, INCOME_TAX, RES_TAX, ANUT, MED_INSUR, EMP_INSUR, OCPT_TYPE, "
				lgStrSQL = lgStrSQL & "			PAY_GRD1, PAY_GRD2, PAY_CD, TAX_CD, INTERNAL_CD, ISRT_EMP_NO, UPDT_EMP_NO, MINUS2_RATE )"
				lgStrSQL = lgStrSQL & " VALUES(" 
				lgStrSQL = lgStrSQL & FilterVar(Replace(StrYYMM,"-",""),"''", "S")		& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(2))),"","S")		& ","				
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(3))),"","S")		& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(4))),"","S")		& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(5))),"","S")		& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(iValue1),0)						& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(iValue2),0)						& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(iValue3),0)						& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(iValue4),0)						& ","							
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(8)),0)					& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(9)),0)					& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(10)),0)					& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(11)),0)					& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(12)),0)					& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(13)),0)					& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(14)),0)					& ","
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(15)),0)					& ","		
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(16)),0)					& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(17)))," ","S")		& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(18)))," ","S")		& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(19)))," ","S")		& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(20)))," ","S")		& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(21)))," ","S")		& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(22)))," ","S")		& ","
				lgStrSQL = lgStrSQL & FilterVar(gUsrID,"","S")							& ","				
				lgStrSQL = lgStrSQL & FilterVar(gUsrID,"","S")							& ", 0 "								
				lgStrSQL = lgStrSQL & ")"
				
			    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
				Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
		   Case "B"								

				lgStrSQL = "INSERT INTO HDF040T(PAY_YYMM, EMP_NO, PROV_TYPE, ALLOW_CD, ALLOW, ISRT_EMP_NO, ISRT_DT, UPDT_EMP_NO, UPDT_DT )"
				lgStrSQL = lgStrSQL & " VALUES(" 
				lgStrSQL = lgStrSQL & FilterVar(Replace(StrYYMM,"-",""),"''", "S")		& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(3))),"","S")		& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(4))),"","S")		& ","				
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(5))),"","S")		& ","				
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(6)),0)					& ","
				lgStrSQL = lgStrSQL & FilterVar(gUsrId, "", "S")                        & ","     
				lgStrSQL = lgStrSQL & "GetDate(), "      
				lgStrSQL = lgStrSQL & FilterVar(gUsrId, "", "S")                        & ","     
				lgStrSQL = lgStrSQL & "GetDate() "        
				lgStrSQL = lgStrSQL & ")" 			
					
			    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
				Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
				
				'==============================================================================================
				' 수당테이블에 "A10(기본급)"이 존재하는 지 체크하고, 없으면 Insert 한다.
				'==============================================================================================
				lgStrSQL = " SELECT	*	"
				lgStrSQL = lgStrSQL & " FROM	HDF040T(NOLOCK)	"
				lgStrSQL = lgStrSQL & " WHERE	PAY_YYMM = " & FilterVar(Replace(StrYYMM,"-",""),"''", "S")
				lgStrSQL = lgStrSQL & " AND		PROV_TYPE = " & FilterVar(Trim(UCase(arrColVal(4))),"","S")
				lgStrSQL = lgStrSQL & " AND		EMP_NO = " & FilterVar(Trim(UCase(arrColVal(3))),"","S")
				lgStrSQL = lgStrSQL & " AND		ALLOW_CD = 'A10'"				
				
			    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then

					lgStrSQL = " "			    	
					lgStrSQL = lgStrSQL & " INSERT INTO HDF040T	(PAY_YYMM, PROV_TYPE, EMP_NO, "
					lgStrSQL = lgStrSQL & " 					 ALLOW_CD, ALLOW, ISRT_EMP_NO, "
					lgStrSQL = lgStrSQL & " 					 ISRT_DT, UPDT_EMP_NO, UPDT_DT) "
					lgStrSQL = lgStrSQL & " SELECT	PAY_YYMM, PROV_TYPE, EMP_NO, 'A10', PAY_TOT_AMT, "
					lgStrSQL = lgStrSQL &			FilterVar(gUsrId, "", "S") & ", "
					lgStrSQL = lgStrSQL &			"GetDate(), "
					lgStrSQL = lgStrSQL &			FilterVar(gUsrId, "", "S") & ", "
					lgStrSQL = lgStrSQL &			"GetDate() "
					lgStrSQL = lgStrSQL & " FROM	HDF070T(NOLOCK) "
					lgStrSQL = lgStrSQL & " WHERE	PAY_YYMM = " & FilterVar(Replace(StrYYMM,"-",""),"''", "S")
					lgStrSQL = lgStrSQL & " AND		PROV_TYPE = " & FilterVar(Trim(UCase(arrColVal(4))),"","S")
					lgStrSQL = lgStrSQL & " AND		EMP_NO = " & FilterVar(Trim(UCase(arrColVal(3))),"","S")		    	

				    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
					Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
					
			    End If				

		   Case "C"
				lgStrSQL = "INSERT INTO HDF060T(SUB_YYMM, EMP_NO, SUB_TYPE, SUB_CD, SUB_AMT, ISRT_EMP_NO, ISRT_DT, UPDT_EMP_NO, UPDT_DT  )"
				lgStrSQL = lgStrSQL & " VALUES(" 
				lgStrSQL = lgStrSQL & FilterVar(Replace(StrYYMM,"-",""),"''", "S")		& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(3))),"","S")		& ","
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(4))),"","S")		& ","				
				lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(5))),"","S")		& ","				
				lgStrSQL = lgStrSQL & UNIConvNum(Trim(arrColVal(6)),0)					& ","
				lgStrSQL = lgStrSQL & FilterVar(gUsrId, "", "S")                        & ","     
				lgStrSQL = lgStrSQL & "GetDate(), "      
				lgStrSQL = lgStrSQL & FilterVar(gUsrId, "", "S")                        & ","     
				lgStrSQL = lgStrSQL & "GetDate() "        
				lgStrSQL = lgStrSQL & ")" 	
				
			    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
				Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

				Call SubCreateCommandObject(lgObjComm)
				Call SubUspAddSub(Replace(StrYYMM,"-",""), Trim(UCase(arrColVal(4))), Trim(UCase(arrColVal(3))), gUsrId)
				Call SubCloseCommandObject(lgObjComm)

	End Select
	

End Sub

'============================================================================================================
' Name : SubUspAddSub
' Desc : 급/상여테이블(HDF070T)의 공제사항을 공제테이블(HDF060T)에 추가
'============================================================================================================
Sub SubUspAddSub(strPay_yymm, strProv_type, strEmp_no, strUSR_id)
    'Dim strMsg_cd

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status


    With lgObjComm
        .CommandText = "USP_HDF060T_ADD_SUB_KO441"
        .CommandType = adCmdStoredProc

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE" ,adInteger  ,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@PAY_YYMM"    ,adVarXChar ,adParamInput ,6  ,strPay_yymm)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@PROV_TYPE"   ,adVarXChar ,adParamInput ,1  ,strProv_type)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@EMP_NO"      ,adVarXChar ,adParamInput ,13 ,strEmp_no)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@USR_ID"      ,adVarXChar ,adParamInput ,13 ,strUSR_id)
        'lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adVarXChar ,adParamOutput,6)

        lgObjComm.Execute ,, adExecuteNoRecords
        
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        'if  IntRetCD < 0 then
            'strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            'Call DisplayMsgBox(strMsg_cd, vbInformation, strMsg_text, "", I_MKSCRIPT)
            'IntRetCD = -1
            'Exit Sub    
        'else          
        '    IntRetCD = 1
        'end if          
    Else           
        'call svrmsgbox(Err.Description, vbinformation, i_mkscript)
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
        'IntRetCD = -1
    End if

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
				lgStrSQL = "DELETE  HDF070T"
				lgStrSQL = lgStrSQL & " WHERE PAY_YYMM = " & FilterVar(Replace(lgKeyStream(0),"-",""),"''", "S") & " "
				lgStrSQL = lgStrSQL & " AND PROV_TYPE LIKE " & FilterVar(lgKeyStream(1) & "%", "''", "S") & " "
				lgStrSQL = lgStrSQL & " AND EMP_NO LIKE " & FilterVar(lgKeyStream(2) & "%", "''", "S") & " "
		   Case "B"
				lgStrSQL = "DELETE  HDF040T "
				lgStrSQL = lgStrSQL & " WHERE PAY_YYMM = " & FilterVar(Replace(lgKeyStream(0),"-",""),"''", "S") & " "
				lgStrSQL = lgStrSQL & " AND PROV_TYPE LIKE " & FilterVar(lgKeyStream(1) & "%", "''", "S") & " "
				lgStrSQL = lgStrSQL & " AND EMP_NO LIKE " & FilterVar(lgKeyStream(2) & "%", "''", "S") & " "
		   Case "C"
				lgStrSQL = "DELETE  HDF060T "
				lgStrSQL = lgStrSQL & " WHERE SUB_YYMM = " & FilterVar(Replace(lgKeyStream(0),"-",""),"''", "S") & " "			
				lgStrSQL = lgStrSQL & " AND SUB_TYPE LIKE " & FilterVar(lgKeyStream(1) & "%", "''", "S") & " "
				lgStrSQL = lgStrSQL & " AND EMP_NO LIKE " & FilterVar(lgKeyStream(2) & "%", "''", "S") & " "
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
							lgStrSQL = lgStrSQL & " FROM H_IF_HDF070T(NOLOCK) "
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
			  End With
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
	   Case "<%=UID_M0004%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
					.ggoSpread.Source     = .frm1.vspdData
					.lgStrPrevKey		  = "<%=lgStrPrevKey%>"
					.ggoSpread.SSShowData "<%=lgstrData%>"
					.lgStrPrevKey         = "<%=lgStrPrevKey%>"
					.DBAutoQueryOk    
			  End With
          End If  
    End Select    
   
</Script>	
