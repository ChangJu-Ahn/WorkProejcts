<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1411mb11.asp
'*  4. Program Name         : BOM변경이력 조회 
'*  5. Program Desc         :
'*  6. Component List        : 
'*  7. Modified date(First) : 2003/03/08
'*  8. Modified date(Last)  : 2003/03/08
'*  9. Modifier (First)     : NamkyuHo
'* 10. Modifier (Last)      : Park Kye Jin (Changed Field Added) (2003.04.10)
'* 11. Comment              :
'**********************************************************************************************-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/AdoVbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDB.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%

On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Const C_SHEETMAXROWS_D = 30

Call HideStatusWnd                                                               '☜: Hide Processing message

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "P", "NOCOOKIE", "MB")

Dim strPlantCd
Dim strBomNo
Dim strChgFromDt
Dim strChgToDt
Dim strECNNo
Dim strItemCd
Dim strChildItemCd

Dim TmpBuffer
Dim iTotalStr

Dim	SaveChangedField
Dim	SavePrntItemCd
Dim	SaveChildItemCd		
Dim	SaveChildItemQty	
Dim	SaveChildUnit		
Dim	SavePrntItemQty		
Dim	SavePrntUnit		
Dim	SaveSafetyLT		
Dim	SaveLossRate		
Dim	SaveSupplyFlg
Dim SaveValidFromDt
Dim	SaveValidToDt
Dim NewActionFlg
Dim	NewPrntItemCd
Dim	NewChildItemCd		
Dim	NewChildItemQty	
Dim	NewChildUnit		
Dim	NewPrntItemQty		
Dim	NewPrntUnit		
Dim	NewSafetyLT		
Dim	NewLossRate		
Dim	NewSupplyFlg	
Dim NewValidFromDt	
Dim	NewValidToDt		

'---------------------------------------Common-----------------------------------------------------------
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space

'------ Developer Coding part (Start ) ------------------------------------------------------------------

lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
lgMaxCount        = C_SHEETMAXROWS_D                                  '☜: Fetch count at a time for VspdData
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
ReDim TmpBuffer(0)
    
'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
Call SubOpenDB(lgObjConn)                    
Call SubCreateCommandObject(lgObjComm)
     
Call SubMakeParameter()
Call SubBomExplode()
     
Call SubBizQuery()

Call SubCloseCommandObject(lgObjComm)
Call SubCloseDB(lgObjConn)       

Response.Write "<Script Language = VBScript>" & vbCrLf
	If Trim(lgErrorStatus) = "NO" And IntRetCd <> -1 Then
	    Response.Write "With Parent" & vbCrLf
	       Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
	       Response.Write ".lgStrPrevKeyIndex = """ & lgStrPrevKeyIndex & """" & vbCrLf
	       Response.Write ".ggoSpread.SSShowDataByClip """ & ConvSPChars(iTotalStr) & """" & vbCrLf
	'                Response.Write ".lgStrPrevKey = """ & lgStrPrevKey & """" & vbCrLf
	       Response.Write ".DBQueryOk()" & vbCrLf        
		Response.Write "End with" & vbCrLf
	End If
Response.Write "</Script>" & vbCrLf

Response.End

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	
	Dim iDx
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                          
	
	strPlantCd		= FilterVar(Request("txtPlantCd"), "''", "S")
	strBomNo		= FilterVar(Request("txtBomNo"), "''", "S")
	strChgFromDt	= FilterVar(Trim(UniConvDate(Request("txtChgFromDt"))) & " 00:00:00", "''", "S")
	strChgToDt		= FilterVar(Trim(UniConvDate(Request("txtChgToDt"))) & " 23:59:59", "''", "S")
    strECNNo		= FilterVar(Request("txtECNNo"), "''", "S")
	strItemCd		= FilterVar(Request("txtItemCd"), "''", "S")
    strChildItemCd  = FilterVar(Request("txtChildItemCd"), "''", "S")

	'--------------
	'공장 체크		
	'--------------	
	lgStrSQL = ""
	Call SubMakeSQLStatements("P_CK",strPlantCd,"","","","","","")           '☜ : Make sql statements
			
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    
		IntRetCD = -1
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()
		Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.Frm1.txtPlantNm.Value  = """"" & vbCrLf   'Set condition area
		Response.Write "parent.Frm1.txtPlantCd.focus" & vbCrLf   'Set condition area
		Response.Write "</Script>" & vbcRLf
		Response.End
	Else
		IntRetCD = 1
		Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.Frm1.txtPlantNm.Value = """ & ConvSPChars(lgObjRs(1)) & """" & vbCrLf 'Set condition area
		Response.Write "</Script>" & vbcRLf
	End If
		
	Call SubCloseRs(lgObjRs) 
	
	'------------------
	' bom type 체크 
	'------------------
	lgStrSQL = ""
			
	Call SubMakeSQLStatements("BT_CK",strBomNo,"","","","","","")           '☜ : Make sql statements
			
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    
		IntRetCD = -1
				
		Call DisplayMsgBox("182622", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()
		Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtBomNo.focus" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		Response.End							
	Else
		IntRetCD = 1
	End If
		
	Call SubCloseRs(lgObjRs) 
	
	'------------------
	'변경일체크 
	'------------------
	IF strChgToDt = FilterVar("1900-01-01 23:59:59", "''", "S") THEN strChgToDt = FilterVar("2999-12-31 23:59:59", "''", "S")
	
	IF strChgFromDt > strChgToDt THEN
		Call DisplayMsgBox("800111", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()
		
		Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.Frm1.txtChgFromDt.Focus" & vbCrLf 
		Response.Write "</Script>" & vbcRLf
		Response.End

		Call SubCloseRs(lgObjRs) 
	END IF

	'------------------
	'설계변경번호체크 
	'------------------
	IF strECNNo = FilterVar("", "''", "S") THEN
		Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.Frm1.txtECNNoDesc.Value  = """"" & vbCrLf		'Set condition area
		Response.Write "</Script>" & vbcRLf
	ELSE
		lgStrSQL = ""
		Call SubMakeSQLStatements("ECN_CK",strECNNo,"","","","","","")           '☜ : Make sql statements
				
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
			    
			IntRetCD = -1
					
			Call DisplayMsgBox("182801", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
			Call SetErrorStatus()
			
			Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.Frm1.txtECNNoDesc.Value  = """"" & vbCrLf   'Set condition area
			Response.Write "parent.frm1.txtECNNo.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			Response.End							
		Else
			IntRetCD = 1
			Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.Frm1.txtECNNoDesc.Value = """ & ConvSPChars(Trim(lgObjRs(1))) & """" & vbCrLf
			Response.Write "</Script>" & vbcRLf
		End If
			
		Call SubCloseRs(lgObjRs) 
	END IF

	'------------------
	'품목체크 
	'------------------
	IF strItemCd = FilterVar("", "''", "S") THEN
		Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.Frm1.txtItemNm.Value  = """"" & vbCrLf		'Set condition area
		Response.Write "</Script>" & vbcRLf
	ELSE
		lgStrSQL = ""
		Call SubMakeSQLStatements("I_CK",strPlantCd,strItemCd,"","","","","")          '☜ : Make sql statements

		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   'If data not exists
			    
			IntRetCD = -1

			Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
			Call SetErrorStatus()

			Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.Frm1.txtItemNm.Value  = """"" & vbCrLf   'Set condition area
			Response.Write "parent.Frm1.txtItemCd.Focus" & vbCrLf 
			Response.Write "</Script>" & vbcRLf
			Response.End
		Else
			IntRetCD = 1
			Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.Frm1.txtItemNm.Value = """ & ConvSPChars(Trim(lgObjRs(0))) & """" & vbCrLf
			Response.Write "</Script>" & vbcRLf
		End If
			
		Call SubCloseRs(lgObjRs) 

	END IF

	'------------------
	'자품목체크 
	'------------------
	IF strChildItemCd = FilterVar("", "''", "S") THEN
		Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.Frm1.txtChildItemNm.Value  = """"" & vbCrLf   'Set condition area
		Response.Write "</Script>" & vbcRLf
	ELSE
		lgStrSQL = ""
		Call SubMakeSQLStatements("I_CK",strPlantCd,strChildItemCd,"","","","","")          '☜ : Make sql statements

		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
			    
			IntRetCD = -1
					
			Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
			Call SetErrorStatus()

			Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.Frm1.txtChildItemNm.Value  = """"" & vbCrLf   'Set condition area
			Response.Write "parent.Frm1.txtChildItemCd.Focus" & vbCrLf 
			Response.Write "</Script>" & vbcRLf
			Response.End
		Else
			IntRetCD = 1
			Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.Frm1.txtChildItemNm.Value = """ & ConvSPChars(Trim(lgObjRs(0))) & """" & vbCrLf
			Response.Write "</Script>" & vbcRLf
		End If
	END IF
		
	Call SubCloseRs(lgObjRs) 

	'---------------------------
	' BOM 이력 결과 조회 
	'---------------------------    
	lgStrSQL = ""
			
	Call SubMakeSQLStatements("E", strPlantCd, strBomNo, strChgFromDt, strChgToDt, strECNNo, strItemCd, strChildItemCd)           '☜ : Make sql statements

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
	
		IntRetCD = -1
				
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()
		Response.End 
	Else
		IntRetCD = 1
		 Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

        lgstrData = ""
        iDx       = 1

        Do While Not lgObjRs.EOF
			'--------------------
			' 변경된 필드 체크 
			'--------------------
			SaveChangedField = ""
			NewActionFlg	= Trim(lgObjRs(3))
			NewPrntItemCd	= Trim(UCase(lgObjRs(0)))
			NewChildItemCd	= Trim(UCase(lgObjRs(8)))
			NewChildItemQty	= UniConvNumberDBToCompany(lgObjRs(13), 6, 3, "", 0)
			NewChildUnit	= Trim(UCase(lgObjRs(14)))
			NewPrntItemQty	= UniConvNumberDBToCompany(lgObjRs(15), 6, 3, "", 0)
			NewPrntUnit		= Trim(UCase(lgObjRs(16)))
			NewSafetyLT		= UniConvNumberDBToCompany(lgObjRs(17), 2, 2, "", 0)
			NewLossRate		= UniConvNumberDBToCompany(lgObjRs(18), 10, 8, "", 0)
			NewSupplyFlg	= Trim(UCase(lgObjRs(19)))
			NewValidFromDt	= UNIDateClientFormat(lgObjRs(20))
			NewValidToDt	= UNIDateClientFormat(lgObjRs(21))	

			If NewActionFlg = "Change" Then
				If NewPrntItemCd = SavePrntItemCd And NewChildItemCd = SaveChildItemCd Then
					If NewChildItemQty <> SaveChildItemQty Then	SaveChangedField = SaveChangedField & "/자품목기준수"
					If NewChildUnit <> SaveChildUnit Then SaveChangedField = SaveChangedField & "/자품목단위"
					If NewPrntItemQty <> SavePrntItemQty Then SaveChangedField = SaveChangedField & "/모품목기준수"
					If NewPrntUnit <> SavePrntUnit Then	SaveChangedField = SaveChangedField & "/모품목단위"
					If NewSafetyLT <> SaveSafetyLT Then	SaveChangedField = SaveChangedField & "/안전L/T"
					If NewLossRate <> SaveLossRate Then	SaveChangedField = SaveChangedField & "/Loss율"
					If NewSupplyFlg <> SaveSupplyFlg Then SaveChangedField = SaveChangedField & "/유무상구분"
					If CDate(NewValidFromDt) <> CDate(SaveValidFromDt) Then SaveChangedField = SaveChangedField & "/시작일"
					If CDate(NewValidToDt) <> CDate(SaveValidToDt) Then	SaveChangedField = SaveChangedField & "/종료일"
					
					If Trim(SaveChangedField) = "" Then SaveChangedField = "Nothing"
				Else
					SaveChangedField = NewActionFlg
				End If
			Else	'Add or Delete
				SaveChangedField	= NewActionFlg
			End If

			SavePrntItemCd		= NewPrntItemCd
			SaveChildItemCd		= NewChildItemCd
			SaveChildItemQty	= NewChildItemQty
			SaveChildUnit		= NewChildUnit
			SavePrntItemQty		= NewPrntItemQty
			SavePrntUnit		= NewPrntUnit
			SaveSafetyLT		= NewSafetyLT
			SaveLossRate		= NewLossRate
			SaveSupplyFlg		= NewSupplyFlg
			SaveValidFromDt		= NewValidFromDt
			SaveValidToDt		= NewValidToDt

			'--------------------
			' 데이터 세팅 
			'--------------------
			lgstrData = ""
            lgstrData = lgstrData & Chr(11) & lgObjRs(0)					' 품목 
			lgstrData = lgstrData & Chr(11) & lgObjRs(1)					' 품목명 
			lgstrData = lgstrData & Chr(11) & lgObjRs(2)					' 규격 
            lgstrData = lgstrData & Chr(11) & lgObjRs(3)					' 변경구분 
            lgstrData = lgstrData & Chr(11) & UniDateClientFormat(lgObjRs(4)) & "  " & FormatDateTime(lgObjRs(4),3) ' 변경일 
            lgstrData = lgstrData & Chr(11) & lgObjRs(5)					' 변경ID
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs(6), 6, 0, "", 0) ' 변경순서 
			lgstrData = lgstrData & Chr(11) & lgObjRs(7)					' 순서 
            lgstrData = lgstrData & Chr(11) & lgObjRs(8)					' 자품목 
            lgstrData = lgstrData & Chr(11) & lgObjRs(9)					' 자품목명 
            lgstrData = lgstrData & Chr(11) & lgObjRs(10)					' 자품목규격 
            lgstrData = lgstrData & Chr(11) & lgObjRs(11)					' 품목계정 
            lgstrData = lgstrData & Chr(11) & lgObjRs(12)					' 조달구분 
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs(13), 15, 4, "", 0)	' 자품목기준수 
            lgstrData = lgstrData & Chr(11) & lgObjRs(14)					' 단위 
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs(15), 15, 4, "", 0)	' 모품목기준수 
            lgstrData = lgstrData & Chr(11) & lgObjRs(16)					' 단위 
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs(17), 2, 2, "", 0)	' 안전L/T
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs(18), 10, 8, "", 0)	' LOSS율 
            lgstrData = lgstrData & Chr(11) & lgObjRs(19)					' 유무상구분 
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(20))	' 시작일 
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(21))	' 종료일 
            lgstrData = lgstrData & Chr(11) & lgObjRs(22)					' 설계변경번호 
            lgstrData = lgstrData & Chr(11) & lgObjRs(23)					' 설계변경내용 
            lgstrData = lgstrData & Chr(11) & lgObjRs(24)					' 설계변경근거 
            lgstrData = lgstrData & Chr(11) & lgObjRs(25)					' 비고            
            lgstrData = lgstrData & Chr(11) & SaveChangedField				' 변경된필드 
            lgstrData = lgstrData & Chr(11) & lgObjRs(4)					'변경일(정렬을 위한)
          
            
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
            
		    lgObjRs.MoveNext
			
			ReDim Preserve TmpBuffer(iDx - 1)
			
			TmpBuffer(iDx - 1) = lgstrData
			
            iDx =  iDx + 1
            If iDx > lgMaxCount Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
   				SavePrntItemCd		= lgObjRs(0)
				SaveChildItemCd		= lgObjRs(8)
				SaveChildItemQty	= lgObjRs(13)
				SaveChildUnit		= lgObjRs(14)
				SavePrntItemQty		= lgObjRs(15)
				SavePrntUnit		= lgObjRs(16)
				SaveSafetyLT		= lgObjRs(17)
				SaveLossRate		= lgObjRs(18)
				SaveSupplyFlg		= lgObjRs(19)
				SaveValidToDt		= lgObjRs(21)		

               Exit Do
            End If
        Loop 
    End If
	
	iTotalStr = Join(TmpBuffer, "")
	
    If iDx <= lgMaxCount Then
       lgStrPrevKeyIndex = ""
        SavePrntItemCd		= ""
		SaveChildItemCd		= ""
		SaveChildItemQty	= ""
		SaveChildUnit		= ""
		SavePrntItemQty		= ""
		SavePrntUnit		= ""
		SaveSafetyLT		= ""
		SaveLossRate		= ""
		SaveSupplyFlg		= ""
		SaveValidToDt		= ""
    End If   

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                             
		
	lgStrSQL = ""
    
End Sub	    
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2,pCode3,pCode4,pCode5,pCode6)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
	Dim iSelCount
	
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Select Case pDataType
		Case "E"
			lgStrSQL = " SELECT A.PRNT_ITEM_CD, D.ITEM_NM, D.SPEC, "
			lgStrSQL = lgStrSQL & " CASE WHEN A.ACTION_FLG = " & FilterVar("A", "''", "S") & "  THEN " & FilterVar("Add", "''", "S") & " WHEN A.ACTION_FLG = " & FilterVar("C", "''", "S") & "  THEN " & FilterVar("Change", "''", "S") & " WHEN A.ACTION_FLG = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("Delete", "''", "S") & " END, "
			lgStrSQL = lgStrSQL & " A.INSRT_DT, A.INSRT_USER_ID, A.CHANGE_SEQ,A.CHILD_ITEM_SEQ, "
			lgStrSQL = lgStrSQL & " A.CHILD_ITEM_CD, E.ITEM_NM, E.SPEC, "
			lgStrSQL = lgStrSQL & "  dbo.ufn_GetCodeName('P1001',  C.ITEM_ACCT), "
			lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName('P1003',  C.PROCUR_TYPE), "
			lgStrSQL = lgStrSQL & " A.CHILD_ITEM_QTY, A.CHILD_ITEM_UNIT, A.PRNT_ITEM_QTY, A.PRNT_ITEM_UNIT, A.SAFETY_LT, A.LOSS_RATE, "
			lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName('M2201',  A.SUPPLY_TYPE), "
			lgStrSQL = lgStrSQL & " A.VALID_FROM_DT, A.VALID_TO_DT, B.ECN_NO, B.ECN_DESC,"
			lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName('P1402',  B.REASON_CD), "
			lgStrSQL = lgStrSQL & " A.REMARK "
			lgStrSQL = lgStrSQL & " FROM P_BOM_HISTORY A, "
			lgStrSQL = lgStrSQL & " P_ECN_MASTER B, "
			lgStrSQL = lgStrSQL & " B_ITEM_BY_PLANT C, "
			lgStrSQL = lgStrSQL & " B_ITEM D, B_ITEM E "
			lgStrSQL = lgStrSQL & " WHERE A.ECN_NO = B.ECN_NO "
			lgStrSQL = lgStrSQL & " AND A.PLANT_CD = C.PLANT_CD"
			lgStrSQL = lgStrSQL & " AND A.CHILD_ITEM_CD = C.ITEM_CD"
			lgStrSQL = lgStrSQL & " AND A.PRNT_ITEM_CD = D.ITEM_CD"
			lgStrSQL = lgStrSQL & " AND A.CHILD_ITEM_CD = E.ITEM_CD"
			lgStrSQL = lgStrSQL & " AND A.PLANT_CD = " & pCode
			lgStrSQL = lgStrSQL & " AND A.PRNT_BOM_NO = " & pCode1
			lgStrSQL = lgStrSQL & " AND A.INSRT_DT BETWEEN " & pCode2 & " AND " & pCode3
			
			IF Trim(strECNNo) = FilterVar("", "''", "S") THEN
			   lgStrSQL = lgStrSQL
			ELSE
			   lgStrSQL = lgStrSQL & " AND A.ECN_NO = " & pCode4
			END IF

			IF Trim(strItemCd) = FilterVar("", "''", "S") THEN
			   lgStrSQL = lgStrSQL
			ELSE
			   lgStrSQL = lgStrSQL & " AND A.PRNT_ITEM_CD = " & pCode5
			END IF

			IF Trim(strChildItemCd) = FilterVar("", "''", "S") THEN
			   lgStrSQL = lgStrSQL
			ELSE
			   lgStrSQL = lgStrSQL & " AND A.CHILD_ITEM_CD = " & pCode6
			END IF
			
			lgStrSQL = lgStrSQL & " ORDER BY A.PRNT_ITEM_CD, A.CHILD_ITEM_SEQ, A.CHILD_ITEM_CD, A.CHANGE_SEQ " 			

		Case "BT_CK"
			lgStrSQL = "SELECT * FROM b_minor WHERE major_cd = " & FilterVar("P1401", "''", "S") & ""
			lgStrSQL = lgStrSQL & " AND minor_cd = " & pCode 

		Case "I_CK"
			lgStrSQL = "SELECT b.item_nm FROM b_item_by_plant a, b_item b, b_minor c, b_minor d "
			lgStrSQL = lgStrSQL & " WHERE a.item_cd =b.item_cd and c.minor_cd = a.item_acct and d.minor_cd = a.procur_type and c.major_cd =" & FilterVar("p1001", "''", "S") & " and d.major_cd =" & FilterVar("p1003", "''", "S") & ""
			lgStrSQL = lgStrSQL & " AND a.plant_cd = " & pCode 
			lgStrSQL = lgStrSQL & " AND a.item_cd = " & pCode1

		Case "P_CK"
			lgStrSQL = "SELECT * FROM b_plant where plant_cd = " & pCode 

		Case "ECN_CK"
			lgStrSQL = "SELECT * FROM p_ecn_master where ecn_no = " & pCode 

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
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status
    If CheckSYSTEMError(pErr,True) = True Then
       ObjectContext.SetAbort
       Call SetErrorStatus
    Else
       If CheckSQLError(pConn,True) = True Then
          ObjectContext.SetAbort
          Call SetErrorStatus
       End If
   End If

End Sub
%>
