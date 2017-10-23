<%@LANGUAGE = VBScript%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : Multi Sample
'*  3. Program ID           : p1501mb9
'*  4. Program Name         : p1501mb9
'*  5. Program Desc         : 자원조회 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2001/11/27
'*  8. Modified date(Last)  : 2003/01/28
'*  9. Modifier (First)     : Jung Yu Kyung
'* 10. Modifier (Last)      : Ryu Sung Won
'* 11. Comment              :
'=======================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->

<!-- #Include file="../../inc/adoVbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "MB")

Dim pPB6S101
Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

Dim R1_P_Plant
	Const EA_b_plant_plant_cd = 0
	Const EA_b_plant_plant_nm = 1
	Const EA_b_plant_cur_cd = 2

Const C_SHEETMAXROWS_D  = 50                                          '☜: Fetch count at a time

Dim TmpBuffer
Dim iTotalStr

Call HideStatusWnd   

On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

'---------------------------------------Common-----------------------------------------------------------
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space
lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
lgMaxCount        = C_SHEETMAXROWS_D
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

'-----------------------
'Com action area
'-----------------------
Set pPB6S101 = Server.CreateObject("PB6S101.cBLkUpPlt")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPB6S101.B_LOOK_UP_PLANT(gStrGlobalCollection, _
							lgKeyStream(0), _
							, _
							,_
							R1_P_Plant)

If CheckSYSTEMError(Err, True) = True Then
	
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "	Parent.frm1.txtPlantNm.Value = """"" & vbCr
	Response.Write "	Parent.frm1.txtCurCd.Value = """"" & vbCr
	Response.Write "	Parent.frm1.txtPlantCd.focus() " & vbCr
	Response.Write "</Script>" & vbCr
	
	Set pPB6S101 = Nothing															'☜: Unload Component
	Response.End
	
Else

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "	Parent.frm1.txtPlantNm.Value = """ & R1_P_Plant(EA_b_plant_plant_nm) & """" & vbCr
	Response.Write "	Parent.frm1.txtCurCd.Value = """ & R1_P_Plant(EA_b_plant_cur_cd) & """" & vbCr
	Response.Write "</Script>" & vbCr
	
	Set pPB6S101 = Nothing															'☜: Unload Component
	
End If

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
Call SubBizQueryMulti()
Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim strPlantCd
    Dim str_start_dt
    Dim str_end_dt
    Dim strRunRccp
    Dim strRunCrp
    
    Dim Rvalue
	
	on error resume next
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------   
    strPlantCd = FilterVar(lgKeyStream(0), "''", "S")
    str_start_dt =  FilterVar(UNIConvDate(lgKeyStream(1)), "''", "S")
    str_end_dt =  FilterVar(UNIConvDate(lgKeyStream(2)), "''", "S")

    If Trim(lgKeyStream(3)) = "A" Then
		strRunRccp = ""
	Else
		strRunRccp = FilterVar(lgKeyStream(3), "''", "S")
	End If

    Call SubMakeSQLStatements(strPlantCd,str_start_dt,str_end_dt,strRunRccp,strRunCrp)                                 '☆ : Make sql statements

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
            lgstrData = lgstrData & Chr(11) & UCase(Trim(lgObjRs("resource_cd")))
            lgstrData = lgstrData & Chr(11) & lgObjRs("description")
            lgstrData = lgstrData & Chr(11) & UCase(Trim(lgObjRs("resource_group_cd")))
            lgstrData = lgstrData & Chr(11) & lgObjRs("rg_nm")
            lgstrData = lgstrData & Chr(11) & FuncCodeName(1,"P1502",lgObjRs("resource_type"))
            lgstrData = lgstrData & Chr(11) & udf_UniConvNumberDBToCompany(lgObjRs("no_of_resource"),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & udf_UniConvNumberDBToCompany(lgObjRs("efficiency"),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & udf_UniConvNumberDBToCompany(lgObjRs("utilization"),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UCase(Trim(lgObjRs("run_rccp")))
            lgstrData = lgstrData & Chr(11) & UCase(Trim(lgObjRs("run_crp")))
            lgstrData = lgstrData & Chr(11) & udf_UniConvNumberDBToCompany(lgObjRs("overload_tol"),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & udf_UniConvNumberDBToCompany(lgObjRs("rsc_base_qty"),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UCase(ConvSPChars(lgObjRs("rsc_base_unit")))
            lgstrData = lgstrData & Chr(11) & udf_UniConvNumberDBToCompany(lgObjRs("mfg_cost"),ggUnitCost.DecPoint,ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("valid_from_Dt"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("valid_to_Dt"))
            
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
    Call SubCloseRs(lgObjRs) 
End Sub    

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pCode,pCode1,pCode2,pCode3,pCode4)
    Dim iSelCount

	iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1

    lgStrSQL = " Select TOP " & iSelCount & " a.*, b.description rg_nm"
    lgStrSQL = lgStrSQL & " From p_resource a, p_resource_group b "
    lgStrSQL = lgStrSQL & " WHERE a.resource_group_cd = b.resource_group_cd"
    lgStrSQL = lgStrSQL & " AND a.plant_cd = " & pCode
    lgStrSQL = lgStrSQL & " AND a.valid_to_Dt >= " & pCode1
    lgStrSQL = lgStrSQL & " AND a.valid_to_dt <= " & pCode2
    If pCode3 <> "" Then
		lgStrSQL = lgStrSQL & " AND a.run_rccp = " & pCode3
    End If
    If pCode4 <> "" Then
		lgStrSQL = lgStrSQL & " AND a.run_crp = " & pCode4
    End If
    lgStrSQL = lgStrSQL & " ORDER BY a.resource_cd ASC, a.resource_group_cd ASC "
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
	           Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
	           ObjectContext.SetAbort
	           Call SetErrorStatus
	        Else
	           If CheckSQLError(pConn,True) = True Then
	'                       Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
	              ObjectContext.SetAbort
	              Call SetErrorStatus
	           End If
	        End If
        Case "MD"
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

'==============================================================================
' 사용자 정의 서버 함수 
'==============================================================================
'==============================================================================
' Function Name : udf_UniConvNumberDBToCompany
' Function Desc : 최대값에 대해 udf_UniConvNumberDBToCompany함수를 사용하면 
'				반올림 정책에 따라 최대값을 넘어 가는 것을 방지 하기 위한 함수 
' 
'==============================================================================
Function udf_UniConvNumberDBToCompany(ByVal pNum,ByVal pDecPoint,ByVal pRndPolicy, ByVal pRndUnit, ByVal pDefault)

	Dim rtnNum
	
	Const maxNum	= 99999999999.9999	'최대값 (필드의 속성에 따라 변경 가능)
	Const maxDecPnt = 4					'소수점 이하 최대자리수 (필드의 속성에 따라 변경 가능,시스템 기준정보에서 적용가능한 최대자리수)
	
	rtnNum = UniConvNumberDBToCompany(pNum, pDecPoint, pRndPolicy, pRndUnit, pDefault)
	
	If rtnNum > UniConvNumberDBToCompany(maxNum,pDecPoint, pRndPolicy, pRndUnit,pDefault) Then	'최대값보다 큰 값일때 적용 
		If pDecPoint <> maxDecPnt Then							'소수점 이하 최대값이 아닐때만 적용 
			rtnNum = int(cdbl(pNum) * cdbl(10 ^ pDecPoint))
			rtnNum = rtnNum * cdbl(pRndUnit) * 10
		End if
		udf_UniConvNumberDBToCompany = UniConvNumberDBToCompany(rtnNum,pDecPoint, pRndPolicy, pRndUnit,pDefault)
	Else
		udf_UniConvNumberDBToCompany = rtnNum
	End If
	
End Function

%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
                .ggoSpread.SSShowDataByClip "<%=ConvSPChars(iTotalStr)%>"
                '.lgStrPrevKey         = 0 
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
