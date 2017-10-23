
<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
'======================================================================================================
'*  1. Module Name          : Interface관리
'*  2. Function Name        : MES Interface 전송관리
'*  3. Program ID           : XI219MB1_KO119
'*  4. Program Name         : 자재LOT정보등록
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2006.05.20
'*  8. Modified date(Last)  : 2006.05.20
'*  9. Modifier (First)     : TGS
'* 10. Modifier (Last)      : TGS
'* 11. Comment              :
'=======================================================================================================

	Dim lgOpModeCRUD
	
	Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
    
    Call HideStatusWnd                                                               '☜: Hide Processing message

	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status

    '---------------------------------------Common-----------------------------------------------------------
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    lgOpModeCRUD = Request("txtMode")                                                '☜: Read Operation Mode (CRUD)
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()

        Case CStr(UID_M0002)                                                         '☜: Save,Update
			Call SubBizSaveMulti()

        Case CStr(UID_M0003)                                                         '☜: Delete
             
    End Select

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
	On Error Resume Next                                                                 '☜: Protect system from crashing
	Err.Clear                                                                            '☜: Clear Error status

	Const C1_SHEETMAXROWS_D  = 100

	'IMPORT GROUP
	Const C1_IG1_sheet_max_rows = 0   'SELECT MAXCOUNT
	Const C1_IG1_print_fr_dt    = 1   '발행기간시작
	Const C1_IG1_print_to_dt    = 2   '발행기간종료
	Const C1_IG1_item_cd        = 3   '품목코드
	Const C1_IG1_bp_cd          = 4   '납품처
	Const C1_IG1_lot_no         = 5   'Lot번호
	Const C1_IG1_del_flag       = 6   '삭제여부
	Const C1_IG1_mes_rcv_flag   = 7   'MES수신여부
	
	Const C1_IT1_print_fr_dt    = 0   '발행기간시작
	Const C1_IT1_print_to_dt    = 1   '발행기간종료
	Const C1_IT1_item_cd        = 2   '품목코드
	Const C1_IT1_bp_cd          = 3   '납품처
	Const C1_IT1_lot_no         = 4   'Lot번호
	Const C1_IT1_del_flag       = 5   '삭제여부
	Const C1_IT1_mes_rcv_flag   = 6   'MES수신여부
	
	Const C1_IG2_next_item_cd     = 0  '품목코드
	Const C1_IG2_next_mat_lot_no  = 1  'LOT번호
	Const C1_IG2_next_seller_cd   = 2  '납품처코드
	Const C1_IG2_next_print_dt    = 3  '발행일자
	Const C1_IG2_next_rcpt_dt     = 4  '납입(입고)일자
	Const C1_IG2_next_rcpt_tm     = 5  '납입(입고)시간
	Const C1_IG2_next_create_type = 6  '생성구분(A:생성,B:수정,C:취소)
	
	'EXPORT GROUP
	Const C1_EG_item_cd          = 0   '품목코드(PK)
	Const C1_EG_item_nm          = 1   '품명
	Const C1_EG_mat_lot_no       = 2   'LOT번호(PK)
	Const C1_EG_seller_cd        = 3   '납품처코드(PK)
	Const C1_EG_seller_nm        = 4   '납품처명 
	Const C1_EG_print_dt         = 5   '발행일자(PK)
	Const C1_EG_rcpt_dt          = 6   '납입(입고)일자(PK)
	Const C1_EG_rcpt_tm          = 7   '납입(입고)시간(PK)
	Const C1_EG_bp_issue_no      = 8   '납품처 발행번호
	Const C1_EG_issue_flag       = 9   '발행구분(G:정산, E:긴급, C:취소)
	Const C1_EG_plant_flag       = 10  '공장구분
	Const C1_EG_plant_cd         = 11  '공장코드(한솔 내부 공장코드)
	Const C1_EG_gate_cd          = 12  'Gate Code
	Const C1_EG_snp              = 13  '포장 단위 수량
	Const C1_EG_box_qty          = 14  'Box 수량
	Const C1_EG_rcpt_qty         = 15  '납입(상차)수량
	Const C1_EG_separate_flag    = 16  '분할구분
	Const C1_EG_delivery_no      = 17  '납품번호
	Const C1_EG_issue_dt         = 18  '출하일시
	Const C1_EG_degree_cd        = 19  '차수
	Const C1_EG_buyer_cd         = 20  '납입처코드
	Const C1_EG_delete_flag      = 21  '삭제여부('N' or null:정상,Y:삭제)
	Const C1_EG_create_type      = 22  '생성구분(A:생성,B:수정,C:취소)(PK)
	Const C1_EG_send_dt          = 23  '최종전송일시
	Const C1_EG_mes_receive_flag = 24  'MES 수신여부
	Const C1_EG_mes_receive_dt   = 25  'MES 수신일시
	Const C1_EG_err_desc         = 26  'MES 수신에러내역
	Const C1_EG_if_seq           = 27  '최종전송순번
	Const C1_EG_insrt_user_id    = 28  '최초생성사용자ID
	Const C1_EG_insrt_dt         = 29  '최초생성일자
	Const C1_EG_updt_user_id     = 30  '최종수정사용자ID
	Const C1_EG_updt_dt          = 31  '최종수정일자
		
	
	Dim ObjPxi2g19Ko119										' 입력/수정용 ComProxy Dll 사용 변수 
	
	Dim lgStrPrevKey
	Dim lgLngMaxRow
	Dim lgstrData
	Dim TempArray
	Dim IG1Array
	Dim TempIG1Array
	Dim IG2ArrayNextKey
	Dim EG1Data
	
	Dim iLngRow

	lgLngMaxRow = Trim(Request("txtMaxRows"))
	
	ReDim IG1Array(C1_IG1_mes_rcv_flag)	
    'Key 값을 읽어온다	

	TempIG1Array = Split(Request("txtKeyStream"),gColSep)
	
	If IsArray(TempIG1Array) Then
		IG1Array(C1_IG1_sheet_max_rows) = C1_SHEETMAXROWS_D
		IG1Array(C1_IG1_print_fr_dt)    = Trim(TempIG1Array(C1_IT1_print_fr_dt))
		IG1Array(C1_IG1_print_to_dt)    = Trim(TempIG1Array(C1_IT1_print_to_dt))
		IG1Array(C1_IG1_item_cd)        = Trim(TempIG1Array(C1_IT1_item_cd))
		IG1Array(C1_IG1_bp_cd)          = Trim(TempIG1Array(C1_IT1_bp_cd))
		IG1Array(C1_IG1_lot_no)         = Trim(TempIG1Array(C1_IT1_lot_no))
		IG1Array(C1_IG1_del_flag)       = Trim(TempIG1Array(C1_IT1_del_flag))
		IG1Array(C1_IG1_mes_rcv_flag)   = Trim(TempIG1Array(C1_IT1_mes_rcv_flag))
    End If
	
	lgStrPrevKey = Trim(Request("lgStrPrevKey"))
	
	ReDim IG2ArrayNextKey(C1_IG2_next_create_type)

	If Trim(lgStrPrevKey) <> "" Then
		TempArray = Split(lgStrPrevKey, Chr(11))
		If IsArray(TempArray) Then
			If Ubound(TempArray) = C1_IG2_next_dock Then
				IG2ArrayNextKey(C1_IG2_next_item_cd)     = Trim(TempArray(C1_IG2_next_item_cd))
				IG2ArrayNextKey(C1_IG2_next_mat_lot_no)  = Trim(TempArray(C1_IG2_next_mat_lot_no))
				IG2ArrayNextKey(C1_IG2_next_seller_cd)   = Trim(TempArray(C1_IG2_next_seller_cd))
				IG2ArrayNextKey(C1_IG2_next_print_dt)    = UNIConvDate(TempArray(C1_IG2_next_print_dt))
				IG2ArrayNextKey(C1_IG2_next_rcpt_dt)     = UNIConvDate(TempArray(C1_IG2_next_rcpt_dt))
				IG2ArrayNextKey(C1_IG2_next_rcpt_tm)     = Trim(TempArray(C1_IG2_next_rcpt_tm))
				IG2ArrayNextKey(C1_IG2_next_create_type) = Trim(TempArray(C1_IG2_next_create_type))
			End If
		End If
	End If
	
    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If  
    
    Set ObjPxi2g19Ko119 = Server.CreateObject("PXI2g19_KO119.cFLkUpLotSvr_KO119")
    
	If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
     
    Call ObjPxi2g19Ko119.F_LOOKUP_MASTER_LOT_SVR("Q1", gStrGlobalCollection, IG1Array, IG2ArrayNextKey, EG1Data)
	
	If CheckSYSTEMError(Err, True) = True Then					
       Set ObjPxi2g19Ko119 = Nothing
       Exit Sub
    End If    
    
    Set ObjPxi2g19Ko119 = nothing    
	
	If IsEmpty(EG1Data) = False Then
		lgstrData = ""
		lgStrPrevKey = ""
		For iLngRow = 0 To UBound(EG1Data, 2)
			If  iLngRow < C1_SHEETMAXROWS_D Then
				lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1Data(C1_EG_item_cd, iLngRow))'품목코드(PK)
				lgstrData = lgstrData & Chr(11) & ""'품목PopUp Button
				lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1Data(C1_EG_item_nm, iLngRow))'품목명
				lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1Data(C1_EG_mat_lot_no, iLngRow))'LOT번호(PK)
				lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1Data(C1_EG_seller_cd, iLngRow))'납품처코드(PK)
				lgstrData = lgstrData & Chr(11) & ""'납품처PopUp Button
				lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1Data(C1_EG_seller_nm, iLngRow))'납품처명
				lgstrData = lgstrData & Chr(11) & ""'생산일자(Hidden)
				lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(EG1Data(C1_EG_print_dt, iLngRow))'발행일자(PK)
				lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(EG1Data(C1_EG_rcpt_dt, iLngRow))'납입일자(PK)
				lgstrData = lgstrData & Chr(11) & ConvToSSSSetTime(EG1Data(C1_EG_rcpt_tm, iLngRow))'납입시간(PK)
				lgstrData = lgstrData & Chr(11) & UNINumClientFormat(EG1Data(C1_EG_rcpt_qty, iLngRow), ggQty.DecPoint, 0)'납입수량
				lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1Data(C1_EG_bp_issue_no, iLngRow))'납품처발행번호
				lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1Data(C1_EG_issue_flag, iLngRow))'발행구분
				lgstrData = lgstrData & Chr(11) & ""'발행구분명
				lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1Data(C1_EG_plant_flag, iLngRow))'공장구분
				lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1Data(C1_EG_plant_cd, iLngRow))'공장코드
				lgstrData = lgstrData & Chr(11) & ""'공장PopUp Button
				lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1Data(C1_EG_gate_cd, iLngRow))'GATE
				lgstrData = lgstrData & Chr(11) & UNINumClientFormat(EG1Data(C1_EG_snp, iLngRow), ggQty.DecPoint, 0)'SNP
				lgstrData = lgstrData & Chr(11) & UNINumClientFormat(EG1Data(C1_EG_box_qty, iLngRow), ggQty.DecPoint, 0)'Box수량
				If Ucase(Trim(EG1Data(C1_EG_separate_flag, iLngRow))) = "Y" Then
					lgstrData = lgstrData & Chr(11) & "1"'분할구분
				Else
					lgstrData = lgstrData & Chr(11) & "0"'분할구분
				End If
				lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1Data(C1_EG_delivery_no, iLngRow))'납품번호
				lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(Mid(EG1Data(C1_EG_issue_dt, iLngRow), 1, 10))'출하일자
				lgstrData = lgstrData & Chr(11) & ConvToSSSSetTime(Mid(EG1Data(C1_EG_issue_dt, iLngRow), 12))'출하시간
				If Ucase(Trim(EG1Data(C1_EG_delete_flag, iLngRow))) = "Y" Then
					lgstrData = lgstrData & Chr(11) & "1"'삭제여부
				Else
					lgstrData = lgstrData & Chr(11) & "0"'삭제여부
				End If

				lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1Data(C1_EG_if_seq, iLngRow))	'최종전송순번

				lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1Data(C1_EG_mes_receive_flag, iLngRow))'MES수신여부

				If Trim(UNIDateClientFormat(Mid(EG1Data(C1_EG_mes_receive_dt, iLngRow), 1, 10))) = "" And _
				   Trim(ConvToSSSSetTime(Mid(EG1Data(C1_EG_mes_receive_dt, iLngRow), 12))) = "00:00" Then
					lgstrData = lgstrData & Chr(11) & ""		'MES수신일시
				Else
					lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(Mid(EG1Data(C1_EG_mes_receive_dt, iLngRow), 1, 10)) & " " & ConvToSSSSetTime(Mid(EG1Data(C1_EG_mes_receive_dt, iLngRow), 12))	'MES수신일시
				End If
				If Trim(UNIDateClientFormat(Mid(EG1Data(C1_EG_send_dt, iLngRow), 1, 10))) = "" And _
				   Trim(ConvToSSSSetTime(Mid(EG1Data(C1_EG_send_dt, iLngRow), 12))) = "00:00" Then
					lgstrData = lgstrData & Chr(11) & ""		'ERP최종전송일시
				Else
					lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(Mid(EG1Data(C1_EG_send_dt, iLngRow), 1, 10)) & " " & ConvToSSSSetTime(Mid(EG1Data(C1_EG_send_dt, iLngRow), 12))	'ERP최종전송일시
				End If

				lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1Data(C1_EG_err_desc, iLngRow))'에러내역
				lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1Data(C1_EG_create_type, iLngRow))'생성구분(PK)
				lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iLngRow + 1
				lgstrData = lgstrData & Chr(11) & Chr(12)
			Else
				lgStrPrevKey = ConvSPChars(EG1Data(C1_EG_buyer_company, iLngRow))
				lgStrPrevKey = lgStrPrevKey & Chr(11) & ConvSPChars(EG1Data(C1_EG_item_cd, iLngRow))
				lgStrPrevKey = lgStrPrevKey & Chr(11) & ConvSPChars(EG1Data(C1_EG_mat_lot_no, iLngRow))
				lgStrPrevKey = lgStrPrevKey & Chr(11) & ConvSPChars(EG1Data(C1_EG_seller_cd, iLngRow))
				lgStrPrevKey = lgStrPrevKey & Chr(11) & UNIDateClientFormat(EG1Data(C1_EG_print_dt, iLngRow))
				lgStrPrevKey = lgStrPrevKey & Chr(11) & UNIDateClientFormat(EG1Data(C1_EG_rcpt_dt, iLngRow))
				lgStrPrevKey = lgStrPrevKey & Chr(11) & ConvSPChars(EG1Data(C1_EG_rcpt_tm, iLngRow))
				lgStrPrevKey = lgStrPrevKey & Chr(11) & ConvSPChars(EG1Data(C1_EG_create_type, iLngRow))
			End If
       Next
    End If

	Response.Write " <Script Language=vbscript>	                                                                   " & vbCr
	Response.Write " With parent                                                                                   " & vbCr
    Response.Write "	.ggoSpread.Source        = .frm1.vspdData                                                  " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData      """ & lgstrData                                        & """    " & vbCr
    If IsArray(IG1Array) Then
		Response.Write "	.frm1.hPrintFrDt.value  = """ & UNIDateClientFormat(IG1Array(C1_IG1_print_fr_dt)) & """    " & vbcr
		Response.Write "	.frm1.hPrintToDt.value  = """ & UNIDateClientFormat(IG1Array(C1_IG1_print_to_dt)) & """    " & vbcr
		Response.Write "	.frm1.hItemCd.value     = """ & Trim(IG1Array(C1_IG1_item_cd))                    & """    " & vbCr
		Response.Write "	.frm1.hBpCd.value       = """ & Trim(IG1Array(C1_IG1_bp_cd))                      & """    " & vbCr
		Response.Write "	.frm1.hLotNo.value      = """ & Trim(IG1Array(C1_IG1_lot_no))                     & """    " & vbCr
		Response.Write "	.frm1.hDelFlag.value    = """ & Trim(IG1Array(C1_IG1_del_flag))                   & """    " & vbCr
		Response.Write "	.frm1.hMesRcvFlag.value = """ & Trim(IG1Array(C1_IG1_mes_rcv_flag))               & """    " & vbCr
	End If
    Response.Write "	.lgStrPrevKey            = """ & lgStrPrevKey                                     & """    " & vbCr
    Response.Write "	Call .DbQueryOk()                                                                          " & vbCr
    Response.Write " End With                                                                                      " & vbCr
	Response.Write " </Script>                                                                                     " & vbCr				  
    
End Sub    	 

'============================================================================================================
' Name : ConvToSSSSetTime(iVal)
' Desc : 
'============================================================================================================
Function ConvToSSSSetTime(iVal)

	Dim TempTime
	
	On Error Resume Next
	Err.Clear 
	
	If Trim(IVal) = ":" Or Trim(IVal) = "00:" Or Trim(IVal) = ":00" Then
		ConvToSSSSetTime = "00:00"
	Else
		TempTime = Split(IVal, ":")
		If IsArray(TempTime) Then
			ConvToSSSSetTime = Right("0" & TempTime(0), 2) & ":" & Right("0" & TempTime(1), 2)
		Else
			ConvToSSSSetTime = "00:00"
		End If
	End If
	
End Function
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

	On Error Resume Next                                                            '☜: Protect system from crashing
    Err.Clear 
    
	Dim ObjPxi2g19Ko119
    Dim iErrorPosition
    Dim iUpdtUserId
    Dim itxtSpread
    '-------------------
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount
    Dim Indx

    Dim iCUCount
    Dim iDCount
             
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
    iDCount  = Request.Form("txtDSpread").Count

    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount + iDCount)
             
    For Indx = 1 To iDCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(Indx)
    Next
    For Indx = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(Indx)
    Next
    itxtSpread = Join(itxtSpreadArr,"")
    '---------------------         
             
	Set ObjPxi2g19Ko119 = Server.CreateObject("PXI2g19_KO119.cMMasterLot")    

	If CheckSYSTEMError(Err,True) = true then 		
		Set ObjPj2g160Ko119 = Nothing
		Exit Sub
	End If
	
	iUpdtUserId = gUsrID
	
	Call ObjPxi2g19Ko119.M_MANAGER_MASTER_LOT_SVR(gStrGlobalCollection, _
	                                               iUpdtUserId, _
	                                               itxtSpread, _
	                                               iErrorPosition)

    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
		Set ObjPxi2g19Ko119 = Nothing
		Response.Write "<Script language=vbScript>            " & vbCr  
		Response.Write " Call Parent.RemovedivTextArea() " & vbCr							'☜: 화면 처리 ASP 를 지칭함 
		Response.Write "</Script>                        "
		Exit Sub
	End If

    Set ObjPj2g160Ko119 = Nothing    

	Response.Write "<Script language=vbScript>                                    " & vbCr  
	Response.Write " With parent                                                  " & vbCr
	If Trim(Request("hPrintFrDt")) <> "" Then Response.Write "	.frm1.txtPrintFrDt.Text    = """ & Trim(Request("hPrintFrDt"))       & """    " & vbcr
    If Trim(Request("hPrintToDt")) <> "" Then Response.Write "	.frm1.txtPrintToDt.Text    = """ & Trim(Request("hPrintToDt"))       & """    " & vbcr
    If Trim(Request("hItemCd"))    <> "" Then Response.Write "	.frm1.txtItemCd.value      = """ & Trim(Request("hItemCd"))          & """    " & vbcr
    If Trim(Request("hBpCd"))      <> "" Then Response.Write "	.frm1.txtBpCd.value        = """ & Trim(Request("hBpCd"))            & """    " & vbcr
    If Trim(Request("hLotNo"))     <> "" Then Response.Write "	.frm1.txtLotNo.value       = """ & Trim(Request("hLotNo"))            & """    " & vbcr
    If Trim(Request("hDelFlag"))   <> "" Then 
		Response.Write "	If """ & Trim(Request("hDelFlag")) & """ = """" Then      " & vbcr
		Response.Write "		.frm1.rdoDelFlagAll.checked = True                    " & vbcr
		Response.Write "	ElseIf Ucase(""" & Trim(Request("hDelFlag")) & """) = ""N"" Then " & vbcr
		Response.Write "		.frm1.rdoDelFlagNomal.checked = True                  " & vbcr
		Response.Write "	Else                                                      " & vbcr
		Response.Write "		.frm1.rdoDelFlagDel.checked = True                    " & vbcr
		Response.Write "	End If                                                    " & vbcr
	End If
	If Trim(Request("hMesRcvFlag")) <> "" Then 
		Response.Write "	If Trim(""" & Trim(Request("hMesRcvFlag")) & """ = """" Then " & vbcr
		Response.Write "		.frm1.rdoMesRcvFlagAll.checked = True                 " & vbcr
		Response.Write "	ElseIf Ucase(""" & Trim(Request("hMesRcvFlag")) & """) = ""Y"" Then " & vbcr
		Response.Write "		.frm1.rdoMesRcvFlagNomal.checked = True               " & vbcr
		Response.Write "	Else                                                      " & vbcr
		Response.Write "		.frm1.rdoMesRcvFlagFail.checked = True                " & vbcr
		Response.Write "	End If                                                    " & vbcr
	End If
    Response.Write "    Call .DBSaveOk()                                          " & vbCr   
    Response.Write " End With                                                     " & vbCr
    Response.Write "</Script> "
End Sub
%>
