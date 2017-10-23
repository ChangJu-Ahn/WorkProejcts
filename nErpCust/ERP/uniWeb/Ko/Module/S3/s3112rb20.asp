<%@ LANGUAGE=VBSCript%>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업																		*
'*  2. Function Name        : 																			*
'*  3. Program ID           : S3112RB20																	*
'*  4. Program Name         : BOM참조(수주내역등록)													*
'*  5. Program Desc         :																			*
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2005/01/20																*
'*  8. Modified date(Last)  : 																			*
'*  9. Modifier (First)     : HJO
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 																			*
'********************************************************************************************************
%>
<% Option Explicit %>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
Call LoadBasisGlobalInf()

Call LoadInfTB19029B("Q", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("Q","*","NOCOOKIE","MB")

On Error Resume Next                                                             
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd                                                               '☜: Hide Processing message
'---------------------------------------Common-----------------------------------------------------------
CONST C_QryCnt	= 1
CONST C_QryPL		= 2
CONST C_QryBOM	= 3
'조회조건의 데이타 존재 여부 체크 
CONST C_QryPlant= 4
CONST C_QryItem= 5
CONST C_QryPoNo= 6
CONST C_QryMItem =7
CONST C_QryDel =8

	
lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

Dim strSpId
Dim intRetCD
Dim strFlag

Dim strPlant, strItem, strSoDt, strGubun, strUnit, strSoldToParty, strCur
Dim strPoNo, strPoNoSeq
Dim dblAmt
	
strPlant		= FilterVar(Request("txtPlant"), "''", "S")	
strItem		= FilterVar(Request("txtItem"), "''", "S")
strSoDt		= FilterVar( UniConvDate(Request("txtSoDt")),"''","S")
strGubun	= FilterVar(Request("txtGubun"), "''", "S")		
strUnit		= FilterVar(Request("txtUnit")	, "''", "S")		
dblAmt		= UniConvNum(Request("txtAMT"),0)	
strSoldToParty =FilterVar(request("txtSoldToParty"),"''","S")
strCur		= FilterVar(request("txtHCurrency"),"''","S")
strPoNo		= FilterVar(request("txtPoNo"),"''","S")
strPoNoSeq= FilterVar(request("txtPoNoSeq"),"''","S")

strFlag=Trim(request("txtMFlg"))


Select Case lgOpModeCRUD	
    Case CStr(UID_M0001)                                                         '☜: Query
		If strFlag= "ITEM"  then
			Call SubBizLookUp()													'☜: Look Up Item by Plant, PoNo and PoNoSeq
		Else
			Call SubBizQuery() 
        End If
    Case CStr(UID_M0002)                                                         '☜: Save,Update
        'Call SubBizSave()
    Case CStr(UID_M0003)                                                         '☜: Delete
        'Call SubBizDelete()

End Select
' ============================================================================================================
' Name : SubBizLookUp
' Desc : Query Data from Db
'============================================================================================================
 Sub SubBizLookUp()
	On Error Resume Next
	
	Dim iStrSvrData, iStrCurrency
	Dim iLngRow
	
    Call SubOpenDB(lgObjConn)

	' 공장코드의 존재유무 Check
	If Request("txtPlant") <> "" Then
	    Call SubMakeSQLStatements(C_QryPlant)
	    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then                 'If data not exists
		    Response.Write "<Script Language=vbscript>" & vbCr
			Response.Write "With parent.frm1"           & vbCr
			Response.Write ".txtPlant.value = """ & ConvSPChars(lgObjRs("PLANT_CD")) & """" & vbCr		
			Response.Write ".txtPlantNm.value = """ & ConvSPChars(lgObjRs("PLANT_NM")) & """" & vbCr
			Response.Write "End With "           & vbCr
		    Response.Write "</Script>" & vbCr
		Else
			If ChkNotFound Then
				Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
				Response.Write "<Script Language=vbscript>" & vbCr
				Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCr
				Response.End
				Response.Write "</Script>" & vbCr
			End If			
		    Exit Sub
		End If
	End If
			' 발주번호의 존재유무 Check
	If Request("txtPoNo") <> "" and Request("txtPoNoSeq") <>"" Then
	    Call SubMakeSQLStatements(C_QryPoNo)
	    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then                 'If data not exists
		    Response.Write "<Script Language=vbscript>" & vbCr
			Response.Write "With parent.frm1"           & vbCr
			Response.Write ".txtPoNo.value = """ & ConvSPChars(lgObjRs(0)) & """" & vbCr		
			Response.Write ".txtPoNoSeq.text= """ & ConvSPChars(lgObjRs(1)) & """" & vbCr
			Response.Write "End With "           & vbCr
		    Response.Write "</Script>" & vbCr
		Else
			If ChkNotFound Then
				Call DisplayMsgBox("970000", vbInformation, "발주번호와 발주순번", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
				Response.Write "<Script Language=vbscript>" & vbCr
				'Response.Write "parent.frm1.txtPoNo.value = """"" & vbCr
				Response.Write "parent.frm1.txtPoNoSeq.text = """"" & vbCr
				Response.Write "parent.frm1.txtItem.value = """"" & vbCr
				Response.Write "parent.frm1.txtPoNo.focus"		
				Response.End		
				Response.Write "</Script>" & vbCr
			End If			
		    Exit Sub
		End If
		
		Call SubMakeSQLStatements(C_QryItem) 
		If Not FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then                 'If data not exists	
			If ChkNotFound Then
					Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   	 	
					Response.Write " Parent.frm1.txtItem.value =""""" 	& vbCr											' 품목 
					Response.Write " Parent.frm1.txtItemNm.value ="""""  & vbCr											' 품목명 
					Response.Write " Parent.frm1.txtItemSpec.value = """"" & vbCr											' 규격 
					Response.Write " Parent.frm1.txtUnit.value =""""" & vbCr													' 기준단위 
					'Response.Write " Parent.frm1.txtAmt.value =" & UNINumClientFormat(0, ggQty.DecPoint, 0) & vbCr			' 필요량						
					Response.Write "</SCRIPT> "		
				'해당되는 자료가 없습니다.
				Call DisplayMsgBox("970000", vbInformation, "발주번호와 순번에 따른 품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
				Exit Sub
			End If	    
		End If
	
		If trim(Request("txtItem"))="" then 				
			Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   	 	
			Response.Write " Parent.frm1.txtItem.value =""" & ConvSPChars(lgObjRs(0))	& """" & vbCr											' 품목 
			Response.Write " Parent.frm1.txtItemNm.value =""" & ConvSPChars(lgObjRs(1))	& """"	& vbCr										' 품목명 
			Response.Write " Parent.frm1.txtItemSpec.value =""" & ConvSPChars(lgObjRs(2))	& """"	& vbCr									' 규격 
			Response.Write " Parent.frm1.txtUnit.value =""" & ConvSPChars(lgObjRs(3))	& """"	& vbCr											' 기준단위 
			Response.Write " Parent.frm1.txtAmt.value =""" & UNINumClientFormat(lgObjRs(4), ggQty.DecPoint, 0)	& """" & vbCr		' 필요량 
			Response.Write " Parent.FncLookUpOk"	& vbCr		
			Response.Write "</SCRIPT> "
		End If
	End If
	
		' 품목정보의 존재유무 Check
	If Request("txtItem") <> "" Then
	    Call SubMakeSQLStatements(C_QryMItem)
	    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then                 'If data not exists
		    Response.Write "<Script Language=vbscript>" & vbCr
			Response.Write "With parent.frm1"           & vbCr
			Response.Write " Parent.frm1.txtItem.value =""" & ConvSPChars(lgObjRs(0))	& """" & vbCr											' 품목 
			Response.Write " Parent.frm1.txtItemNm.value =""" & ConvSPChars(lgObjRs(1))	& """"	& vbCr										' 품목명 
			Response.Write " Parent.frm1.txtItemSpec.value =""" & ConvSPChars(lgObjRs(2))	& """"	& vbCr									' 규격 
			Response.Write " Parent.frm1.txtUnit.value =""" & ConvSPChars(lgObjRs(3))	& """"	& vbCr											' 기준단위 
			'Response.Write " Parent.frm1.txtAmt.value =""" & UNINumClientFormat(0, ggQty.DecPoint, 0)	& """" & vbCr		' 필요량 
			Response.Write "End With "           & vbCr
		    Response.Write "</Script>" & vbCr
		Else
			If ChkNotFound Then
				Call DisplayMsgBox("970000", vbInformation, "모품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
				Response.Write "<Script Language=vbscript>" & vbCr
				'Response.Write "parent.frm1.txtItem.value = """"" & vbCr
				Response.Write "parent.frm1.txtItemNM.value = """"" & vbCr		
						
				Response.Write " Parent.frm1.txtItemNm.value =""" 	& vbCr										' 품목명 
				Response.Write " Parent.frm1.txtItemSpec.value =""" 	& """"	& vbCr									' 규격 
				Response.Write " Parent.frm1.txtUnit.value =""" 	& """"	& vbCr											' 기준단위 
				'Response.Write " Parent.frm1.txtAmt.value =""" & UNINumClientFormat(0, ggQty.DecPoint, 0)	& """" & vbCr		' 필요량		
				Response.Write "parent.frm1.txtItem.focus"							
				Response.end
				Response.Write "</Script>" & vbCr
				
			End If			
		    Exit Sub
		End If
	End If
	

	lgObjRs.Close
	lgObjConn.Close
	Set lgObjRs = Nothing
	Set lgObjConn = Nothing
 End Sub
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	On Error Resume Next
	
	Dim iStrSvrData, iStrCurrency
	Dim iLngRow
	
    Call SubOpenDB(lgObjConn)

	' 공장코드의 존재유무 Check
	If Request("txtPlant") <> "" Then
	    Call SubMakeSQLStatements(C_QryPlant)
	    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then                 'If data not exists
		    Response.Write "<Script Language=vbscript>" & vbCr
			Response.Write "With parent.frm1"           & vbCr
			Response.Write ".txtPlant.value = """ & ConvSPChars(lgObjRs("PLANT_CD")) & """" & vbCr		
			Response.Write ".txtPlantNm.value = """ & ConvSPChars(lgObjRs("PLANT_NM")) & """" & vbCr
			Response.Write "End With "           & vbCr
		    Response.Write "</Script>" & vbCr
		Else
			If ChkNotFound Then
				Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
				Response.Write "<Script Language=vbscript>" & vbCr
				Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCr
				Response.End
				Response.Write "</Script>" & vbCr
			End If			
		    Exit Sub
		End If
	End If
	
				' 발주번호의 존재유무 Check
	If Request("txtPoNo") <> "" and Request("txtPoNoSeq") <>"" Then
	    Call SubMakeSQLStatements(C_QryPoNo)
	    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then                 'If data not exists
		    Response.Write "<Script Language=vbscript>" & vbCr
			Response.Write "With parent.frm1"           & vbCr
			Response.Write ".txtPoNo.value = """ & ConvSPChars(lgObjRs(0)) & """" & vbCr		
			Response.Write ".txtPoNoSeq.text = """ & ConvSPChars(lgObjRs(1)) & """" & vbCr
			Response.Write "End With "           & vbCr
		    Response.Write "</Script>" & vbCr
		Else
			If ChkNotFound Then
				Call DisplayMsgBox("970000", vbInformation, "발주번호와 발주순번", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
				Response.Write "<Script Language=vbscript>" & vbCr
				'Response.Write "parent.frm1.txtPoNo.value = """"" & vbCr
				Response.Write "parent.frm1.txtPoNoSeq.text = """"" & vbCr				
				Response.Write "parent.frm1.txtPoNo.focus"							
				Response.end
				Response.Write "</Script>" & vbCr
				
			End If			
		    Exit Sub
		End If
	End If
	' 품목정보의 존재유무 Check
	If Request("txtItem") <> "" Then
	    Call SubMakeSQLStatements(C_QryMItem)
	    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then                 'If data not exists
		    Response.Write "<Script Language=vbscript>" & vbCr
			Response.Write "With parent.frm1"           & vbCr
			Response.Write " Parent.frm1.txtItem.value =""" & ConvSPChars(lgObjRs(0))	& """" & vbCr											' 품목 
			Response.Write " Parent.frm1.txtItemNm.value =""" & ConvSPChars(lgObjRs(1))	& """"	& vbCr										' 품목명 
			Response.Write " Parent.frm1.txtItemSpec.value =""" & ConvSPChars(lgObjRs(2))	& """"	& vbCr									' 규격 
			Response.Write " Parent.frm1.txtUnit.value =""" & ConvSPChars(lgObjRs(3))	& """"	& vbCr											' 기준단위 
			'Response.Write " Parent.frm1.txtAmt.value =""" & UNINumClientFormat(0, ggQty.DecPoint, 0)	& """" & vbCr		' 필요량 
			Response.Write "End With "           & vbCr
		    Response.Write "</Script>" & vbCr
		Else
			If ChkNotFound Then
				Call DisplayMsgBox("970000", vbInformation, "모품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
				Response.Write "<Script Language=vbscript>" & vbCr
				'Response.Write "parent.frm1.txtItem.value = """"" & vbCr
				Response.Write "parent.frm1.txtItemNM.value = """"" & vbCr		
						
				Response.Write " Parent.frm1.txtItemNm.value =""" 	& vbCr										' 품목명 
				Response.Write " Parent.frm1.txtItemSpec.value =""" 	& """"	& vbCr									' 규격 
				Response.Write " Parent.frm1.txtUnit.value =""" 	& """"	& vbCr											' 기준단위 
				'Response.Write " Parent.frm1.txtAmt.value =""" & UNINumClientFormat(0, ggQty.DecPoint, 0)	& """" & vbCr		' 필요량		
				Response.Write "parent.frm1.txtItem.focus"							
				Response.end
				Response.Write "</Script>" & vbCr
				
			End If			
		    Exit Sub
		End If
	End If
'-------------------------------------------------------------------------------------------------------------------------
	'실제 메인 쿼리 데이타 존재 여부 조회 
	Call SubMakeSQLStatements(C_QryCnt)
	
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then            
		If lgObjRs(0)<>0 Then																	'If data  exists.
		  'PL 정보를 참조 
			Call SubMakeSQLStatements(C_QryPL)
		Else																							'If data  not exists.				   
			'BOM 정보를 참조 
			Call SubBomExplode
			Call SubMakeSQLStatements(C_QryBOM)
		End If
	End If	

 
	If Not FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then                 'If data not exists	
		If ChkNotFound Then
			'해당되는 자료가 없습니다.
			Call DisplayMsgBox("800076", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found. 
			Exit Sub
		End If	    
	End If

	' 화폐단위 
'	iStrCurrency = Request("txtCurrency")
	iStrSvrData = ""
	iLngRow = 0
'
	While (Not lgObjRs.EOF)
		iLngRow = iLngRow + 1
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs(0))			' check
  		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs(1))			' 자품목 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs(2))			' 자품목명 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs(3))			' 규격 
   		iStrSvrData = iStrSvrData & gColSep & UNINumClientFormat(lgObjRs(4), ggQty.DecPoint, 0)			' 필요량 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs(5))			' 기준단위 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs(6))			' 지급구분 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs(7))			' 조달구분 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs(8))			' HS코드 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs(9))			' 품목계정 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(lgObjRs(10))		' VAT유형 
   		iStrSvrData = iStrSvrData & gColSep & lgObjRs(11)		' VAT율 
   		
   		iStrSvrData = iStrSvrData & gColSep & iLngRow 
   		iStrSvrData = iStrSvrData & gColSep & gRowSep
   		
		lgObjRs.MoveNext
	Wend
	
	'조회가끝나면 P_BOM_FOR_EXPLOSION의 테이블을 비워준다.
	If strSpId <>"" then
		Call SubMakeSQLStatements(C_QryDel)
		If 	FncOpenRs("D",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		    Call DisplayMsgBox("800407", vbInformation, "", "", I_MKSCRIPT)              '☜ : An Error Occur
		    lgErrorStatus     = "YES"
		    Exit Sub
		End If
	End If
	
	lgObjRs.Close
	lgObjConn.Close
	Set lgObjRs = Nothing
	Set lgObjConn = Nothing
	
	Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   	
    Response.Write " Parent.ggoSpread.Source = Parent.frm1.vspdData " & vbCr
    Response.Write " Parent.ggoSpread.SSShowDataByClip """ & iStrSvrData & """" & vbCr
    Response.Write " Parent.DbQueryOk" & vbCr       
	Response.Write "</SCRIPT> "
End Sub    

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
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
    Call SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

Function ChkNotFound()
	ChkNotFound = False
	
	' Error가 발행한 경우는 recordset이 닫혔있다.
	If lgObjRs.State = adStateOpen  Then
		lgObjRs.Close
		lgObjConn.Close
		Set lgObjRs = Nothing
		Set lgObjConn = Nothing
		ChkNotFound = True
	Else
		Set lgObjRs = Nothing
		Set lgObjConn = Nothing
	End If
End Function
  
'============================================================================================================
' Name : SubBomExplode
' Desc : Query Data from Db
'============================================================================================================
Sub SubBomExplode()

    Dim strMsg_cd
    Dim strMsg_text
    
    Dim strPlant, stritem, strBaseDt
    Dim strSrchType, strBaseQty, strBomNo
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Call SubCreateCommandObject(lgObjComm)
    
    strPlant = Trim(Request("txtPlant"))									' 
	strItem = Trim(Request("txtItem"))									' 	
	strBaseDt = UniConvDate(Request("txtSoDt"))
	strBaseQty = UniConvNum(1,0)
	
	'strSrchType = "2"
	strSrchType = "5"  '20051025 박정순 수정..
	strBomNo = "1"
	
    With lgObjComm
        .CommandText = "usp_BOM_explode_main"
        .CommandType = adCmdStoredProc

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@srch_type",	advarXchar,adParamInput,2, strSrchType)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_cd",	advarXchar,adParamInput,4, strPlant)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@par_item_cd",	advarXchar,adParamInput,18, strItem)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@par_bom_no",advarXchar,adParamInput,4,strBomNo)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@base_dt_s",	advarXchar,adParamInput,10,strBaseDt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@base_qty",	adInteger,adParamInput,15,1)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd",	advarXchar,adParamOutput,6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text",	advarXchar,adParamOutput,60)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@user_id",	advarXchar,adParamOutput,13)
		
        lgObjComm.Execute ,, adExecuteNoRecords
        
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value      
       	
        If  IntRetCD <> 0 then
            
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            strMsg_text = lgObjComm.Parameters("@msg_text").Value
            strSpId = FilterVar(lgObjComm.Parameters("@user_id").Value, "''", "S")
            
            If strMsg_cd <> MSG_OK_STR Then
				Call DisplayMsgBox(strMsg_cd, vbInformation, strMsg_text, "", I_MKSCRIPT)
			End If

			IntRetCD = -1
            Response.End 
        Else
			strSpId = FilterVar(lgObjComm.Parameters("@user_id").Value, "''", "S")
			IntRetCD = 1
        End If
    Else        
		strSpId = "''"
        Call SvrMsgBox(Err.Description & "THIS IS ERROR", vbinformation, i_mkscript)        
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)        
        IntRetCD = -1
    End if
    
    Call SubCloseCommandObject(lgObjComm)

    
End Sub	
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(ByVal pvIntWhere)
	Dim iStrSelectList, iStrFromList, iStrWhere, iStrOrderBy
	
	Select Case pvIntWhere
		Case C_QryCnt
			iStrSelectList	="	SELECT  COUNT(*)  "
			iStrFromList	="	FROM  M_PL_HDR K, M_PL_DTL A  "
			iStrWhere		="	WHERE K.PL_NO = A.PL_NO  "
			iStrWhere		= iStrWhere & "		AND K.USAGE_FLG = 'Y'   "
			if trim(Request("txtGubun")) <>"" then
				iStrWhere		= iStrWhere & "		AND A.SPPL_TYPE  = " & strGubun 
			end if
			iStrWhere		= iStrWhere & "		AND K.PLANT_CD = " & strPlant 
			iStrWhere		= iStrWhere & "		AND K.ITEM_CD = " & strItem  
			iStrWhere		= iStrWhere & "		AND A.PAR_ITEM_UNIT = " & strUnit 
			iStrWhere		= iStrWhere & "		AND K.BP_CD =" & strSoldToParty
			iStrWhere		= iStrWhere & "		AND K.VALID_FROM_DT  = ( SELECT MAX(Z.VALID_FROM_DT)  FROM M_PL_HDR Z "
			iStrWhere		= iStrWhere & "			WHERE K.PL_NO = Z.PL_NO  AND Z.USAGE_FLG = 'Y'  AND Z.VALID_FROM_DT <= " & strSoDt & ")"				        

			iStrOrderBy = ""
		Case C_QryPL
		
			iStrSelectList	=	"	SELECT 0, A.ITEM_CD, B.ITEM_NM, B.SPEC, (A.PAR_ITEM_QTY *  A.CHILD_ITEM_QTY) * " & dblAmt & ",   "
			iStrSelectList	=	iStrSelectList & "		A.CHILD_ITEM_UNIT,  Z.MINOR_NM,  C.MINOR_NM , B.HS_CD, D.ITEM_ACCT, B.VAT_TYPE, B.VAT_RATE  "
			iStrFromList	=	"	FROM M_PL_HDR K, M_PL_DTL A , B_ITEM B, B_MINOR C  , B_ITEM_BY_PLANT D, B_MINOR Z   "
			iStrWhere		=	"	WHERE K.PL_NO = A.PL_NO "
			iStrWhere		=	iStrWhere	& "	AND K.USAGE_FLG = 'Y'		AND  A.ITEM_CD = B.ITEM_CD    "
			iStrWhere		=	iStrWhere	& "	AND D.PROCUR_TYPE = C.MINOR_CD		AND C.MAJOR_CD = 'P1003' "
			iStrWhere		=	iStrWhere	& "	AND B.PHANTOM_FLG <> 'Y'	AND A.PLANT_CD  = D.PLANT_CD     "
			iStrWhere		=	iStrWhere	& "	AND A.ITEM_CD = D.ITEM_CD	AND A.SPPL_TYPE  = Z.MINOR_CD	"
			iStrWhere		=	iStrWhere	& "	AND Z.MAJOR_CD = 'M2201'    "
			if trim(Request("txtGubun")) <>"" then 
				iStrWhere		=	iStrWhere	& "	AND A.SPPL_TYPE  = " & strGubun 
			end if
			iStrWhere		=	iStrWhere	& "	AND K.PLANT_CD = " & strPlant
			iStrWhere		=	iStrWhere	& "	AND K.ITEM_CD    = " & strItem 
			iStrWhere		=	iStrWhere	& "	AND A.PAR_ITEM_UNIT = " & strUnit
			iStrWhere		=  iStrWhere & "		AND K.BP_CD =" & strSoldToParty
			iStrWhere		=	iStrWhere	& "	AND K.VALID_FROM_DT  = ( SELECT MAX(Z.VALID_FROM_DT)  FROM M_PL_HDR Z  "
			iStrWhere		=	iStrWhere	& "		WHERE K.PL_NO = Z.PL_NO  AND Z.USAGE_FLG = 'Y'  AND Z.VALID_FROM_DT <=  " & strSoDt & ")"		
				
			iStrOrderBy = ""
			
		Case C_QryBOM				
			iStrSelectList	=	"	SELECT 0, A.CHILD_ITEM_CD, B.ITEM_NM, B.SPEC,  SUM(A.CHILD_ITEM_QTY) * " & dblAmt & " AS SUM_CHILD_ITEM_QTY, "
			iStrSelectList	=	iStrSelectList	& "		A.CHILD_ITEM_UNIT,   Z.MINOR_NM,C.MINOR_NM ,  B.HS_CD, D.ITEM_ACCT, B.VAT_TYPE, B.VAT_RATE  "
			iStrFromList	=	"	FROM P_BOM_FOR_EXPLOSION A, B_ITEM B, B_MINOR C  , B_ITEM_BY_PLANT D, B_MINOR Z		"
			iStrWhere		=	"	WHERE A.CHILD_ITEM_CD = B.ITEM_CD    AND D.PROCUR_TYPE = C.MINOR_CD "
			iStrWhere		=	iStrWhere	& "		AND C.MAJOR_CD = 'P1003'             AND B.PHANTOM_FLG <> 'Y'  "
			iStrWhere		=	iStrWhere	& "		AND A.CHILD_ITEM_CD   = D.ITEM_CD    AND A.PLANT_CD  = D.PLANT_CD  "
			iStrWhere		=	iStrWhere	& "		AND A.SUPPLY_TYPE  = Z.MINOR_CD      AND Z.MAJOR_CD = 'M2201'  "
			if trim(Request("txtGubun")) <>"" then
				iStrWhere		=	iStrWhere	& "		AND A.SUPPLY_TYPE = " & strGubun
			end if
			iStrWhere		=	iStrWhere	& "		AND A.PLANT_CD =  " & strPlant
			iStrWhere		=	iStrWhere	& "		AND A.USER_ID  =  " & strSpId
			iStrWhere		=	iStrWhere	& "	GROUP BY A.CHILD_ITEM_CD, A.CHILD_BOM_NO, B.ITEM_NM, B.SPEC, A.CHILD_ITEM_UNIT,  C.MINOR_NM , A.PRNT_BOM_NO,   "
			iStrWhere		=	iStrWhere	& "		B.HS_CD, D.ITEM_ACCT, B.VAT_TYPE, B.VAT_RATE, Z.MINOR_NM  "
			
			iStrOrderBy = ""
		Case C_QryPlant
			iStrSelectList	= " SELECT PLANT_CD, PLANT_NM "
			iStrFromList	= " FROM dbo.B_PLANT "
			iStrWhere		= " WHERE PLANT_CD =  " & strPlant 
			iStrOrderBy	= ""
		Case C_QryItem
			iStrSelectList	= "	SELECT A.ITEM_CD, B.ITEM_NM, B.SPEC, A.PO_BASE_UNIT, A.PO_QTY "
			iStrFromList	= "	FROM M_PUR_ORD_DTL A, B_ITEM  B "
			iStrWhere		= "	WHERE A.ITEM_CD = B.ITEM_CD  "
			iStrWhere		=	iStrWhere & "		AND  A.PO_NO = " & strPoNo
			iStrWhere		=	iStrWhere & "		AND A.PO_SEQ_NO =" & strPoNoSeq
			iStrWhere		=	iStrWhere & "		AND A.PLANT_CD=" & strPlant	
			If strItem <>"''" Then
				iStrWhere		=	iStrWhere & "		AND A.ITEM_CD=" & strItem
			End If
			iStrOrderBy = ""
		Case C_QryPoNo
			iStrSelectList	= "	SELECT A.PO_NO,A.PO_SEQ_NO "
			iStrFromList =	 "	 FROM M_PUR_ORD_DTL A ,   M_PUR_ORD_HDR F"
			iStrWhere =	"	WHERE  A.PO_NO = F.PO_NO        "
			iStrWhere		= iStrWhere & "		AND A.CLS_FLG = 'N'    AND F.RELEASE_FLG = 'Y' "
			iStrWhere		= iStrWhere & "		AND F.PO_DT <= " & strSoDt		
			iStrWhere		=	iStrWhere & "		AND  A.PO_NO = " & strPoNo
			iStrWhere		=	iStrWhere & "		AND A.PO_SEQ_NO =" & strPoNoSeq
			iStrWhere		=	iStrWhere & "		AND A.PLANT_CD=" & strPlant	
			iStrOrderBy = ""
		Case C_QryMItem
			
			iStrSelectList	= "	SELECT A.ITEM_CD, A.ITEM_NM, A.SPEC,  A.BASIC_UNIT "
 			iStrFromList	= "	FROM  B_ITEM A, B_ITEM_BY_PLANT  B "
			iStrWhere		= "	WHERE A.ITEM_CD = B.ITEM_CD  			 "
			iStrWhere		=	iStrWhere & "		AND B.PLANT_CD=" & strPlant				
			iStrWhere		=	iStrWhere & "		AND A.ITEM_CD=" & strItem			
			iStrOrderBy = ""
		Case C_QryDel			
			iStrSelectList	= "	DELETE  "
 			iStrFromList	= "	P_BOM_FOR_EXPLOSION  "
			iStrWhere		= "	WHERE	 USER_ID  =  " & strSpId			
			iStrOrderBy = ""
	End Select	
	lgStrSql = iStrSelectList & iStrFromList & iStrWhere & iStrOrderBy
End Sub
%>
