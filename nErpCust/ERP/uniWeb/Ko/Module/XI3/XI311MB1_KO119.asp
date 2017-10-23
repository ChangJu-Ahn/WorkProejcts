<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/IncSvrMain.asp" -->

<%
    Call LoadBasisGlobalInf()
    Call HideStatusWnd                                                               '☜: Hide Processing message

    Dim strSQL
    Dim strMES
    Dim objNumber
    Dim strParam(1)

	Dim istrMode
	Dim strRetMsg
	Dim IntRetCd

	Dim txtPlantCd
	Dim txtUserId

    Call LoadInfTB19029B( "I", "*", "NOCOOKIE", "MB")
    Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

	On Error Resume Next														
 	Err.Clear																

	txtPlantCd	= UCase(Trim(Request("txtPlant")))
	txtUserId	= UCase(Trim(Request("txtUserId")))

	istrMode = Request("txtMode")

	Select Case istrMode
			Case CStr("T")					'MES의 검사요청 자료를 ERP로 복사

					Call SubOpenDB(lgObjConn)				' 데이터 베이스 커넥션 개체 생성
					Call SubCreateCommandObject(lgObjComm)

					With lgObjComm
						.CommandText = "usp_INSPECTION_REQ_MES_RCV_KO119"
						.CommandType = adCmdStoredProc
						.CommandTimeout = 1800	

						lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",	adInteger,adParamReturnValue)
						lgObjComm.Parameters.Append lgObjComm.CreateParameter("@CON_PLANT_CD",	adVarChar, adParamInput,    4, txtPlantCd)
						lgObjComm.Parameters.Append lgObjComm.CreateParameter("@CON_USER_ID",	adVarChar, adParamInput,   13, txtUserId)	   
						lgObjComm.Parameters.Append lgObjComm.CreateParameter("@CON_ERROR_DESC",adVarChar, adParamOutput, 200)

						lgObjComm.Execute ,, adExecuteNoRecords
					End With
	
					If Err.number = 0 Then
						intRetCd = lgObjComm.Parameters("RETURN_VALUE").Value
						
						If intRetCd <> 0 Then
							strRetMsg = lgObjComm.Parameters("@CON_ERROR_DESC").Value
							If strRetMsg <> "" Then
								Call DisplayMsgBox(strRetMsg, vbInformation, "", "", I_MKSCRIPT)
							End If	
						End If
					Else
						Call DisplayMsgBox(Err.Description, vbInformation, "", "", I_MKSCRIPT)	
					End If
		
					Response.Write "<Script Language=vbscript>	"	& vbcr
					Response.Write "With parent.frm1			"	& vbcr
					Response.Write "	parent.FncQuery			"	& vbcr
					Response.Write "End With					"	& vbcr
					Response.Write "</Script>					"	& vbcr	 

					Call SubCloseCommandObject(lgObjComm)
					Call SubCloseDB(lgObjConn) 


			Case CStr("A")					'MES에서 수신한 검사요청 자료를 ERP의 검사요청의뢰 등록

					Call SubOpenDB(lgObjConn)				' 데이터 베이스 커넥션 개체 생성
					Call SubCreateCommandObject(lgObjComm)

					With lgObjComm
						.CommandText = "usp_INSPECTION_REQ_ERP_UPLOAD_KO119"
						.CommandType = adCmdStoredProc
						.CommandTimeout = 1800	

						lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",	adInteger,adParamReturnValue)
						lgObjComm.Parameters.Append lgObjComm.CreateParameter("@CON_PLANT_CD",	adVarChar, adParamInput,    4, txtPlantCd)
						lgObjComm.Parameters.Append lgObjComm.CreateParameter("@CON_USER_ID",	adVarChar, adParamInput,   13, txtUserId)	   
						lgObjComm.Parameters.Append lgObjComm.CreateParameter("@CON_ERROR_DESC",adVarChar, adParamOutput, 200)

						lgObjComm.Execute ,, adExecuteNoRecords
					End With
	
					If Err.number = 0 Then
						intRetCd = lgObjComm.Parameters("RETURN_VALUE").Value
						
						If intRetCd <> 0 Then
							strRetMsg = lgObjComm.Parameters("@CON_ERROR_DESC").Value
							If strRetMsg <> "" Then
								Call DisplayMsgBox(strRetMsg, vbInformation, "", "", I_MKSCRIPT)
							End If	
						End If
					Else
						Call DisplayMsgBox(Err.Description, vbInformation, "", "", I_MKSCRIPT)	
					End If
		
					Response.Write "<Script Language=vbscript>	"	& vbcr
					Response.Write "With parent.frm1			"	& vbcr
					Response.Write "	parent.FncQuery			"	& vbcr
					Response.Write "End With					"	& vbcr
					Response.Write "</Script>					"	& vbcr	 

					Call SubCloseCommandObject(lgObjComm)
					Call SubCloseDB(lgObjConn) 


			Case CStr("R")
			
					Dim iLngRow, lgStrPrevKey, lgArrPrevKey
					
					Const C1_SHEETMAXROWS_D  = 100
					
					Const C_plant_cd       = 0
					Const C_prodt_order_no = 1
					Const C_lot_no         = 2
					Const C_request_seq    = 3
					Const C_create_type    = 4
					
					lgLngMaxRow = Trim(Request("txtMaxRows"))
    
					If Trim(Request("lgStrPrevKey")) = "" Then
						lgStrPrevKey = ""
					Else
						lgStrPrevKey = Trim(Request("lgStrPrevKey"))
						lgArrPrevKey = Split(lgStrPrevKey, Chr(11))
					End If

					Call SubOpenDB(lgObjConn)				' 데이터 베이스 커넥션 개체 생성
					
					strSQL = ""
					strSQL = strSQL & " SELECT TOP " & C1_SHEETMAXROWS_D + 1
					strSQL = strSQL & "    a.prodt_order_no"
					strSQL = strSQL & "  , a.request_seq"
					strSQL = strSQL & "  , a.item_cd"
					strSQL = strSQL & "  , b.item_nm"
					strSQL = strSQL & "  , a.lot_no"
					strSQL = strSQL & "  , a.lot_size"
					strSQL = strSQL & "  , a.create_type"
					strSQL = strSQL & "  , a.insp_req_no"
					strSQL = strSQL & "  , a.send_dt"
					strSQL = strSQL & "  , a.erp_apply_flag"
					strSQL = strSQL & "  , a.err_desc"
					strSQL = strSQL & "  , a.erp_receive_dt"
					strSQL = strSQL & "  , a.plant_cd"
					strSQL = strSQL & " FROM"
					strSQL = strSQL & "  t_if_rcv_insp_req_ko119 a (nolock)"
					strSQL = strSQL & "  LEFT OUTER JOIN b_item b  (nolock) ON b.item_cd = a.item_cd"
					strSQL = strSQL & " WHERE"
					strSQL = strSQL & "      a.plant_cd = '" & Request("txtPlant") & "'"
					strSQL = strSQL & "  AND a.send_dt BETWEEN (CASE '" & Request("txtConSoFrDt") & "' WHEN '' THEN '1900-01-01' ELSE '" & Request("txtConSoFrDt") & "'          END)"
					strSQL = strSQL & "                AND     (CASE '" & Request("txtConSoToDt") & "' WHEN '' THEN '9999-12-31' ELSE '" & Request("txtConSoToDt") & " 23:59:59' END)"
					strSQL = strSQL & "  AND a.item_cd LIKE    (CASE '" & Request("txtItemCD")    & "' WHEN '' THEN '%'          ELSE '" & Request("txtItemCD")    & "'          END)"
					strSQL = strSQL & "  AND a.prodt_order_no LIKE (CASE '" & Request("txtMakOrdNo")  & "' WHEN '' THEN '%' ELSE '" & Request("txtMakOrdNo") & "'  END)"
					strSQL = strSQL & "  AND a.insp_req_no    LIKE (CASE '" & Request("txtInspReqNo") & "' WHEN '' THEN '%' ELSE '" & Request("txtInspReqNo") & "' END)"
					strSQL = strSQL & "  AND a.erp_apply_flag LIKE (CASE '" & Request("txtCfmFlag")   & "' WHEN '' THEN '%' ELSE '" & Request("txtCfmFlag") & "'   END)"
					If Trim(lgStrPrevKey) <> "" Then
						strSQL = strSQL & "  AND (a.plant_cd > '" & lgArrPrevKey(C_plant_cd) & "'"
						strSQL = strSQL & "  OR   (a.plant_cd = '" & lgArrPrevKey(C_plant_cd) & "'"
						strSQL = strSQL & "  AND   (a.prodt_order_no > '" & lgArrPrevKey(C_prodt_order_no) & "'"
						strSQL = strSQL & "  OR     (a.prodt_order_no = '" & lgArrPrevKey(C_prodt_order_no) & "'"
						strSQL = strSQL & "  AND     (a.lot_no > '" & lgArrPrevKey(C_lot_no) & "'"
						strSQL = strSQL & "  OR       (a.lot_no = '" & lgArrPrevKey(C_lot_no) & "'"
						strSQL = strSQL & "  AND       (a.request_seq > " & lgArrPrevKey(C_request_seq)
						strSQL = strSQL & "  OR         (a.request_seq = " & lgArrPrevKey(C_request_seq)
						strSQL = strSQL & "  AND         a.create_type >= '" & lgArrPrevKey(C_create_type) & "'))))))))"
					End If
					strSQL = strSQL & " ORDER BY"
					strSQL = strSQL & "    a.plant_cd"
					strSQL = strSQL & "  , a.prodt_order_no"
					strSQL = strSQL & "  , a.lot_no"
					strSQL = strSQL & "  , a.request_seq"
					strSQL = strSQL & "  , a.create_type"

					If FncOpenRs("R", lgObjConn, lgObjRs, strSQL, "X", "X") = False Then
						Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
					Else

						Response.Write "<Script Language=""VBScript"">"         & vbCrLf
						Response.Write "With Parent"                            & vbCrLf
						Response.Write "    .ggoSpread.Source = .frm1.vspdData" & vbCrLf
						Response.Write "    .ggoSpread.SSShowDataByClip """
						
						iLngRow = 0
						lgStrPrevKey = ""

						Do While Not lgObjRs.EOF
							If iLngRow < C1_SHEETMAXROWS_D Then

								If Ucase(lgObjRs("erp_apply_flag")) = "Y" Then
									Response.Write Chr(11) & "1"
								Else
									Response.Write Chr(11) & "0"
								End If

								Response.Write Chr(11) & lgObjRs("prodt_order_no")
								Response.Write Chr(11) & lgObjRs("request_seq")
								Response.Write Chr(11) & lgObjRs("item_cd")
								Response.Write Chr(11) & Replace(lgObjRs("item_nm"), """", """""")
								Response.Write Chr(11) & lgObjRs("lot_no")
								Response.Write Chr(11) & UniNumClientFormat(lgObjRs("lot_size"), ggQty.DecPoint, 0)
								Response.Write Chr(11) & lgObjRs("create_type")
								Response.Write Chr(11) & lgObjRs("insp_req_no")
								Response.Write Chr(11) & lgObjRs("send_dt")
								Response.Write Chr(11) & lgObjRs("erp_receive_dt")
								Response.Write Chr(11) & lgObjRs("erp_apply_flag")
								Response.Write Chr(11) & lgObjRs("err_desc")

								Response.Write Chr(11) & Chr(12)
							Else
								lgStrPrevKey = lgStrPrevKey & Trim(lgObjRs("plant_cd"))       & Chr(11)
								lgStrPrevKey = lgStrPrevKey & Trim(lgObjRs("prodt_order_no")) & Chr(11)
								lgStrPrevKey = lgStrPrevKey & Trim(lgObjRs("lot_no"))         & Chr(11)
								lgStrPrevKey = lgStrPrevKey & Trim(lgObjRs("request_seq"))    & Chr(11)
								lgStrPrevKey = lgStrPrevKey & Trim(lgObjRs("create_type"))    & Chr(11)
							End If
			
							iLngRow = iLngRow + 1
							lgObjRs.MoveNext
						Loop
						Response.Write """"                                               & vbCrLf
						Response.Write "	.lgStrPrevKey = """ & lgStrPrevKey & """    " & vbCrLf
						Response.Write "    .DBQueryOk  "                                 & vbCrLf
						Response.Write "End with"                                         & vbCrLf
						Response.Write "</Script>"                                        & vbCrLf
					End If

					Call SubCloseRs(lgObjRs)
					Call SubCloseDB(lgObjConn)
	End Select

	Response.End

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub
%>