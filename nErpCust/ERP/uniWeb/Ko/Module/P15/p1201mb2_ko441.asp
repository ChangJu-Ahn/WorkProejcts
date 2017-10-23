<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1201mb2_ko441.asp
'*  4. Program Name         : Component Allocation (Query)
'*  5. Program Desc         :
'*  6. Component List       : PP1C505.cPListCmpReqByRtng
'*  7. Modified date(First) : 2000/03/28
'*  8. Modified date(Last)  : 2008/01/31
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : HAN cheol
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf

Dim pPP1C505																'☆ : 조회용 Component Dll 사용 변수 

Dim I1_plant_cd, I2_item_cd, I3_bom_no, I4_rout_no, I5_opr_no, I6_valid_to_dt, I7_next_seq
Dim E1_next_seq, EG1_list

' EG1_list
Const P141_EG1_E1_select_char = 0  ' View Name : export_item ief_supplied
Const P141_EG1_E2_issued_sl_cd = 3
Const P141_EG1_E2_issued_unit = 4
Const P141_EG1_E3_sl_nm = 7
Const P141_EG1_E4_item_cd = 8      ' View Name : export_item_for_prnt b_item
Const P141_EG1_E4_item_nm = 9
Const P141_EG1_E4_spec = 10
Const P141_EG1_E5_item_nm = 12
Const P141_EG1_E5_spec = 13
Const P141_EG1_E6_seq = 15
Const P141_EG1_E6_child_item_seq = 17
Const P141_EG1_E6_child_item_cd = 18
Const P141_EG1_E6_valid_from_dt = 19
Const P141_EG1_E6_valid_to_dt = 20

Dim iLngNextKey	' 다음 값 
Dim iLngPrevKey	' 이전 값 
Dim iLngMaxRow		' 현재 그리드의 최대Row
Dim iLngRow
Dim iLngGrpCnt          
Dim strData
Dim TmpBuffer
Dim iTotalStr

Const SheetMaxRowsD = 50	'Server PAD에서 정의한 Group View Size PAD에서 BOM을 읽을 때 만약에 자품목이 3개 있으면 실제 
						    '자품목 할당할 수 있는 품목이 2개라도 Next 값은 3번째를 넘겨 주므로 아래 비교 문장이 맞지 않아 
						    ' Group View Size보다 작으면 다음 Query를 안타도록 함. 		

If Request("txtPlantCd") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)             
	Response.End 
ElseIf Request("txtItemCd") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)            
	Response.End 
ElseIf Request("txtRoutNo") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)           
	Response.End 
ElseIf Request("txtOprNo") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)           
	Response.End 
End If
	
iLngPrevKey = Request("lgIntPrevKey")

'-----------------------
'Data manipulate  area(import view match)
'-----------------------
I1_plant_cd		= UCase(Trim(Request("txtPlantCd")))
I2_item_cd		= UCase(Trim(Request("txtItemCd")))
I3_bom_no 		= UCase(Trim(Request("txtBomNo")))
I4_rout_no		= UCase(Trim(Request("txtRoutNo")))
I5_opr_no		= UCase(Trim(Request("txtOprNo")))
I6_valid_to_dt	= UniConvDate(Request("txtBaseDt"))
I7_next_seq 	= CLng(iLngPrevKey)

'-----------------------
'Com action area
'-----------------------
Set pPP1C505 = Server.CreateObject("PP1C505.cPListCmpReqByRtng")    

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPP1C505.P_LIST_COMP_REQ_BY_ROUT(gStrGlobalCollection, SheetMaxRowsD+1, I1_plant_cd, I2_item_cd, I3_bom_no, I4_rout_no, _
                     I5_opr_no, I6_valid_to_dt, I7_next_seq, E1_next_seq, EG1_list)

If CheckSYSTEMError(Err, True) = True Then
	Set pPP1C505 = Nothing															'☜: Unload Component
	Response.End
End If

Set pPP1C505 = Nothing															'☜: Unload Component

iLngMaxRow = Request("txtMaxRows")											'Save previous Maxrow                                                

If Not IsNull(EG1_list) Then
	iLngGrpCnt = UBound(EG1_list, 1)

	If (EG1_list(iLngGrpCnt, P141_EG1_E6_seq) = E1_next_seq) Or (iLngGrpCnt < SheetMaxRowsD) Then
		iLngNextKey = 0
	Else
		IntNextKey = E1_next_seq
	End If
	
	ReDim TmpBuffer(iLngGrpCnt)
	    
	For iLngRow = 0 To iLngGrpCnt
		strData = ""
		strData = strData & Chr(11) & ""															'1:Check Box
		strData = strData & Chr(11) & ConvSPChars(EG1_list(iLngRow, P141_EG1_E6_child_item_cd))					'2:Child Item Code
		strData = strData & Chr(11) & ConvSPChars(EG1_list(iLngRow, P141_EG1_E5_item_nm))			'3:Child Item Name
		strData = strData & Chr(11) & ConvSPChars(EG1_list(iLngRow, P141_EG1_E5_spec))				'5:Child Item Spec
		strData = strData & Chr(11) & ConvSPChars(EG1_list(iLngRow, P141_EG1_E2_issued_sl_cd))		'6:Issued Sl Code
		strData = strData & Chr(11) & ConvSPChars(EG1_list(iLngRow, P141_EG1_E3_sl_nm))				'7:Issued Sl Name
		strData = strData & Chr(11) & ConvSPChars(EG1_list(iLngRow, P141_EG1_E2_issued_unit))		'8:Issued Unit
		strData = strData & Chr(11) & ConvSPChars(EG1_list(iLngRow, P141_EG1_E4_item_cd))			'9:Prnt Item Code
		strData = strData & Chr(11) & ConvSPChars(EG1_list(iLngRow, P141_EG1_E4_item_nm))			'10:Prnt Item Name	
		strData = strData & Chr(11) & ConvSPChars(EG1_list(iLngRow, P141_EG1_E4_spec))				'11:Prnt Item Spec
		strData = strData & Chr(11) & EG1_list(iLngRow, P141_EG1_E6_child_item_seq)					'12:Child Item Seq
		strData = strData & Chr(11) & UNIDateClientFormat(EG1_list(iLngRow, P141_EG1_E6_valid_from_dt))
		strData = strData & Chr(11) & UNIDateClientFormat(EG1_list(iLngRow, P141_EG1_E6_valid_to_dt))
		        
		strData = strData & Chr(11) & EG1_list(iLngRow, P141_EG1_E1_select_char)					'10:Hidden
		strData = strData & Chr(11) & iLngMaxRow + iLngRow
		strData = strData & Chr(11) & Chr(12)
		
		TmpBuffer(iLngRow) = strData
	Next
End If

iTotalStr = Join(TmpBuffer, "")
Response.Write "<Script Language=VBScript>" & vbCrLf
Response.Write "With parent" & vbCrLf										'☜: 화면 처리 ASP 를 지칭함 

If IsEmpty(EG1_list) = False Then
	Response.Write ".ggoSpread.Source = .frm1.vspdData2" & vbCrLf
	Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """" & vbCrLf
End If
		
Response.Write ".lgIntPrevKey = """ & ConvSPChars(IntNextKey) & """" & vbCrLf
Response.Write ".frm1.hBaseDt.Value = """ & Request("txtBaseDt") & """" & vbCrLf

' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 
Response.Write "If .frm1.vspdData2.MaxRows < .VisibleRowCnt(.frm1.vspdData2, 0) And .lgIntPrevKey <> 0 Then" & vbCrLf
	Response.Write ".DbDtlQuery" & vbCrLf
Response.Write "Else" & vbCrLf
	Response.Write ".DbDtlQueryOk(" & iLngMaxRow + 1 & ")" & vbCrLf
Response.Write "End If" & vbCrLf

Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf

Response.End
%>
