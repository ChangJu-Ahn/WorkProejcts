<%'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : Prereceipt
'*  3. Program ID           : f7105mb3
'*  4. Program Name         : 선수금 기초등록의 정보삭제 
'*  5. Program Desc         : 선수금 기초등록의 정보삭제 
'*  6. Modified date(First) : 2000/09/25
'*  7. Modified date(Last)  : 2002/06/19
'*  8. Modifier (First)     : 조익성 
'*  9. Modifier (Last)      : 정용균 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

Response.Expires = -1														'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True														'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.


'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<%																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next														'☜: 
Err.Clear 

Call LoadBasisGlobalInf()
Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim iPAFG705 																'☆ : 조회용 ComPlus Dll 사용 변수 
Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim iCommandSent
Dim iTransType 
Dim iarrPrrcpt 

Const A835_I5_prrcpt_no = 0
Const A835_I5_prrcpt_dt = 1
Const A835_I5_ref_no = 2
Const A835_I5_doc_cur = 3
Const A835_I5_xch_rate = 4
Const A835_I5_prrcpt_amt = 5
Const A835_I5_loc_prrcpt_amt = 6
Const A835_I5_prrcpt_sts = 7
Const A835_I5_conf_fg = 8
Const A835_I5_prrcpt_fg = 9
Const A835_I5_prrcpt_desc = 10
Const A835_I5_prrcpt_type = 11
Const A835_I5_vat_type = 12
Const A835_I5_vat_amt = 13
Const A835_I5_vat_loc_amt = 14
Const A835_I5_issued_dt = 15
Const A835_I5_c_limit_fg = 16

	strMode = Request("txtMode")											'☜ : 현재 상태를 받음 

	Set iPAFG705 = Server.CreateObject("PAFG705.cFMngPrSvr")	    

	If CheckSYSTEMError(Err, True) = True Then					
	   Response.End 
	End If    

	'-----------------------
	'Data manipulate  area(import view match)
	'-----------------------
	Redim iarrPrrcpt(A835_I5_c_limit_fg)

	iCommandSent = "DELETE"

	iTransType	                  = "FR003"
	iArrPrrcpt(A835_I5_prrcpt_no) = Trim(Request("txtPrrcptNo"))
	iArrPrrcpt(A835_I5_prrcpt_fg) = "CT"

	Call iPAFG705.F_MANAGE_PRRCPT_SVR(gStrGloBalCollection,iCommandSent,iTransType,,,,iarrPrrcpt)

	If CheckSYSTEMError(Err, True) = True Then					
	    Set iPAFG705 = Nothing
	    Response.End 
	End If   

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write " Call parent.DbDeleteOk() " & vbCr
	Response.Write "</Script> "                 & vbCr

	Set iPAFG705 = Nothing                                                   '☜: Unload Complus
%>

