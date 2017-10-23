<% 
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Monthly Exchange Rate)
'*  3. Program ID           : B1703mb1.asp
'*  4. Program Name         : B1703mb1.asp
'*  5. Program Desc         :
'*  6. Comproxy List        :
'                             +B17031CtrlMnthlyExchangeRate
'                             +B17038ListMnthlyExchangeRate
'*  7. Modified date(First) : 2000/09/05
'*  8. Modified date(Last)  : 2002/12/11
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
'**********************************************************************************************

%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

On Error Resume Next														'☜: 
Err.Clear

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strSpread

Dim StrNextKey		' 다음 값 
Dim StrNextToKey
Dim lgStrPrevKey	' 이전 값 
Dim lgStrPrevToKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount          

Call LoadBasisGlobalInf()

Call loadInfTB19029B("I", "B","NOCOOKIE","MB")

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strSpread = Trim(Request("txtSpread"))

Select Case strMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
    Dim I1_b_currency 
    Dim I2_b_monthly_exchange_rate

    Const B369_EG1_E1_currency = 0    
    Const B369_EG1_E1_currency_desc = 1
    Const B369_EG1_E2_currency = 2    
    Const B369_EG1_E2_currency_desc = 3
    Const B369_EG1_E3_apprl_yrmnth = 4
    Const B369_EG1_E3_multi_divide = 5
    Const B369_EG1_E3_std_rate = 6
    Const B369_EG1_E3_buy_rate = 7
    Const B369_EG1_E3_sell_rate = 8
    Const B369_EG1_E3_cash_buy_rate = 9
    Const B369_EG1_E3_cash_sell_rate = 10
    Const B369_EG1_E3_usd_rate = 11
    Const B369_EG1_E3_scope_average = 14
    
    
	Dim ObjPB2G071
	Dim Export_Array
	
    I1_b_currency = Request("txtCurrency")
	I2_b_monthly_exchange_rate = Request("txtValidDt")

    Set ObjPB2G071 = server.CreateObject ("PB2G071.cBListMnthlyExchRate")    
    on error resume next
    Err.Clear 
    Export_Array = ObjPB2G071.B_LIST_MNTHLY_EXCHANGE_RATE  (gStrGlobalCollection,I1_b_currency,I2_b_monthly_exchange_rate)

    Set ObjPB2G071 = nothing
    If CheckSYSTEMError(Err,True) = True Then 
		Response.End                          
    End If   
    on error goto 0
    
%>
<Script Language=vbscript>
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp
    Dim strData
    Dim tmpYYYYMM
    Dim tmpYYYYMMData



	With parent																	'☜: 화면 처리 ASP 를 지칭함 
		'.Clear
		 LngMaxRow = 0
<%      
        GroupCount = Ubound(Export_Array,1)
        
	    For LngRow = 0 To GroupCount
%>          
            tmpYYYYMMData = "<%=ConvSPChars(Export_Array(LngRow,B369_EG1_E3_apprl_yrmnth))%>"
            
		    tmpYYYYMM = .parent.gDateFormatYYYYMM

		    tmpYYYYMM = .parent.UniConvYYYYMMDDToDate(tmpYYYYMM, left(tmpYYYYMMData,4),right(tmpYYYYMMData, 2), "01")
                       
		    strData = strData & Chr(11) & tmpYYYYMM															'0
		    strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B369_EG1_E1_currency )))%>"   '1
		    strData = strData & Chr(11) & " " '2 PopupButton
		    strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B369_EG1_E2_currency )))%>" '3
		    strData = strData & Chr(11) & " " '4 PopupButton
		    strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B369_EG1_E3_multi_divide )))%>"  '5

		    strData = strData & Chr(11) & "<%=UNINumClientFormat(Trim(Export_Array(LngRow,B369_EG1_E3_std_rate )), ggExchRate.DecPoint,0)%>"      '6
		    strData = strData & Chr(11) & "<%=UNINumClientFormat(Trim(Export_Array(LngRow,B369_EG1_E3_buy_rate )), ggQty.DecPoint,0)%>"      '7
		    strData = strData & Chr(11) & "<%=UNINumClientFormat(Trim(Export_Array(LngRow,B369_EG1_E3_sell_rate )), ggQty.DecPoint,0)%>"     '8
		    strData = strData & Chr(11) & "<%=UNINumClientFormat(Trim(Export_Array(LngRow,B369_EG1_E3_cash_buy_rate )), ggQty.DecPoint,0)%>"  '9
		    strData = strData & Chr(11) & "<%=UNINumClientFormat(Trim(Export_Array(LngRow,B369_EG1_E3_cash_sell_rate )), ggQty.DecPoint,0)%>" '10
		    strData = strData & Chr(11) & "<%=UNINumClientFormat(Trim(Export_Array(LngRow,B369_EG1_E3_usd_rate )), ggQty.DecPoint,0)%>"      '11		
		    strData = strData & Chr(11) & "<%=UNINumClientFormat(Trim(Export_Array(LngRow,B369_EG1_E3_scope_average )), ggExchRate.DecPoint,0)%>"      '14		
		    strData = strData & Chr(11) & LngMaxRow + <%=LngRow%> + 1
		    strData = strData & Chr(11) & Chr(12)
            
<%      		
        Next
%>    	

		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData strData
		.frm1.hValidDt.value = "<%=Request("txtValidDt")%>"
		.frm1.hCurrency.value = "<%=ConvSPChars(Request("txtCurrency"))%>"			
			
		.DbQueryOk
	End With
	
</Script>	
<%    
    
Case CStr(UID_M0002)																'☜: 저장 요청을 받음 
             									
	If Request("txtMaxRows") = "" Then
		Call DisplayMsgBox("700117", vbInformation, "", "", I_MKSCRIPT)
		Response.End 
	End If


    Dim Obj2PB2G071
    Dim iErrorPosition

    Set Obj2PB2G071 = server.CreateObject ("PB2G071.cBCtrlMnthlyExchRate")    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear        

    Call Obj2PB2G071.B_CTRL_MNTHLY_EXCHANGE_RATE (gStrGlobalCollection,strSpread,iErrorPosition)
    Set Obj2PB2G071 = nothing

    If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
		Response.End 
    End If
    on error goto 0                                                             


	Dim arrVal, arrTemp																'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim strStatus																	'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
	Dim	lGrpCnt																		'☜: Group Count
	Dim strUsrId
	
	strUsrId = Request("txtInsrtUserId")
	arrTemp = Split(strSpread, gRowSep)									'☆: Spread Sheet 내용을 담고 있는 Element명 
%>
<Script Language=vbscript>
	With parent																		'☜: 화면 처리 ASP 를 지칭함 
		'window.status = "저장 성공"
		.DbSaveOk
	End With
	
</Script>
<%					

End Select

%>
