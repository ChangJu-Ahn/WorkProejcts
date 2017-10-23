<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Daily Exchange Rate)
'*  3. Program ID           : B1702mb1.asp
'*  4. Program Name         : B1702mb1.asp
'*  5. Program Desc         :
'*  6. Comproxy List        :
'                             +B17021CtrlDailyExchangeRate
'                             +B17028ListDailyExchangeRate
'*  7. Modified date(First) : 2000/09/05
'*  8. Modified date(Last)  : 2002/12/11
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************


Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.

%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

On Error Resume Next														'��: 
Err.Clear

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strSpread
Dim StrNextKey		' ���� �� 
Dim StrNextToKey
Dim lgStrPrevKey	' ���� �� 
Dim lgStrPrevToKey	' ���� �� 
Dim LngMaxRow		' ���� �׸����� �ִ�Row
Dim LngRow
Dim GroupCount          

Call LoadBasisGlobalInf()

Call loadInfTB19029B("I", "B","NOCOOKIE","MB")
strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
strSpread = Trim(Request("txtSpread"))

Select Case strMode
Case CStr(UID_M0001)														'��: ���� ��ȸ/Prev/Next ��û�� ���� 
    Dim I1_b_currency
    Dim I2_b_daily_exchange_rate

    Const B366_EG1_E1_currency = 0    
    Const B366_EG1_E1_currency_desc = 1
    Const B366_EG1_E2_currency = 2    
    Const B366_EG1_E2_currency_desc = 3
    Const B366_EG1_E3_apprl_dt = 4    
    Const B366_EG1_E3_multi_divide = 5
    Const B366_EG1_E3_std_rate = 6
    Const B366_EG1_E3_buy_rate = 7
    Const B366_EG1_E3_sell_rate = 8
    Const B366_EG1_E3_cash_buy_rate = 9
    Const B366_EG1_E3_cash_sell_rate = 10
    Const B366_EG1_E3_usd_rate = 11
    Const B366_EG1_E3_scope_average = 14
 
	Dim ObjPB2G081
	Dim Export_Array
    
    
	I1_b_currency = Request("txtCurrency")
	I2_b_daily_exchange_rate = UNIConvDate(Request("txtValidDt"))
	
    Set ObjPB2G081 = server.CreateObject ("PB2G081.cBListDailyExchRate")    
    on error resume next
    Err.Clear 
    Export_Array = ObjPB2G081.B_LIST_DAILY_EXCHANGE_RATE(gStrGlobalCollection,I1_b_currency,I2_b_daily_exchange_rate)
    Set ObjPB2G081 = nothing

    If CheckSYSTEMError(Err,True) = True Then                               
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
    End If
    on error goto 0
%>
<Script Language=vbscript>
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp
    Dim strData

	With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
		
		LngMaxRow = 0


<%      
        GroupCount = Ubound(Export_Array,1)
	    For LngRow = 0 To GroupCount
%>
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(Trim(Export_Array(LngRow,B366_EG1_E3_apprl_dt )))%>" '1

		strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B366_EG1_E1_currency )))%>"   '2
		strData = strData & Chr(11) & " " '2 PopupButton
		strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B366_EG1_E2_currency )))%>" '4
		strData = strData & Chr(11) & " " '4 PopupButton
		strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B366_EG1_E3_multi_divide )))%>"  '6

		strData = strData & Chr(11) & "<%=UNINumClientFormat(Trim(Export_Array(LngRow,B366_EG1_E3_std_rate )), ggQty.DecPoint,0)%>"      '7
		strData = strData & Chr(11) & "<%=UNINumClientFormat(Trim(Export_Array(LngRow,B366_EG1_E3_buy_rate )), ggQty.DecPoint,0)%>"      '8
		strData = strData & Chr(11) & "<%=UNINumClientFormat(Trim(Export_Array(LngRow,B366_EG1_E3_sell_rate )), ggQty.DecPoint,0)%>"     '9
		strData = strData & Chr(11) & "<%=UNINumClientFormat(Trim(Export_Array(LngRow,B366_EG1_E3_cash_buy_rate )), ggQty.DecPoint,0)%>"  '10
		strData = strData & Chr(11) & "<%=UNINumClientFormat(Trim(Export_Array(LngRow,B366_EG1_E3_cash_sell_rate )), ggQty.DecPoint,0)%>" '11
		strData = strData & Chr(11) & "<%=UNINumClientFormat(Trim(Export_Array(LngRow,B366_EG1_E3_usd_rate )), ggQty.DecPoint,0)%> "      '12		
		strData = strData & Chr(11) & "<%=UNINumClientFormat(Trim(Export_Array(LngRow,B366_EG1_E3_scope_average )), ggExchRate.DecPoint,0)%>"      '14		
		
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
    
Case CStr(UID_M0002)																'��: ���� ��û�� ���� 
									
	If Request("txtMaxRows") = "" Then
		Call DisplayMsgBox("700117", vbInformation, "", "", I_MKSCRIPT)
		Response.End 
	End If

    Dim Obj2PB2G081
    Dim iErrorPosition

    Set Obj2PB2G081 = server.CreateObject ("PB2G081.cBCtrlDailyExchRate")    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear        
    
    Call Obj2PB2G081.B_CTRL_DAILY_EXCHANGE_RATE(gStrGlobalCollection,strSpread,iErrorPosition)
    Set Obj2PB2G081 = nothing

    If CheckSYSTEMError2(Err,True,iErrorPosition & "��","","","","") = True Then
		Response.End 
    End If
    on error goto 0                                                             
%>
<Script Language=vbscript>
	With parent																		'��: ȭ�� ó�� ASP �� ��Ī�� 
		'window.status = "���� ����"
		.DbSaveOk
	End With
</Script>
<%					

End Select

%>
