<%@ Language=VBScript %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sale,Production
'*  2. Function Name        : Sales Order,....
'*  3. Program ID           :
'*  4. Program Name         : Master L/C 등록 Transaction 처리용 ASP
'*  5. Program Desc         :
'*  6. Comproxy List        : +B17013
'*  7. Modified date(First) : 2000/09/14
'*  8. Modified date(Last)  : 2000/09/14
'*  9. Modifier (First)     : Kim Hyungsuk
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'**********************************************************************************************

Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

On Error Resume Next
	
'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncServer.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 


	Call HideStatusWnd

	Dim strMode
	
	Dim b17013										' Master L/C Header 조회용 Object
	Dim b17014										' Master L/C Header 조회용 Object
	Dim Row

	strMode = Trim(Request("StrMode"))

	Set b17013 = Server.CreateObject("B17013.B17013CalcExchRateByUser")

	'-----------------------------------------------------------------------------------
	'Com action result check area(OS,internal)
	'-----------------------------------------------------------------------------------
	    If Err.Number <> 0 Then
			Set b17013 = Nothing												'☜: ComProxy UnLoad
			Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)		'⊙:
			Response.End														'☜: Process End
		End If
	    
	'-----------------------------------------------------------------------------------
	'Data manipulate  area(import view match)
	'-----------------------------------------------------------------------------------
	    
	    b17013.ImportBCurrencyCurrency = Request("txtCurrency")
	    b17013.ImportBDailyExchangeRateApprlDt = UNIConvDate(Request("txtApplDt"))
	    b17013.ImportBDailyExchangeRateStdRate = UNIConvNum(Request("txtXchRate"), 0)
	    b17013.ImportBNumericFormatDataType = "2"
	    b17013.ImportExchangeVariableNumValue152 = UNIConvNum(Request("txtDocAmt"), 0)
	    b17013.ImportToBCurrencyCurrency = Request("txtLocCurrency")
		Row = Request("Row")
		
	'------------------------------------------------------------------------------------
	'Com action area
	'------------------------------------------------------------------------------------
	    b17013.ComCfg = gConnectionString
	    b17013.Execute	

	'------------------------------------------------------------------------------------
	'Com action result check area(OS,internal)
	'------------------------------------------------------------------------------------
		If Err.Number <> 0 Then
		   Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)      '⊙:
		   Set b17013 = Nothing													'☜: ComProxy UnLoad
		   Response.End															'☜: Process End
		End If

	'-------------------------------------------------------------------------------------
	'Com action result check area(DB,internal)
	'-------------------------------------------------------------------------------------
		If Not (b17013.OperationStatusMessage = MSG_OK_STR) Then
		   Call DisplayMsgBox(b17013.OperationStatusMessage, vbInformation, "", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
		   Set b17013 = Nothing																		'☜: ComProxy UnLoad
		   Response.End																				'☜: Process End
		End If
		
		If Request("txtAmendFlg") = "LC" Then

%>
			<Script Language=VBScript>
				With parent
					.ggoSpread.Source = .frm1.vspdData
					
					.frm1.vspdData.Row = <%=Row%>
					.frm1.vspdData.Col = .C_LocAmt
					msgbox "444"
					If "<%=UNINumClientFormat(b17013.ExportExchangeVariableNumValue152, ggAmtOfMoney.DecPoint, 0)%>" <> "" Then
						.frm1.vspdData.text = "<%=UNINumClientFormat(b17013.ExportExchangeVariableNumValue152, ggAmtOfMoney.DecPoint, 0)%>" 
						
						'msgbox "<%=b17013.ExportExchangeVariableNumValue152%>"
						'msgbox "<%=UNINumClientFormat(b17013.ExportExchangeVariableNumValue152, ggAmtOfMoney.DecPoint, 0)%>" 			
					End If
				
					
				
					Call .SumLCAmt()	
				End	With
			</Script>
<%
			Set b17013 = Nothing														'☜: Unload Comproxy
			Response.End																'☜: Process End
		ElseIf Request("txtAmendFlg") = "AMEND" Then
%>
			<Script Language=VBScript>
				With parent
					.ggoSpread.Source = .frm1.vspdData
					
					.frm1.vspdData.Row = <%=Row%>
					.frm1.vspdData.Col = .C_AtLocAmt

					If "<%=UNINumClientFormat(b17013.ExportExchangeVariableNumValue152, ggAmtOfMoney.DecPoint, 0)%>" <> "" Then
						.frm1.vspdData.text = "<%=UNINumClientFormat(b17013.ExportExchangeVariableNumValue152, ggAmtOfMoney.DecPoint, 0)%>" 
					End If
					Call .TotalSum()	
				End	With
			</Script>
<%
			Set b17013 = Nothing														'☜: Unload Comproxy
			Response.End																'☜: Process End
		End If
%>

	
