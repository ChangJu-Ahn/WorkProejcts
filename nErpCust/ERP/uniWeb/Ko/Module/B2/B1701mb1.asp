<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Currency)
'*  3. Program ID           : B1701mb1
'*  4. Program Name         : 통화코드관리 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'                             +B17011ControlCurrency
'                             +B17018ListCurrency
'*  7. Modified date(First) : 2000/09/05
'*  8. Modified date(Last)  : 2002/12/16
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************

%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

On Error Resume Next														'☜: 

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strSpread
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount          

Call LoadBasisGlobalInf()

strMode = Request("txtMode")
strSpread = Trim(Request("txtSpread"))
'Response.Write "||strMode:" & strMode												'☜ : 현재 상태를 받음 
'Response.Write "||UID_M0001:"& UID_M0001
'Response.Write "||UID_M0002:"& UID_M0002
'Response.End 


Select Case strMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
    Const B362_EG1_E1_currency = 0 
    Const B362_EG1_E1_currency_desc = 1
 

	Dim ObjPB2G061
    Dim I1_b_currency_currency
	Dim Export_Array

	I1_b_currency_currency = Request("txtCurrency")
        
%>
<Script Language=vbscript>
	parent.frm1.txtCurrencyNm.value = "<%=ConvSPChars(LookUpCurrency(Request("txtCurrency")))%>"
</Script>
<%  

    Set ObjPB2G061 = server.CreateObject ("PB2G061.cBListCurrency")    
    on error resume next
    Err.Clear 
    Export_Array = ObjPB2G061.B_LIST_CURRENCY(gStrGlobalCollection,I1_b_currency_currency)
    Set ObjPB2G061 = nothing

    If CheckSYSTEMError(Err,True) = True Then                               
		Response.End														'☜: 비지니스 로직 처리를 종료함 
    End If
    on error goto 0
    
%>
<Script Language=vbscript>

    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp
    Dim strData
	
	With parent																	'☜: 화면 처리 ASP 를 지칭함 
		'.Clear
		 LngMaxRow = 0
<%      
        GroupCount = Ubound(Export_Array,1)
	    For LngRow = 0 To GroupCount
%>        
            strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B362_EG1_E1_currency)))%>"'CURRENCY
            strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B362_EG1_E1_currency_desc)))%>"'CURRENCY_DESC
            strData = strData & Chr(11) & LngMaxRow + <%=LngRow%> + 1
            strData = strData & Chr(11) & Chr(12)
            
<%      		
        Next
%>    		
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData strData
		.frm1.hCurrency.value = "<%=ConvSPChars(Request("txtCurrency"))%>"
		.DbQueryOk
		
	End With

</Script>	
<%    
    
Case CStr(UID_M0002)																'☜: 저장 요청을 받음 
									
    Dim Obj2PB2G061
    Dim iErrorPosition

    Set Obj2PB2G061 = server.CreateObject ("PB2G061.cBControlCurrency")    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear        
    Call Obj2PB2G061.B_CONTROL_CURRENCY(gStrGlobalCollection,strSpread,iErrorPosition)
    Set Obj2PB2G061 = nothing

    If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
		Response.End 
    End If
    on error goto 0                                                                  '☜: Unload Comproxy
    
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


<%
'==============================================================================
' Function : LookUp...
' Description : 저장시 Lookup
'==============================================================================
Function LookUpCurrency(Byval strCode)
    Const B251_I1_currency = 0
    Const B251_I1_currency_desc = 1

    Const B251_E1_currency = 0
    Const B251_E1_currency_desc = 1

	Dim ObjPB0C003	
	Dim I1_b_currency
	Dim E1_b_currency
	
    ReDim I1_b_currency(B251_I1_currency_desc)
    ReDim E1_b_currency(B251_E1_currency_desc)
    
    I1_b_currency(B251_I1_currency) = strCode
    I1_b_currency(B251_I1_currency_desc)= ""

    Set ObjPB0C003 = server.CreateObject ("PB0C003.CB0C003")    
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status
    E1_b_currency = ObjPB0C003.B_SELECT_CURRENCY (gStrGlobalCollection,I1_b_currency)
    Set ObjPB0C003 = nothing    

    If Err.number <> 0 and inStr(Err.Description ,"121400") > 0 then
  	LookUpCurrency = ""
    Else
        If CheckSYSTEMError(Err,True) = True Then                                              
        	Exit Function
	    End If
        on error goto 0

	    LookUpCurrency = E1_b_currency(B251_E1_currency_desc)
    End If
End Function
%>
