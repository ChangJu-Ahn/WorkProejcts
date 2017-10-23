<% 
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Numeric Format등록)
'*  3. Program ID           : B1903mb1.asp
'*  4. Program Name         : B1903mb1.asp
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19028ListNumericFormat
'                             +B19021ControlNumericFormat
'*  7. Modified date(First) : 2000/09/23
'*  8. Modified date(Last)  : 2002/12/10
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************
%>
<% Option Explicit %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd

Dim PB1G031											'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strMode	
Dim strSpread										'☜ : 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim FrmFlag

Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount     

Dim iErrPosition  ''Error Row
DIM Import_Array  ''Import Array
DIM Export_Array  ''Export Array

''Import		           
Const B376_IG_Data_Type = 0   
Const B376_IG_Currency = 1   
Const B376_IG_Module_Cd = 2
Const B376_IG_Form_Type = 3

''Export
Const B376_EG_Data_Cd = 0
Const B376_EG_Data_Type = 1
Const B376_EG_Currency = 2
Const B376_EG_Currency_Desc = 3
Const B376_EG_Module_Cd = 4
Const B376_EG_Module_Nm = 5
Const B376_EG_Form_Cd = 6
Const B376_EG_Form_Type = 7
Const B376_EG_Decimals = 8
Const B376_EG_Rounding_Unit = 9
Const B376_EG_Rounding_Policy = 10
Const B376_EG_Rounding_PolicyNm = 11
Const B376_EG_Data_format = 12 
call LoadBasisGlobalInf()

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strSpread = Request("txtSpread")

Select Case strMode
    Case CStr(UID_M0001)													'☜: 현재 조회/Prev/Next 요청을 받음 
        
        REDIM Import_Array(B376_IG_Form_Type)
        
	    FrmFlag	    = Request("FrmFlag")
	    
	    '-----------------------
        'Data manipulate  area(import view match)
        '-----------------------        
	    Import_Array(B376_IG_Data_Type) = Request("cboDataType")
	    Import_Array(B376_IG_Currency)  = Request("txtCurrency")
	    Import_Array(B376_IG_Module_Cd) = Request("cboModuleCd")
	    Import_Array(B376_IG_Form_Type) = Request("cboFormType")
	    '''''''''''''''''''''''''''
	    Set PB1G031 = Server.CreateObject("PB1G031.cBListNumericFormat")	
	    On Error Resume Next    
%>	    
	<Script Language=vbscript>
        parent.frm1.txtCurrency.value = "<%=ConvSPChars(Request("txtCurrency"))%>"
    	parent.frm1.txtCurrencyNm.value = "<%=ConvSPChars(LookUpCurrency(Request("txtCurrency")))%>"    		
    </Script>
    
<%    
		Err.Clear 
		Export_Array = PB1G031.B_LIST_NUMERIC_FORMAT(gStrGlobalCollection,Import_Array)
		Set PB1G031 = Nothing
		
		If CheckSYSTEMError(Err,True) = True Then                               
			Response.End														'☜: 비지니스 로직 처리를 종료함 
		End If
		On Error Goto 0
		
	    GroupCount = uBound(Export_Array,1) 
	    	    
    
%>

    <Script Language=vbscript>
        Dim LngLastRow      
        Dim LngMaxRow       
        Dim LngRow          
        Dim strTemp
        Dim strData
    	DIM strTempDataFormat
    	
    	With parent																	'☜: 화면 처리 ASP 를 지칭함 
    		LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow     		
    		
<%      
	        For LngRow = 0 To GroupCount
%>
                strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B376_EG_Data_Cd))%>"	'1
                strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B376_EG_Data_Type))%>"	'2
                strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B376_EG_Currency))%>"	'3
                strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B376_EG_Currency_Desc))%>"	'4
                strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B376_EG_Module_Cd))%>"	'5
                strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B376_EG_Module_Nm))%>"	'6
                strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B376_EG_Form_Cd))%>"	'7
                strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B376_EG_Form_Type))%>"	'8
                strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B376_EG_Decimals))%>"           	'9
                
        		''Rounding Unit
                strData = strData & Chr(11) & "<%=UNINumClientFormat(Export_Array(LngRow,B376_EG_Rounding_Unit),Export_Array(LngRow,B376_EG_Decimals)+1,0)%>"
                    
		        strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B376_EG_Rounding_Policy))%>" 	    '11
                strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B376_EG_Rounding_PolicyNm))%>"     	'12
                
                strTempDataFormat = parent.FormatChanging("<%=ConvSPChars(Export_Array(LngRow,B376_EG_Decimals))%>","<%=ConvSPChars(Export_Array(LngRow,B376_EG_Data_Cd))%>")  '****0425수정        		
        		strData = strData & Chr(11) & strTempDataFormat            '**************9

		        strData = strData & Chr(11) & LngMaxRow + <%=LngRow%> + 1										'14
                strData = strData & Chr(11) & Chr(12)	
<%      
            Next
%>    
            
		    If "<%=Request("FrmFlag")%>" = "1" Then
		    	.ggoSpread.Source = .frm1.vspdData		 
		    	.ggoSpread.SSShowData strData
		    Else
		    	.ggoSpread.Source = .frm1.vspdData2
		    	.ggoSpread.SSShowData strData
		    End If
		
		    .frm1.hDataType.value = "<%=Request("cboDataType")%>"
		    .frm1.hCurrency.value = "<%=ConvSPChars(Request("txtCurrency"))%>"					
		    .frm1.hModuleCd.value = "<%=ConvSPChars(Request("cboModuleCd"))%>"
		    .frm1.hFormType.value = "<%=Request("cboFormType")%>"			
	   
		    .DbQueryOk(LngMaxRow + 1)
		    
	    End With
    </Script>	
    
<%  
    Case CStr(UID_M0002)																'☜: 저장 요청을 받음									
        Err.Clear																		'☜: Protect system from crashing
		Dim dateFormat, decimalChar		
		Dim lgIntFlgMode
        
		Dim PB0C008
		Dim strLogCntUser
		
		Err.Clear 
		On Error Resume Next		
			
		Set PB0C008 = Server.CreateObject("PB0C008.CB0C008")
		        
		strLogCntUser = PB0C008.Z_GET_CHECK_LONGIN_USER_COUNT(gStrGlobalCollection)

		If Err.Number <> 0 Then
			Set PB0C008 = Nothing												'☜: ComProxy Unload
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)						'⊙:
			Response.End														'☜: 비지니스 로직 처리를 종료함 
		End If
		    
		Set PB0C008 = Nothing
		    
		If strLogCntUser > 1 Then
%>		
			<Script Language=vbscript>
				With parent
				
					.frm1.txtLogInCnt.value = <%=strLogCntUser%>
					.CheckLogInUser
				End With
			</Script>
<%			
			Response.End 
		End If

        If Request("txtMaxRows") = "" Then
	    	Call DisplayMsgBox("700117", vbInformation, "", "", I_MKSCRIPT)
	    	Response.End 
	    End If
	
	    On Error Resume Next
        Set PB1G031 = Server.CreateObject("PB1G031.cBCtlNumericFormat")    
        
        If CheckSYSTEMError(Err,True) = True Then
            Set PB1G031 = nothing
            Response.End  
        End If	
	    On Error Goto 0
    
        On Error Resume Next
        Call PB1G031.B_CONTROL_NUMERIC_FORMAT(gStrGlobalCollection,strSpread,iErrPosition)
        Set PB1G031 = nothing
        If CheckSYSTEMError2(Err,True,iErrPosition & "행","","","","") = True Then            
            Response.End  
        End If
 	    On Error Goto 0

%>
    <Script Language=vbscript>
    	With parent																		'☜: 화면 처리 ASP 를 지칭함 
    		.DbSaveOk
    	End With
    </Script>
<%					
End Select
%>

<%

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
    I1_b_currency(B251_I1_currency_desc) = ""

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
