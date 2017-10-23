<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Count Format)
'*  3. Program ID           : B1902mb1.asp
'*  4. Program Name         : B1902mb1.asp
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19038ListCountFormat
'                             +B19031ControlCountFormat
'*  7. Modified date(First) : 2000/09/18
'*  8. Modified date(Last)  : 2000/09/18
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Lee Seok Gon
'* 11. Comment              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd

Dim PB1G021																	'☆ : 조회용 ComProxy Dll 사용 변수 

Dim strMode	
Dim strSpread																'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim GroupCount          
'''''''''''''''''''''''''''''''''''''''''''''''
Dim iErrPosition
DIM Import_Array
DIM Export_Array
''Import		           
Const B380_IG1_module_cd = 0  ''MODULE_CD
Const B380_IG1_form_type = 1  ''FORM_TYPE    
''Export
Const B380_EG1_Minor_Name = 0
Const B380_EG1_Minor_Cd = 1
Const B380_EG1_Decimals = 2
Const B380_EG1_Rounding_Unit = 3
Const B380_EG1_Rounding_Policy = 4
Const B380_EG1_Rounding_PolicyNm = 5
Const B380_EG1_Data_Format = 6
'''''''''''''''''''''''''''''''''''''''''''''''		
call LoadBasisGlobalInf()

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strSpread = Request("txtSpread")

Select Case strMode
	Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
		''VALUE SETTING		
		REDIM  Import_Array(B380_IG1_form_type)						
		
		Import_Array(B380_IG1_module_cd) = Request("cboModuleCd")
		Import_Array(B380_IG1_form_type) = Request("cboFormType")		
		
		Set PB1G021 = Server.CreateObject("PB1G021.cBListCountFormat")		
		
		On Error Resume Next
		Err.Clear 
		Export_Array = PB1G021.B_LIST_COUNT_FORMAT(gStrGlobalCollection,Import_Array)
		Set PB1G021 = Nothing
		
		If CheckSYSTEMError(Err,True) = True Then                               
			Response.End														'☜: 비지니스 로직 처리를 종료함 
		End If
		On Error Goto 0
		
		GroupCount = uBound(Export_Array,1)
		
%>
	<Script Language=vbscript>
		Dim LngMaxRow       
		Dim LngRow          
		Dim strData
		DIM strTempDataFormat
		
		With parent																	'☜: 화면 처리 ASP 를 지칭함		
			If "<%=Request("cboModuleCd")%>" = "*" And 	"<%=Request("cboFormType")%>" = "I" Then   ''TAB1
				.frm1.txtDec.value  = "<%=ConvSPChars(Export_Array(LngRow,B380_EG1_Decimals))%>"
				.frm1.cboFlag.value = "<%=ConvSPChars(Export_Array(LngRow,B380_EG1_Rounding_Policy))%>"         
								
				CALL parent.txtDec_Change()				
				.DbQueryOk	    	
			Else
				LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow 	
			<%  
				For LngRow = 0 To GroupCount		
			%>   
			    	strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B380_EG1_Minor_Name))%>"    '1                					
					strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B380_EG1_Minor_Cd))%>"	'2
					strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B380_EG1_Decimals))%>"	'5
                    
                    ''Rounding Unit
                    strData = strData & Chr(11) & "<%=UNINumClientFormat(Export_Array(LngRow,B380_EG1_Rounding_Unit),Export_Array(LngRow,B380_EG1_Decimals)+1,0)%>"
					                              
					strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B380_EG1_Rounding_Policy))%>"		'7
					strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B380_EG1_Rounding_PolicyNm))%>"	'7
		
					strTempDataFormat = parent.FormatChanging("<%=ConvSPChars(Export_Array(LngRow,B380_EG1_Decimals))%>") 
					strData = strData & Chr(11) & strTempDataFormat            '**************9

			        strData = strData & Chr(11) & LngMaxRow + <%=LngRow%> + 1  '10
			        strData = strData & Chr(11) & Chr(12)
			        
			<%      
			    Next
			%>        
				.ggoSpread.Source = .frm1.vspdData 
				.ggoSpread.SSShowData strData
		
				.frm1.hModuleCd.value = "<%=ConvSPChars(Request("cboModuleCd"))%>"
				.frm1.hFormType.value = "<%=ConvSPChars(Request("cboFormType"))%>"			
				.DbQueryOk
			End If		
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
        Set PB1G021 = Server.CreateObject("PB1G021.cBControlCountFormat")    
        
        If CheckSYSTEMError(Err,True) = True Then
            Set PB1G021 = nothing
            Response.End  
        End If	
	    On Error Goto 0
    
        On Error Resume Next
        Call PB1G021.B_CONTROL_COUNT_FORMAT(gStrGlobalCollection,strSpread,iErrPosition)
        Set PB1G021 = nothing
        If CheckSYSTEMError2(Err,True,iErrPosition & "행","","","","") = True Then            
            'Response.Write iErrPosition
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
