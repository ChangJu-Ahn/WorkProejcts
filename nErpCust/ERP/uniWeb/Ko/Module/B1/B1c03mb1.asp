<%
'**********************************************************************************************
'*  1. Module Name          : Sale,Production
'*  2. Function Name        : Sales Order,....
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        :
'                             +B1c031ControlMessage
'                             +B1c038ListMessage
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2002/12/16
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************

%>
<% Option Explicit %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd

Err.Clear

Dim PB1G071											'�� : ��ȸ�� ComProxy Dll ��� ���� 

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strSpread 

Dim lgStrPrevLang	' ���� ��  ''LANG_CD
Dim lgStrPrevKey	' ���� ��  ''MSG_TYPE

Dim LngMaxRow		' ���� �׸����� �ִ�Row
Dim LngRow
Dim GroupCount 

'''''''''''''''''''''''''''''''''''''''''''''''
Dim iErrPosition
Dim Import_Array
Const B430_I1_lang_cd = 0    
Const B430_I2_msg_cd = 1     
Const B430_I2_msg_type = 2
Const B430_I2_severity = 3
Const B430_I2_msg_text = 4

Dim Export_Array
Const B430_EG1_E1_lang_cd = 0    
Const B430_EG1_E2_msg_cd = 1     
Const B430_EG1_E2_msg_type = 2
Const B430_EG1_E2_msg_typeNm = 3 
Const B430_EG1_E2_severity = 4
Const B430_EG1_E2_severityNm = 5 
Const B430_EG1_E2_msg_text = 6
''''''''''''''''''''''''''''''''''''''''''''''''
Const C_SHEETMAXROWS_D = 100         

call LoadBasisGlobalInf()

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
strSpread = Request("txtSpread")

Select Case strMode
    Case CStr(UID_M0001)														'��: ���� ��ȸ/Prev/Next ��û�� ���� 
	    
	    Redim Import_Array(B430_I2_msg_text)        
        lgStrPrevKey = Request("lgStrPrevKey")
	    lgStrPrevLang = Request("lgStrPrevLang")
	
	    'If Not(lgStrPrevKey = "" And lgStrPrevLang = "" ) Then
	    '    Import_Array(B430_I2_msg_cd) = lgStrPrevKey
        '    Import_Array(B430_I1_lang_cd) = lgStrPrevLang
	    'Else
	    	Import_Array(B430_I2_msg_cd) = Request("txtCode")
	    	Import_Array(B430_I1_lang_cd) = Request("txtLang")		
	    'End If
	    Import_Array(B430_I2_msg_type) = Request("txtType")
	    Import_Array(B430_I2_severity) = Request("txtLevel")
        Import_Array(B430_I2_msg_text) = Request("txtText")
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
        Set PB1G071 = Server.CreateObject("PB1G071.cBListMessage")
	    On Error Resume Next    
''============================================================================	    
	    
	    Err.Clear
		Export_Array = PB1G071.B_LIST_MESSAGE(gStrGlobalCollection,C_SHEETMAXROWS_D,Import_Array,lgStrPrevLang,lgStrPrevKey)
		Set PB1G071 = Nothing
		
		If CheckSYSTEMError(Err,True) = True Then                               
			Response.End														'��: �����Ͻ� ���� ó���� ������ 
		End If
		On Error Goto 0
		
		GroupCount = uBound(Export_Array,1) 
        
        'Response.Write GroupCount
        'Response.End 
%>
    <Script Language=vbscript>
        Dim LngMaxRow       
        Dim strData

    	With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
    		
    		LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow                                                
<%      
    	For LngRow = 0 To GroupCount
    	    If LngRow < C_SHEETMAXROWS_D Then
%>
    		    strData = strData & Chr(11) & "<%=UCase(RTrim(ConvSPChars(Export_Array(LngRow,B430_EG1_E1_lang_cd))))%>"    '1
    		    strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B430_EG1_E2_msg_cd))%>"     '2
    		    strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B430_EG1_E2_msg_type))%>"   '3
    		    strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B430_EG1_E2_msg_typeNm))%>" '4
    		    strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B430_EG1_E2_severity))%>"   '5
    		    strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B430_EG1_E2_severityNm))%>" '6    		    
    		    strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B430_EG1_E2_msg_text))%>"   '7

    		    strData = strData & Chr(11) & LngMaxRow + <%=LngRow%> + 1
    		    strData = strData & Chr(11) & Chr(12)
    		    
		        .lgStrPrevKey  = "" '1
		        .lgStrPrevLang = "" '2
<%
		    Else
%>		        
		        .lgStrPrevLang  = "<%=ConvSPChars(Export_Array(LngRow,B430_EG1_E1_lang_cd))%>" '1
		        .lgStrPrevKey = "<%=ConvSPChars(Export_Array(LngRow,B430_EG1_E2_msg_cd))%>"  '2
<%		            
            End If    		    
        Next
%>    
    		.ggoSpread.Source = .frm1.vspdData 
    		.ggoSpread.SSShowData strData
            
    		If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData,0) And Not(.lgStrPrevKey = "" And .lgStrPrevLang = "") Then
    			.DbQuery
    		Else
    			.frm1.hLang.value     = "<%=ConvSPChars(Request("txtLang"))%>"
    			.frm1.hMsg.value      = "<%=ConvSPChars(Request("txtCode"))%>"			
    			.frm1.hMsgType.value  = "<%=ConvSPChars(Request("txtType"))%>"
    			.frm1.hMsgLevel.value = "<%=ConvSPChars(Request("txtLevel"))%>"
    			.frm1.hMsgText.value  = "<%=ConvSPChars(Request("txtText"))%>"
    		   
    			.DbQueryOk
    		End If
    	End With
    </Script>	
<%    
    Case CStr(UID_M0002)																'��: ���� ��û�� ���� 
									
	    
	    If Request("txtMaxRows") = "" Then
	    	Call DisplayMsgBox("700117", vbInformation, "", "", I_MKSCRIPT)
	    	Response.End 
	    End If
	    
	    On Error Resume Next
        Set PB1G071 = Server.CreateObject("PB1G071.cBControlMessage")    
        
        If CheckSYSTEMError(Err,True) = True Then
            Set PB1G071 = nothing
            Response.End  
        End If	
	    On Error Goto 0
    
        On Error Resume Next
        Call PB1G071.B_CONTROL_MESSAGE(gStrGlobalCollection,strSpread,iErrPosition)
        Set PB1G071 = nothing
        
        If CheckSYSTEMError2(Err,True,iErrPosition & "��","","","","") = True Then            
            'Response.Write iErrPosition
            Response.End  
        End If
 	    On Error Goto 0                                             '��: Unload Comproxy
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
