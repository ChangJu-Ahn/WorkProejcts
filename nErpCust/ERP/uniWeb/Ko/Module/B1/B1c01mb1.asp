    <%
'**********************************************************************************************
'*  1. Module Name          : Sale,Production
'*  2. Function Name        : Sales Order,....
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        :
'                             +B17021CtrlAutoNumberingRule
'                             +B17028ListAutoNumbering
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2002/12/10
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************
%>
<% Option Explicit %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

Err.Clear

Dim PB1G061										'�� : ��ȸ�� ComProxy Dll ��� ���� 

Dim strMode	
Dim strSpread																'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim LngMaxRow		' ���� �׸����� �ִ�Row
Dim LngRow
Dim GroupCount      

Dim lgStrPrevKey	' ���� �� 
Dim lgStrPrevKey2	
Dim iErrPosition  ''Error Row
'''''''''''''''''''''''''''''''''''''''''''''''
Dim Import_Array
Const B424_IG1_lang_cd = 0
Const B424_IG1_caption_cd = 1
Const B424_IG1_original_caption = 2
Const B424_IG1_abbreviated_text = 3   

Dim Export_Array
Const B424_EG1_lang_cd = 0
Const B424_EG1_caption_cd = 1
Const B424_EG1_original_caption = 2
Const B424_EG1_abbreviated_text = 3
Const B424_EG1_redundant_text = 4
Const B424_EG1_description = 5
''''''''''''''''''''''''''''''''''''''''''''''''
Const C_SHEETMAXROWS_D = 100

call LoadBasisGlobalInf()

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
strSpread = Request("txtSpread")

Select Case strMode
    Case CStr(UID_M0001)														'��: ���� ��ȸ/Prev/Next ��û�� ���� 
    
        Redim Import_Array(B424_IG1_abbreviated_text)
        
        lgStrPrevKey = Request("lgStrPrevKey")
	    lgStrPrevKey2 = Request("lgStrPrevKey2")
	
	    'If Not(lgStrPrevKey = "" And lgStrPrevKey2 = "" ) Then
	    '    Import_Array(B424_IG1_lang_cd) = lgStrPrevKey
        '    Import_Array(B424_IG1_caption_cd) = lgStrPrevKey2
	    'Else
	    	Import_Array(B424_IG1_lang_cd) = Request("txtLang")
	    	Import_Array(B424_IG1_caption_cd) = Request("txtCapCd")		
	    'End If
	    Import_Array(B424_IG1_original_caption) = Request("txtOrgCap")
        Import_Array(B424_IG1_abbreviated_text) = Request("txtShtTxt")
	    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
        Set PB1G061 = Server.CreateObject("PB1G061.cBListDataDictionary")	
	    On Error Resume Next    
	    
	    Err.Clear
		Export_Array = PB1G061.B_LIST_DATA_DICTIONARY(gStrGlobalCollection,C_SHEETMAXROWS_D,Import_Array,lgStrPrevKey,lgStrPrevKey2)
		Set PB1G061 = Nothing
		
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
		            strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B424_EG1_lang_cd ))%>" '1
		            strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B424_EG1_caption_cd))%>" '2
		            strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B424_EG1_original_caption))%>" '3
		            strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B424_EG1_abbreviated_text))%>" '4
		            strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B424_EG1_redundant_text))%>"  '5
		            strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B424_EG1_description))%>"  '6
		
		            strData = strData & Chr(11) & LngMaxRow + <%=LngRow%> + 1
		            strData = strData & Chr(11) & Chr(12)
		            
		            .lgStrPrevKey  = "" '1
		            .lgStrPrevKey2 = "" '2
<%
		        Else
%>		        
		            .lgStrPrevKey  = "<%=ConvSPChars(Export_Array(LngRow,B424_EG1_lang_cd))%>"    '1
		            .lgStrPrevKey2 = "<%=ConvSPChars(Export_Array(LngRow,B424_EG1_caption_cd))%>" '2
<%		            
		        End If
            Next
%>    
		    .ggoSpread.Source = .frm1.vspdData 
		    .ggoSpread.SSShowData strData
		    
            If .frm1.vspdData.MaxRows <  .C_SHEETMAXROWS And Not(.lgStrPrevKey = "" And .lgStrPrevKey2 = "") Then
		    	.DbQuery
		    Else
			    .frm1.hLang.value = "<%=Request("txtLang")%>"
			    .frm1.hCaptionCd.value = "<%=(Request("txtCapCd"))%>"			
			    .frm1.hOrgCaption.value = "<%=ConvSPChars(Request("txtOrgCap"))%>"
			    .frm1.hShortText.value = "<%=ConvSPChars(Request("txtShtTxt"))%>"		
			    
			    .DbQueryOk
		    End if		
	    End With
    </Script>	
    
<%    
    Case CStr(UID_M0002)																'��: ���� ��û�� ����									
    
	    
	    If Request("txtMaxRows") = "" Then
	    	Call DisplayMsgBox("700117", vbInformation, "", "", I_MKSCRIPT)
	    	Response.End 
	    End If
	    
	    On Error Resume Next
        Set PB1G061 = Server.CreateObject("PB1G061.cBCtlDataDictionary")    
        
        If CheckSYSTEMError(Err,True) = True Then
            Set PB1G061 = nothing
            Response.End  
        End If	
	    On Error Goto 0
    
        On Error Resume Next
        Call PB1G061.B_CONTROL_DATA_DICTIONARY(gStrGlobalCollection,strSpread,iErrPosition)
        Set PB1G061 = nothing
        
        If CheckSYSTEMError2(Err,True,iErrPosition & "��","","","","") = True Then            
            'Response.Write iErrPosition
            Response.End  
        End If
 	    On Error Goto 0
	        
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
