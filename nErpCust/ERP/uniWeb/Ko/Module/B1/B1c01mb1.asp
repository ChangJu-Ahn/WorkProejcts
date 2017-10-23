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
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Err.Clear

Dim PB1G061										'☆ : 조회용 ComProxy Dll 사용 변수 

Dim strMode	
Dim strSpread																'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount      

Dim lgStrPrevKey	' 이전 값 
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

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strSpread = Request("txtSpread")

Select Case strMode
    Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
    
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
			Response.End														'☜: 비지니스 로직 처리를 종료함 
		End If
		On Error Goto 0
		
		GroupCount = uBound(Export_Array,1) 
        
        'Response.Write GroupCount
        'Response.End 
%>

    <Script Language=vbscript>
        Dim LngMaxRow
        Dim strData

    	With parent																	'☜: 화면 처리 ASP 를 지칭함    		
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
    Case CStr(UID_M0002)																'☜: 저장 요청을 받음									
    
	    
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
        
        If CheckSYSTEMError2(Err,True,iErrPosition & "행","","","","") = True Then            
            'Response.Write iErrPosition
            Response.End  
        End If
 	    On Error Goto 0
	        
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
