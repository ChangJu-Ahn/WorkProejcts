<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : 
'*  3. Program ID           : B1C02mb1
'*  4. Program Name         : 언어코드관리 
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             +B25011ManagePlant
'                             +B25011ManagePlant
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 2000/04/28
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

Dim PB1G051																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strSpread

Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount          
Dim iErrPosition
''Import
Dim str_I1_lang_cd
''Export
Dim Export_Array

Const B427_EG1_E1_lang_cd = 0
Const B427_EG1_E1_lang_nm = 1


call LoadBasisGlobalInf()

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strSpread = Request("txtSpread")

Select Case strMode
    Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 

        str_I1_lang_cd = Request("txtLanguage")
	    '''''''''''''''''''''''''''
	    Set PB1G051 = Server.CreateObject("PB1G051.cBListLanguage")	
	    On Error Resume Next    
%>
    <Script Language=vbscript>
    	parent.frm1.txtLanguageNm.value = "<%=ConvSPChars(LookUpLanguage(Request("txtLanguage")))%>"
    </Script>
<%  
		Err.Clear 
		Export_Array = PB1G051.B_LIST_LANGUAGE(gStrGlobalCollection,str_I1_lang_cd)
		Set PB1G051 = Nothing
		
		If Err.number <> 0 and inStr(Err.Description ,"970000") > 0 then
  	        Call DisplayMsgBox("970000", vbOKOnly, "언어코드", "", I_MKSCRIPT)
  	        Response.End
		Else
		    If CheckSYSTEMError(Err,True) = True Then                               
		        Response.End														'☜: 비지니스 로직 처리를 종료함 
		    End if
		End If
		
		On Error Goto 0
		    
	    GroupCount = uBound(Export_Array,1) 
	    	    
	    'Response.Write "GroupCnt=" & GroupCount
	    'Response.End        
%>
    <Script Language=vbscript>
        Dim LngRow          
        Dim strData
    	
    	With parent
    		LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow  
		
        <%  
        	For LngRow = 0 To GroupCount
        %>      
                strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B427_EG1_E1_lang_cd))%>"	
                strData = strData & Chr(11) & "<%=ConvSPChars(Export_Array(LngRow,B427_EG1_E1_lang_nm))%>"			'3
                strData = strData & Chr(11) & LngMaxRow + <%=LngRow%> + 1                                 '11
                strData = strData & Chr(11) & Chr(12)
        <%      
            Next
        %>    
        	.ggoSpread.Source = .frm1.vspdData 
        	.ggoSpread.SSShowData strData
        	.frm1.hLanguage.value = "<%=Request("txtLanguage")%>"
        	.DbQueryOk    		
    	End With    	
    </Script>	
    
<%  
    Case CStr(UID_M0002)																'☜: 저장 요청을 받음									
        Err.Clear																		'☜: Protect system from crashing

        If Request("txtMaxRows") = "" Then
	    	Call DisplayMsgBox("700117", vbInformation, "", "", I_MKSCRIPT)
	    	Response.End 
	    End If
	
	    On Error Resume Next
        Set PB1G051 = Server.CreateObject("PB1G051.cBControlLanguage")    
        
        If CheckSYSTEMError(Err,True) = True Then
            Set PB1G051 = nothing
            Response.End  
        End If	
	    On Error Goto 0
    
        On Error Resume Next
        Call PB1G051.B_CONTROL_LANGUAGE(gStrGlobalCollection,strSpread,iErrPosition)
        Set PB1G051 = nothing
        If CheckSYSTEMError2(Err,True,iErrPosition & "행","","","","") = True Then            
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


<%
'==============================================================================
' Function : LookUp...
' Description : 저장시 Lookup(Language)
'==============================================================================
Function LookUpLanguage(Byval strCode)
    Const B428_E1_lang_cd = 0
    Const B428_E1_lang_nm = 1

	Dim PB1G051
	Dim Import_Value
	Dim Export_Value
	
    Set PB1G051 = Server.CreateObject ("PB1G051.cBLookLanguage")    
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status
    Export_Value = PB1G051.B_LOOKUP_LANGUAGE(gStrGlobalCollection,strCode)
    Set PB1G051 = nothing    

    If Err.number <> 0 and inStr(Err.Description ,"970000") > 0 then
  	    LookUpLanguage = ""
    Else
        If CheckSYSTEMError(Err,True) = True Then
            Exit Function
	    End If
        on error goto 0

	    LookUpLanguage = Export_Value(B428_E1_lang_nm)
    End If		
End Function
%>
