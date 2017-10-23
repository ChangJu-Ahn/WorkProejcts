<% 
'**********************************************************************************************
'*  1. Module Name          : Master Data
'*  2. Function Name        : Master Data
'*  3. Program ID           : B1201ma1.asp
'*  4. Program Name         : Auto Numbering
'*  5. Program Desc         :
'*  6. Comproxy List        :
'                             +B12011CtrlAutoNumberingRule
'                             +B12018ListAutoNumbering
'*  7. Modified date(First) : 2000/09/04
'*  8. Modified date(Last)  : 2002/12/10
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              : 
'**********************************************************************************************
%>
<% Option Explicit %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
Call HideStatusWnd
    
Dim PB1G041												'☆ : 조회용/CUD ComProxy Dll 사용 변수 

Dim strMode												'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strSpread 
Dim strData

Dim LngRow
Dim LngMaxRow
Dim GroupCount          

'''''''''''''''''''''''''''''''''''''''''''''''
Dim iErrPosition
''Import		
Dim str_I_AutoType
Dim str_I_EffDate

''Export
DIM Export_Array

Const B343_EG1_E1_auto_no_type = 0  
Const B343_EG1_E1_effect_from_dt = 1
Const B343_EG1_E1_auto_no_nm = 2
Const B343_EG1_E1_auto_flag = 3
Const B343_EG1_E1_job_prefix = 4
Const B343_EG1_E1_date_type = 5
Const B343_EG1_E1_serial_digit = 6
Const B343_EG1_E1_serial_len = 7

Const B343_EG1_E2_date_info = 8     
Const B343_EG1_E2_serial_no = 9
Const B343_EG1_E2_auto_no = 10

Const B343_EG1_E3_minor_nm = 11

Const B343_EG1_E4_reference = 12  
'''''''''''''''''''''''''''''''''''''''''''''''	
call LoadBasisGlobalInf()

strMode   = Request("txtMode")	
strSpread = Request("txtSpread")											'☜ : 현재 상태를 받음 
LngMaxRow = CInt(Request("txtMaxRows"))										'☜ : 

Select Case strMode
    Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 

	    ''VALUE SETTING		
		str_I_AutoType = Request("txtMinor")
		str_I_EffDate = UNIConvDate(Request("txtValidDt"))
		
		Set PB1G041 = Server.CreateObject("PB1G041.cBListAutoNumbering")		
		
		On Error Resume Next
		Err.Clear 
		Export_Array = PB1G041.B_LIST_AUTO_NUMBERING(gStrGlobalCollection,str_I_AutoType,str_I_EffDate)
		Set PB1G041 = Nothing
		
		If CheckSYSTEMError(Err,True) = True Then                               
			Response.End														'☜: 비지니스 로직 처리를 종료함 
		End If
		On Error Goto 0
		
		GroupCount = uBound(Export_Array,1)
		
		'Response.Write "GroupCnt=" & GroupCount
		'Response.End		
    
        For LngRow = 0 To GroupCount
        	strData = strData & Chr(11) & ConvSPChars(Export_Array(LngRow,B343_EG1_E1_auto_no_type))
        	strData = strData & Chr(11) & " " '2 PopupButton
        		
            If Export_Array(LngRow,B343_EG1_E1_auto_flag) = "O" Then
        	   strData = strData & Chr(11) & "1" '3
            Else
        	   strData = strData & Chr(11) & "0" '3
            End If
                            
        	strData = strData & Chr(11) & ConvSPChars(Export_Array(LngRow,B343_EG1_E3_minor_nm))
        	strData = strData & Chr(11) & ConvSPChars(Export_Array(LngRow,B343_EG1_E4_reference))
        	strData = strData & Chr(11) & UNIDateClientFormat(Export_Array(LngRow,B343_EG1_E1_effect_from_dt))
        	strData = strData & Chr(11) & ConvSPChars(Export_Array(LngRow,B343_EG1_E1_job_prefix))
        	strData = strData & Chr(11) & ConvSPChars(Export_Array(LngRow,B343_EG1_E1_date_type))
        	strData = strData & Chr(11) & ConvSPChars(Export_Array(LngRow,B343_EG1_E1_serial_digit))
        	strData = strData & Chr(11) & ConvSPChars(Export_Array(LngRow,B343_EG1_E1_serial_len))
        	strData = strData & Chr(11) & ConvSPChars(Export_Array(LngRow,B343_EG1_E1_auto_no_nm))
        	strData = strData & Chr(11) & ConvSPChars(Export_Array(LngRow,B343_EG1_E2_date_info))
        	strData = strData & Chr(11) & ConvSPChars(Export_Array(LngRow,B343_EG1_E2_serial_no))
        	strData = strData & Chr(11) & ConvSPChars(Export_Array(LngRow,B343_EG1_E2_auto_no))
        	strData = strData & Chr(11) & LngMaxRow + LngRow + 1
        	strData = strData & Chr(11) & Chr(12)
        Next
%>
    <Script Language=vbscript>
        With parent																	'☜: 화면 처리 ASP 를 지칭함 
    	    .ggoSpread.Source = .frm1.vspdData 
    		.ggoSpread.SSShowData "<%=strData%>"
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
        Set PB1G041 = Server.CreateObject("PB1G041.cBCtrlAutoNoRule")    
        
        If CheckSYSTEMError(Err,True) = True Then
            Set PB1G041 = nothing
            Response.End  
        End If	
	    On Error Goto 0
    
        On Error Resume Next
        Call PB1G041.B_CTRL_AUTO_NUMBERING_RULE(gStrGlobalCollection,strSpread,iErrPosition)
        Set PB1G041 = nothing
        
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

<%
'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function
%>