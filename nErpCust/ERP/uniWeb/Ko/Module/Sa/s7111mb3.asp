<%
'********************************************************************************************************
'*  1. Module Name          : Sales & Distribution														*
'*  2. Function Name        :																			*
'*  3. Program ID           : S7111MB3    																*
'*  4. Program Name         : NEGO 등록																	*
'*  5. Program Desc         : 수출 B/L 등록 회계 Posting Transaction 처리용 ASP							*
'*  7. Modified date(First) : 2000/05/20																*
'*  8. Modified date(Last)  : 2000/05/20																*
'*  9. Modifier (First)     : An Chang Hwan																*
'* 10. Modifier (Last)      : An Chang Hwan																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/05/20 : Coding Start												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->
<%
																				'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")   
Call HideStatusWnd                                                                 '☜: Hide Processing message

Dim S71115																		' 수출통관 OpenAR Posting용 Object
Dim I2_s_nego_nego_no 

Err.Clear																		'☜: Protect system from crashing

'---------------------------------- C/C Header Data Query ----------------------------------

I2_s_nego_nego_no= Trim(Request("txtNegoNo"))

Set S71115 = Server.CreateObject("PSAG115.SNegoPostSvr")

If CheckSYSTEMError(Err,True) = True Then 
	Response.End						
End if

Call S71115.S_POST_NEGO_SVR( gStrGlobalCollection, I2_s_nego_nego_no )

If CheckSYSTEMError(Err,True) = True Then 
	Set S71115 = Nothing
	Response.End						
End if


'-----------------------
'Result data display area
'-----------------------
%>
<Script Language=VBScript>
	Call parent.PostingOk()														'☜: 회계Posting 성공 
</Script>
<%
Set S71115 = Nothing															'☜: ComProxy UnLoad

Response.End																	'☜: Process End
%>