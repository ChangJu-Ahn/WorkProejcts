<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b11mb9.asp
'*  4. Program Name         : Update Item by Plant (Detail)
'*  5. Program Desc         :
'*  6. Component List       : PB3S111.cBMngItemByPltDtl 
'*  7. Modified date(First) : 2001/03/13
'*  8. Modified date(Last)  : 2002/11/14
'*  9. Modifier (First)     : Jung Yu Kyung
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call LoadBasisGlobalInf

On Error Resume Next
Err.Clear

Call HideStatusWnd

Dim pPB3G108

Dim strSpread
Dim iErrorPosition

strSpread = Request("txtSpread")

'-----------------------
'Com action result check area(OS,internal)
'-----------------------
Set pPB3G108 = Server.CreateObject("PB3G108.cBMngItemMulti")    
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPB3G108.B_MANAGE_ITEM_MULTI(gStrGlobalCollection, strSpread, iErrorPosition)

Select Case Trim(Cstr(Err.Description))
	Case "B_MESSAGE" & Chr(11) & "970023"
		Call DisplayMsgBox("970023", vbInformation, iErrorPosition & "青 : 前格辆丰老", "傍厘喊前格辆丰老", I_MKSCRIPT)
		Set pPB3G108 = Nothing
		Response.End
	Case Else
		If CheckSYSTEMError2(Err, True, iErrorPosition & "青 : ", "", "", "", "") = True Then
			Set pPB3G108 = Nothing
			Response.End
		End If
End Select

Set pPB3G108 = Nothing

Response.Write "<Script Language=vbscript>" & vbCrLf
Response.Write "	parent.DbSaveOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End

%>
