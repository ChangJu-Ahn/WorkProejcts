<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name			: Inventory
'*  2. Function Name		: 
'*  3. Program ID			: B1B07MB2.asp
'*  4. Program Name			: Basic Data
'*  5. Program Desc			: Save B_ITEM_ACCT_TRACKING
'*  6. Comproxy List		: PB3S116.cBSetItemAcctTracking
'*  7. Modified date(First)	: 2006/06/27
'*  8. Modified date(Last) 	: 2006/06/27
'*  9. Modifier (First)		: LEE SEUNG WOOK
'* 10. Modifier (Last)		: LEE SEUNG WOOK
'* 11. Comment		:
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call LoadBasisGlobalInf

Call HideStatusWnd

On Error Resume Next

Dim oPB3S116
Dim iErrorPosition
Dim itxtSpread

Err.Clear

itxtSpread = ""

itxtSpread = Request("txtSpread")

Set oPB3S116 = Server.CreateObject("PB3S116.cBSetItemAcctTracking")
    
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If
	
Call oPB3S116.B_MANAGE_ITEMACCT_TRACKING(gStrGlobalCollection, _
								itxtSpread, _
								iErrorPosition)

If CheckSYSTEMError(Err,True) = True Then
	Set oPB3S116 = Nothing
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call parent.SheetFocus(" & iErrorPosition & ",1)" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

If Not (oPB3S116 Is nothing) Then
	Set oPB3S116 = Nothing
End If
        
Response.Write "<Script Language=VBScript>" & vbCrLF
Response.Write "Call parent.DbSaveOk" & vbCrLF
Response.Write "</Script>" & vbCrLF
Response.End
%>