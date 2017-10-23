<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4112mb2
'*  4. Program Name         : 
'*  5. Program Desc         : Insert, Delete, Update Production Order
'*  6. Comproxy List        : PP4C103_LKO391.cPMngProdOrd
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002-07-02
'*  9. Modifier (First)     : 2002-09-16
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              : Chen, Jae Hyun
'**********************************************************************************************
'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%																					'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
Call HideStatusWnd

On Error Resume Next

' =======================================
' �ٲٱ� oPP4C103 => oPP4C103_LKO391
'        PP4C103 => PP4C103_LKO391
Dim oPP4C103_LKO391
Dim iErrorPosition

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount
Dim iDCount

Dim ii

Err.Clear																		'��: Protect system from crashing

itxtSpread = ""
             
' Call ServerMesgBox("HANC : " &  "100" , vbInformation, I_MKSCRIPT)

iCUCount = Request.Form("txtCUSpread").Count
iDCount  = Request.Form("txtDSpread").Count
             
itxtSpreadArrCount = -1
             
ReDim itxtSpreadArr(iCUCount + iDCount)
             
' Call ServerMesgBox("HANC : " &  "200" , vbInformation, I_MKSCRIPT)
For ii = 1 To iDCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
Next
For ii = 1 To iCUCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
Next

' Call ServerMesgBox("HANC : " &  "300" , vbInformation, I_MKSCRIPT)
itxtSpread = Join(itxtSpreadArr,"")

' Call ServerMesgBox("HANC : " &  "400" , vbInformation, I_MKSCRIPT)
'20080116::hanc Set oPP4C103_LKO391 = Server.CreateObject("PP4C103_LKO391.cPMngProdOrd")
Set oPP4C103_LKO391 = Server.CreateObject("PP4C103_KO441.cPMngProdOrd")     '20080307::hanc::PP4C103_KO441
' Call ServerMesgBox("HANC : " &  "500" , vbInformation, I_MKSCRIPT)

'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err,True) = True Then
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If
	
'Call ServerMesgBox("HANC : " &  "600" , vbInformation, I_MKSCRIPT)
Call oPP4C103_LKO391.P_MANAGE_PRODUCTION_ORDER(gStrGlobalCollection, _
										itxtSpread, _
										, _
										iErrorPosition)
' Call ServerMesgBox("HANC : " &  "700" , vbInformation, I_MKSCRIPT)

If CheckSYSTEMError2(Err, True, iErrorPosition & "��", "", "", "", "") = True Then
	Set oPP4C103_LKO391 = Nothing
	Call Parent.RemovedivTextArea
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "Call parent.SheetFocus(" & iErrorPosition & ",1)" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

' Call ServerMesgBox("HANC : " &  "800" , vbInformation, I_MKSCRIPT)
If Not (oPP4C103_LKO391 Is Nothing) Then
	Set oPP4C103_LKO391 = Nothing								'��: Unload Comproxy	
End If	
        
Response.Write "<Script Language=VBScript>" & vbCrLF
Response.Write "parent.DbSaveOk" & vbCrLF
Response.Write "</Script>" & vbCrLF
Response.End
%>
