<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrnumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b13mb2.asp
'*  4. Program Name         : ��üǰ���� 
'*  5. Program Desc         :
'*  6. Component List       : +PP1G303.cBMngAltItem.B_MANAGE_ALTERNATIVE_ITEM_Svr 
'*  7. Modified date(First) : 2000/11/06
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Lee Hwa Jung
'* 11. Comment              :
'**********************************************************************************************

'Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
'Response.Buffer = True								'�� : ASP�� ���ۿ� ����Ǿ� �������� �ٷ� Client�� ��������.


'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 

On Error Resume Next														'��: 
Err.Clear																'��: Protect system from crashing

Dim I1_b_item_cd
Dim I2_b_plant_cd
Dim IG0_import_group
Dim pPP1G303
Dim iErrorPosition

Dim intFlgMode

Call LoadBasisGlobalInf() 
Call LoadinfTB19029B("I", "*", "NOCOOKIE", "MB")
Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
 		
intFlgMode = CInt(Request("txtFlgMode"))
    
If intFlgMode = OPMD_CMODE Then
	I2_b_plant_cd = Trim(UCase(Request("txtPlantCd")))
	I1_b_item_cd= Trim(UCase(Request("txtItemCd")))
Else
	I2_b_plant_cd = Trim(UCase(Request("hPlantCd")))
	I1_b_item_cd= Trim(UCase(Request("hItemCd")))
End If
    
    
If I2_b_plant_cd = "" Then
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)				'��:
	Response.End
End If
If I1_b_item_cd = "" Then
	Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)					'��:
	Response.End
End If		
    
IG0_import_group = Request("txtSpread")
    
Set pPP1G303 = Server.CreateObject("PP1G303.cBMngAltItem")    
    
If CheckSYSTEMError(Err, True) = True Then
    
	Response.End 
End if

Call pPP1G303.B_MANAGE_ALTERNATIVE_ITEM_Svr  (gStrGlobalCollection, I1_b_item_cd, _
				I2_b_plant_cd, IG0_import_group)
    
If CheckSYSTEMError2(Err, True, "", "", "", "", "") = True Then
	Set pPP1G303 = Nothing															'��: Unload Component
	Response.End
End If
	
Set pPP1G303 = Nothing
            
Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.DbSaveOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
%>
' Server Side ������ ���⼭ ���� 

'==============================================================================
' ����� ���� ���� �Լ� 
'==============================================================================
<Script Language=vbscript RUNAT=server>

'==============================================================================
' Function : SheetFocus
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
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
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function
</Script>