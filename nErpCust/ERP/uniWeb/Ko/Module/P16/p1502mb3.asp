<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1502mb2.asp
'*  4. Program Name         : ManageResourceGroup ���� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/09/07
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : Hong Eun Sook
'* 10. Modifier (Last)      : RYU SUNG WON
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call LoadBasisGlobalInf

Dim oPP1G604				'�� : ����� ComProxy Dll ��� ���� 

Dim StrNextKey				' ���� �� 
Dim lgStrPrevKey			' ���� �� 
Dim LngRow
Dim GroupCount 

Dim lgIntFlgMode
Dim iCommandSent
Dim I1_P_Resource_Group		'Array
Dim I2_B_Plant_Plant_Cd		'String

ReDim I1_P_Resource_Group(7)
Const C_I1_resource_group_cd = 0    '[CONVERSION INFORMATION]  View Name : import p_resource_group

Call HideStatusWnd

On Error Resume Next
Err.Clear

    If Request("txtResourceGroupCd2") = "" Then						'��: ��ȸ�� ���� ���� ���Դ��� üũ 
		Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)      
		Response.End 
	End If
	
	iCommandSent = "DELETE"
	I1_P_Resource_Group(C_I1_resource_group_cd) = Request("txtResourceGroupCd2")
	I2_B_Plant_Plant_Cd = Request("txtPlantCd")
	
	Set oPP1G604 = Server.CreateObject("PP1G604.cPMngRsrcGrp")
    
    If CheckSYSTEMError(Err,True) = True Then
       Response.End 
    End If
    
    Call oPP1G604.P_MANAGE_RESOURCE_GROUP(gStrGlobalCollection, _
									iCommandSent, _
									I1_P_Resource_Group, _
									I2_B_Plant_Plant_Cd)
	
	If CheckSYSTEMError(Err,True) = True Then
       Set oPP1G604 = Nothing		                                '��: Unload Comproxy DLL
       Response.End		
    End If
   
    Set oPP1G604 = Nothing	

				
Response.Write "<Script Language=vbscript>" & vbCr
Response.Write "	With parent			  " & vbCr					'��: ȭ�� ó�� ASP �� ��Ī�� 
Response.Write "		.DbDeleteOk		  " & vbCr
Response.Write " 	End With			  " & vbCr
Response.Write "</Script>				  " & vbCr

%>
<Script Language=vbscript RUNAT=server>

</Script>