
<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : A5406mb1
'*  4. Program Name         : �̰����(�������)
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 2002/11/05
'*  8. Modified date(Last)  : 2002/11/05
'*  9. Modifier (First)     : KIM HO YOUNG
'* 10. Modifier (Last)      : KIM HO YOUNG
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2002/11/05 : ..........
'**********************************************************************************************


Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ����Ǿ� �������� �ٷ� Client�� ��������.


'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 

%>
<%
'#########################################################################################################
'												1. Include
'##########################################################################################################
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'#########################################################################################################
'												2. ���Ǻ� 
'##########################################################################################################

													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd	
On Error Resume Next														'��: 
Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

'#########################################################################################################
'												2.1 ���� üũ 
'##########################################################################################################

If strMode = "" Then
'	Response.End 
End If

'#########################################################################################################
'												2. ���� ó�� ����� 
'##########################################################################################################

'#########################################################################################################
'												2.1. ����, ��� ���� 
'##########################################################################################################
' ������ ���� 
Dim pDelGlCardAcct																'�� : ��ȸ�� ComProxy Dll ��� ���� 

Dim IntRows
Dim IntCols
Dim sList
Dim vbIntRet
Dim intCount
Dim IntCount1
Dim LngMaxRow
Dim LngMaxRow1
Dim StrNextKey
Dim lgStrPrevKey
Dim lgIntFlgMode
dim test

' Com+ Conv. ���� ���� 
Dim pvStrGlobalCollection 
Dim I1_cls_no
Dim arrCount
					'��: ���� ��ȸ/Prev/Next ��û�� ���� 
	'#########################################################################################################
	'												2.2. ��û ���� ó�� 
	'##########################################################################################################
	lgStrPrevKey = Request("lgStrPrevKey")

	'#########################################################################################################
	'												2.3. ���� ó�� 
	'##########################################################################################################

	Set pDelGlCardAcct = Server.CreateObject("PAUG035.cADelGlCardAcctSvr")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Set pDelGlCardAcct = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)	'��:
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

		LngMaxRow  = CLng(Request("txtMaxRows"))												'��: Fetechd Count      
		LngMaxRow1  = CLng(Request("txtMaxRows1"))

		I1_cls_no = Request("txtClsNo")

		On Error Resume next
		Call pDelGlCardAcct.A_DELETE_GL_CARD_ACCT_SVR(gStrGlobalCollection,Trim(I1_cls_no))
						
	'-----------------------
	'Com Action Area
	'-----------------------

		If CheckSYSTEMError(Err,True) = True Then
		
			Set pDelGlCardAcct = Nothing																	'��: ComProxy Unload
			Response.End																			'��: �����Ͻ� ���� ó���� ������ 
		End If

		Set pDelGlCardAcct = Nothing

    Response.Write "<Script Language=VBScript> " & vbCr         
    Response.Write "With parent "				 & vbCr	
	Response.Write " .DbDeleteOK() "								  & vbCr
    Response.Write "End With "				 & vbCr	  
    Response.Write "</Script>"       																	  & vbCr	
		
	%>		

