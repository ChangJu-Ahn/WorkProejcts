<%
'======================================================================================================
'*  1. Module Name          : accounting
'*  2. Function Name        : 
'*  3. Program ID           : a7125ma1
'*  4. Program Name         : �����ڻ����󼼳������ 
'*  5. Program Desc         : �����ڻ����󼼳������ 
'*  6. Modified date(First) : 2000/08/23
'*  7. Modified date(Last)  : 2000/08/23
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'======================================================================================================= -->
Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next								'��: 

Dim pAS0101											'�Է�/������ ComProxy Dll ��� ���� 

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim StrNextKey		' ���� �� 
Dim lgStrPrevKey	' ���� �� 
Dim LngMaxRow		' ���� �׸����� �ִ�Row
Dim LngRow
Dim GroupCount          
Dim iPAAG011


Call HideStatusWnd

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")


strMode = Request("txtMode")						'�� : ���� ���¸� ���� 

GetGlobalVar

On Error Resume Next

Select Case strMode

    
Case CStr(UID_M0002)																'��: ���� ��û�� ���� 
									
    Err.Clear																		'��: Protect system from crashing

    LngMaxRow = CInt(Request("txtMaxRows_2"))											'��: �ִ� ������Ʈ�� ���� 

    Set iPAAG011 = Server.CreateObject("PAAG011.cAMngAsItmSvr")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		set iPAAG011 = Nothing
		Response.End
	End IF	
	
	arrTemp = Split(Request("txtSpread_m"), gRowSep)									'Spread Sheet ������ ��� �ִ� Element�� 

    call iPAAG011.A_MANAGE_ASSET_ITEM_SVR( gStrGloBalCollection , Request("txtSpread_m"))
            
        '-----------------------
        'Com action result check area(OS,internal)
        '-----------------------                   
    If CheckSYSTEMError(Err,True) = True Then
		set pAS0101 = Nothing
		Response.End			
	End IF
                                              

    Set iPAAG011 = Nothing														    '��: Unload Comproxy
    
%>
<Script Language=vbscript>
	With parent																	    '��: ȭ�� ó�� ASP �� ��Ī�� 
		.DbSaveOk
	End With
</Script>
<%					

End Select

%>
