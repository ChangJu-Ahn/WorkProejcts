<% 
'======================================================================================================
'*  1. Module Name          : CIS
'*  2. Function Name        : ǰ���ڵ屸������ 
'*  3. Program ID           : B81101Mb1.asp
'*  4. Program Name         : B81101Mb1.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 2004/01/25
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Lee Wol san
'*  9. Modifier (Last)      :
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

%>
<% Option Explicit %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

Call HideStatusWnd									'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strSpread
Dim StrNextKey		' ���� �� 
Dim lgStrPrevKey	' ���� �� 
Dim LngMaxRow		' ���� �׸����� �ִ�Row
Dim LngRow         
Dim RowData(5)
Dim RowDataPre
Dim lgSelectList
call LoadBasisGlobalInf()

    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    'Multi SpreadSheet
    strSpread = Request("txtSpread")
    lgLngMaxRow       = Request("txtMaxRows")    
   
                            
  
	Call SubOpenDB(lgObjConn)                                                       '��: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)    
                                                             '��: Query
             Call SubBizQuery()
             
        Case CStr(UID_M0002) 
                                                                '��: Save,Update
             Call SubBizSave()
             Call SubBizSaveMulti()
             
        Case CStr(UID_M0003)
                                                                 '��: Delete
             Call SubBizDelete()
    End Select

Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection

Sub SubBizQuery()
   ' On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    Call SubBizQueryMulti()
End Sub    

Sub SubBizSave()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub

Sub SubBizDelete()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub
'-----------------------------------------------------------------------------------------
Sub SubBizQueryMulti()
'-----------------------------------------------------------------------------------------
    Dim strData
    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    DIM arr,arrCnt
    dim i,j
    'BlankchkFlg = False
    
    'On Error Resume Next  
	Err.Clear                                                                  '��: Clear Error status
	LngRow = 0
	Call SubMakeSQLStatements()
	Call SubOpenDB(lgObjConn)   
	
		adoRec.Open lgStrSQL, lgObjConn,adOpenStatic, adLockReadOnly 
		IF adoRec.RecordCount >0 then
			arr=adoRec.GetRows 
			arrCnt = adoRec.RecordCount 
		end if
		
		
	Call SubCloseDB(lgObjConn)   
	if arrCnt>0 then
		for i=0 to arrCnt-1
		   LngRow = LngRow + 1
			for j=0 to uBound(arr,1)
			strData = strData & Chr(11) & arr(j,i)
			next 
			strData =  strData & Chr(11) & LngRow &  Chr(11) & Chr(12) 
		next 
	else
	Call DisplayMsgBox("971001", vbOKOnly, "ǰ���ڵ屸������", "", I_MKSCRIPT)
    end if
%>

<Script Language=vbscript>
    Dim LngRow          
    Dim strTemp
    Dim strData
	With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	.frm1.vspdData.ReDraw = False
	 strData = "<%=ConvSPChars(strData)%>"
    .ggoSpread.Source = .frm1.vspdData 
  	.ggoSpread.SSShowData strData
	.frm1.vspdData.ReDraw = True
	.DbQueryOk
	End With
</Script>	
<% 

End Sub    
%>	
	
<%	
Sub SubBizSaveMulti()
		
  dim iErrorPosition
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear 
    Call ObjPY1G101.B_CIS_Ctrl(gStrGlobalCollection,strSpread,iErrorPosition)
    
   If CheckSYSTEMError(Err,True) = True Then                                              
		Response.End 
    End If
	
                                                           
%>
<Script Language=vbscript>
	With parent																	    '��: ȭ�� ó�� ASP �� ��Ī�� 
		.DbSaveOk
	End With
</Script>
<%					

End Sub


Sub SubMakeSQLStatements()

	lgStrSQL = lgStrSQL & "SELECT  A.ITEM_ACCT,B.MINOR_NM,'',A.ITEM_KIND,C.MINOR_NM,'',"
	lgStrSQL = lgStrSQL & " A.ITEM_LVL1,A.ITEM_LVL2 ,A.ITEM_LVL3,A.ITEM_SEQNO,A.ITEM_LVL_D, A.ITEM_VER,"
	lgStrSQL = lgStrSQL & " A.ITEM_LVL1+A.ITEM_LVL2 +A.ITEM_LVL3+A.ITEM_SEQNO+A.ITEM_LVL_D+ A.ITEM_VER TOTAL,"
	lgStrSQL = lgStrSQL & " CASE A.ITEM_R WHEN 'Y' THEN '1' ELSE '0' END,"
	lgStrSQL = lgStrSQL & " CASE A.ITEM_T WHEN 'Y' THEN '1' ELSE '0' END,"
	lgStrSQL = lgStrSQL & " CASE A.ITEM_P WHEN 'Y' THEN '1' ELSE '0' END,"
	lgStrSQL = lgStrSQL & " CASE A.ITEM_Q WHEN 'Y' THEN '1' ELSE '0' END,"
	lgStrSQL = lgStrSQL & " CREATE_ITEM,"
	lgStrSQL = lgStrSQL & " CASE CREATE_ITEM WHEN 'R' THEN '����' WHEN 'T' THEN '���' "
	lgStrSQL = lgStrSQL & " WHEN 'P' THEN '����' WHEN 'Q' THEN 'ǰ��' END"
	lgStrSQL = lgStrSQL & " FROM B_CIS_CONFIG A "
	lgStrSQL = lgStrSQL & " INNER JOIN B_MINOR B ON A.ITEM_ACCT=B.MINOR_CD AND    B.MAJOR_CD = N'P1001' "
	lgStrSQL = lgStrSQL & " INNER JOIN B_MINOR C ON A.ITEM_KIND=C.MINOR_CD AND    C.MAJOR_CD = N'Y1001' "

End Sub


%>

<OBJECT RUNAT=server PROGID="PY1G101.cBCtrlBiz" id=ObjPY1G101></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=adoRec></OBJECT>