<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           : B81105MB1
'*  4. Program Name         : ����ڵ��
'*  5. Program Desc         : ����ڵ��
'*  6. Component List       : 
'*  7. Modified date(First) : 2005/01/23
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : lee wol san
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="./B81COMM.ASP" -->


<%	
call LoadBasisGlobalInf()
'call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
'call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

   ' Dim lgOpModeCRUD
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '�� : DBAgent Parameter ���� 
    Dim rs1, rs2, rs3, rs4,rs5
	Dim istrData
	Dim lgStrPrevKey	' ���� �� 
	Dim iLngMaxRow		' ���� �׸����� �ִ�Row
	Dim GroupCount  
    Dim lgPageNo
	Dim iErrorPosition
	Dim arrRsVal(11)
	Dim strSpread
	Dim lgStrGbn 
	Dim iRowStr
	
	
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call HideStatusWnd                                                               '��: Hide Processing message
	
    lgOpModeCRUD  = Request("txtMode") 
	 strSpread = Request("txtSpread")
	 lgStrGbn  = Request("lgStrGbn")
	
						                                              '
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)
             Call SubBizSaveMulti()
    End Select

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
    Dim arr  
    ReDim arr(5)
    Dim lgstrdata
'   on error resume next
   
   
   call FixUNISQLData(lgStrGbn)
   
   
   if lgStrGbn="A" then
		Call QueryData("")	
			Response.Write "<Script Language=vbscript>" & vbCr
			Response.Write "With parent" & vbCr
			Response.Write "	.ggoSpread.Source = .frm1.vspdData1 "			& vbCr
			Response.Write "    .frm1.vspdData1.Redraw = False   "                  & vbCr
			Response.Write "	.ggoSpread.SSShowData     """ & iRowStr	 & """" & ",""F""" & vbCr  
			Response.Write "	.lgPageNo  = """ & lgPageNo & """" & vbCr  
			Response.Write "	.DbQueryOk " & vbCr 
			Response.Write "	.vspdData1_Click  2,  1  " & vbCr 
			Response.Write  "   .frm1.vspdData1.Redraw = True " & vbCr   
			Response.Write "End With"		& vbCr
			Response.Write "</Script>"		& vbCr   
	
	else 
		Call QueryData("�����")	
	
			Response.Write "<Script Language=vbscript>" & vbCr
			Response.Write "With parent" & vbCr
			Response.Write "	.ggoSpread.Source = .frm1.vspdData2 "			& vbCr
			Response.Write "    .frm1.vspdData2.Redraw = False   "                  & vbCr
			Response.Write "	.ggoSpread.SSShowData     """ & iRowStr	 & """" & ",""F""" & vbCr  
			Response.Write "	.lgPageNo  = """ & lgPageNo & """" & vbCr  
			Response.Write "	.DbQueryOk " & vbCr 
			Response.Write  "   .frm1.vspdData2.Redraw = True " & vbCr   
			Response.Write "End With"		& vbCr
			Response.Write "</Script>"		& vbCr   
	
	end if 
	 %>
   <script language="vbScript">
	
	with parent.frm1
	parent.ggoSpread.Source = .vspdData2
	
	   //for j=5 to  8
        // .vspdData1.col = j
      	//if.vspdData1.Text="1" then
		//	parent.ggoSpread.SpreadUnLock j, -1 ,j
		//else
		//	parent.ggoSpread.SpreadLock j, -1 ,j
		//end if
       //next 
  
     end with
   </script>
   <%
End Sub    


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data into Db
'============================================================================================================

	
Sub SubBizSaveMulti()
  
    dim Col,Row
    col =Request("hCol"):Row = Request("hRow")
    on error Resume Next                                                             '��: Protect system from crashing
    Err.Clear 
    '===============
    'item check
    '===============
    call chkGridCd()

    Call PY1G105.B_CIS_CTRL(gStrGlobalCollection,strSpread)
    If CheckSYSTEMError(Err,True) = True Then                                              
		Response.End 
    End If
 
    on error goto 0              
                                                         
%>
<Script Language=vbscript>
	With parent																	    '��: ȭ�� ó�� ASP �� ��Ī�� 
		.DbSaveOk
	    .vspdData1_Click  <%=col%>,  <%=row%>
	End With
</Script>

<%
End Sub


'----------------------------------------------------------------------------------------------------------
' chkGridCd
' Grid CD Value check.
'----------------------------------------------------------------------------------------------------------
sub chkGridCd()
  
    dim RowStr,ColStr
    Dim i
	RowStr=split(strSpread,"")
    Call SubOpenDB(lgObjConn) 
		for i=0 to uBound(RowStr)-1
			ColStr=split(RowStr(i),"")
			if ColStr(0)="C" or ColStr(0)="U" then
			call GetNameChkGrid("USR_NM","Z_USR_MAST_REC","USR_ID='"&ColStr(4)&"' AND USR_KIND='U'" ,ColStr(1),1,"parent.frm1.vspdData2","�����") '
			end if
		next
    Call SubCloseDB(lgObjConn) 
   
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO�� Record Set�̿��Ͽ� Query�� �ϰ� Record Set�� �Ѱܼ� MakeSpreadSheetData()���� Spreadsheet�� �����͸� 
' �Ѹ� 
' ADO ��ü�� �����Ҷ� prjPublic.dll������ �̿��Ѵ�.(�󼼳����� vb�� �ۼ��� prjPublic.dll �ҽ� ����)
'----------------------------------------------------------------------------------------------------------
Sub QueryData(pMsg)
    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim iStr
    
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
	  'Call ServerMesgBox(gDsnNo , vbInformation, I_MKSCRIPT)
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
		Response.end
    End If 
    
  If  rs0.EOF And rs0.BOF  Then
        Call DisplayMsgBox("900014", vbOKOnly, pMsg, "", I_MKSCRIPT)
        
        rs0.Close
        Set rs0 = Nothing
		Response.end
    ELSE

        Call  MakeSpreadSheetData()
    End If  
End Sub


'----------------------------------------------------------------------------------------------------------
'QueryData()�� ���ؼ� Query�� �Ǹ� MakeSpreadSheetData()�� ���ؼ� �����͸� ���������Ʈ�� �ѷ��ִ� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100            
    Dim iLoopCount                                                                     
   
	Dim PvArr,arr
	Dim j,i
	
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
   
   iLoopCount = -1
   ReDim PvArr(C_SHEETMAXROWS_D - 1)
	arr=rs0.getRows()

   iRowStr = ""
   for j=0 to uBound(arr,2) 
		
        iLoopCount =  iLoopCount + 1
     
		for i=0 to uBound(arr,1) 
			iRowStr = iRowStr &	Chr(11) & ConvSPChars(Trim(arr(i,j)))
		next
		iRowStr = iRowStr &	Chr(11) & iLngMaxRow + iLoopCount + 1                             
		iRowStr = iRowStr &	Chr(11) & Chr(12)                          
        
        If iLoopCount < C_SHEETMAXROWS_D Then
	        PvArr(iLoopCount) = iRowStr
        Else
           lgPageNo = lgPageNo + 1
          ' Exit for
        End If
      
	next

	istrData = Join(PvArr, "")
    If iLoopCount < C_SHEETMAXROWS_D Then                                      '��: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF
End Sub
	    
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query�ϱ� ����  DB Agent �迭�� �̿��Ͽ� Query���� ����� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData(pVal)
    Dim strVal
	Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(0,4)                                                 '��: DB-Agent�� ���۵� parameter�� ���� ���� 
    Dim item_acct,item_kind 
     item_acct= filterVar(Request("item_acct"),"''","S")
     item_kind= filterVar(Request("item_kind"),"''","S")
     
     
     if pVal="A" then 
                                                               '    parameter�� ���� ���� ������ 
		UNISqlId(0) = "B81105MA101A" 											' header
		UNIValue(0,0)=" A.ITEM_ACCT,B.MINOR_NM,A.ITEM_KIND,C.MINOR_NM, "
		UNIValue(0,0)=UNIValue(0,0) & " CASE A.ITEM_R WHEN 'Y' THEN '1' ELSE '0' END,"
		UNIValue(0,0)=UNIValue(0,0) & " CASE A.ITEM_T WHEN 'Y' THEN '1' ELSE '0' END,"
		UNIValue(0,0)=UNIValue(0,0) & " CASE A.ITEM_P WHEN 'Y' THEN '1' ELSE '0' END,"
		UNIValue(0,0)=UNIValue(0,0) & " CASE A.ITEM_Q WHEN 'Y' THEN '1' ELSE '0' END "
		
		UNIValue(0,1)=" A.ITEM_ACCT" 'ORDER BY 
 
     elseif pVal="B" then 
       
		UNISqlId(0) = "B81105MA101B" 											' header
		UNIValue(0,0)=" USER_ID,'',B.USR_NM,'',"
		UNIValue(0,0)=UNIValue(0,0) & " CASE A.ITEM_R WHEN 'Y' THEN '1' ELSE '0' END,"
		UNIValue(0,0)=UNIValue(0,0) & " CASE A.ITEM_T WHEN 'Y' THEN '1' ELSE '0' END,"
		UNIValue(0,0)=UNIValue(0,0) & " CASE A.ITEM_P WHEN 'Y' THEN '1' ELSE '0' END,"
		UNIValue(0,0)=UNIValue(0,0) & " CASE A.ITEM_Q WHEN 'Y' THEN '1' ELSE '0' END,REMARK "
		
		UNIValue(0,1)= item_acct
		UNIValue(0,2)= item_kind
		UNIValue(0,3)=" A.INSRT_DT DESC " 'ORDER BY
		
       
     end if
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

End Sub


%>


<OBJECT RUNAT=server PROGID="prjPublic.cCtlTake" id=lgADF></OBJECT>
<OBJECT RUNAT=server PROGID="PY1G105.cBCtrlBiz" id=PY1G105></OBJECT>
