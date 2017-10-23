<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           : B82106MB1
'*  4. Program Name         : ǰ�񱸼��ڵ庯���Ƿ���ȸ 
'*  5. Program Desc         : ǰ�񱸼��ڵ庯���Ƿ���ȸ 
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
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../B81/B81COMM.ASP" -->



<%	
call LoadBasisGlobalInf()
'call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
'call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '�� : DBAgent Parameter ���� 
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
	
	
    'On Error Resume Next                                                             '��: Protect system from crashing
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
   
        '-----------------------------
	  ' �� �׸� NAME SET 
	  '-----------------------------
	    Call SubOpenDB(lgObjConn) 
	     
		call GetNameChk("MINOR_NM","B_MINOR","MINOR_CD="&filterVar(Request("txtreq_user"),"''","S") & " AND MAJOR_CD=" & filterVar("Y1006","''","S") ,	Request("txtreq_user"),"txtreq_user","","N") '�Ƿ��� 
		call GetNameChk("item_nm","B_CIS_ITEM_MASTER","item_cd="&filterVar(Request("txtitem_cd"),"''","S"),	Request("txtItem_cd"),"txtItem_cd","","N") 'ǰ��code
		Call SubCloseDB(lgObjConn)  

		call FixUNISQLData()
 		Call QueryData("DATA")	
		
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With parent" & vbCr
		Response.Write "	.ggoSpread.Source = .frm1.vspdData "			& vbCr
		Response.Write "    .frm1.vspdData.Redraw = False   "                  & vbCr
		Response.Write "	.ggoSpread.SSShowData     """ & iRowStr	 & """" & ",""F""" & vbCr  
		Response.Write "	.lgPageNo  = """ & lgPageNo & """" & vbCr  
		Response.Write "	.DbQueryOk " & vbCr 
		Response.Write  "   .frm1.vspdData.Redraw = True " & vbCr   
		Response.Write "End With"		& vbCr
		Response.Write "</Script>"		& vbCr   

	
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
	Const C_SHEETMAXROWS_D  = 1000            
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
 			if i=2 or i=11 or i=12 then 
 			iRowStr = iRowStr &	Chr(11) & UniConvDateDbToCompany(Trim(arr(i,j)),"")
 			else
 			iRowStr = iRowStr &	Chr(11) & ConvSPChars(Trim(arr(i,j)))
 			end if 
			
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
Sub FixUNISQLData()
    Dim strVal
	Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(0,7)                                                 '��: DB-Agent�� ���۵� parameter�� ���� ���� 
    Dim item_id,item_spec,status
    Dim req_id
    
     item_id	= filterVar(Request("txtItem_cd"),"''","S")
     req_id		= filterVar(Request("txtreq_user")&"%","''","S")
     item_spec	= filterVar(Request("txtitem_spec")&"%","''","S")
     status		= Request("rbo_status")
     
	UNISqlId(0) = "B82106MA101" 											' header
	UNIValue(0,0)="REQ_NO,dbo.ufn_GetCodeName('Y1006' , REQ_ID ),REQ_DT,dbo.ufn_s_CIS_GetStatus(STATUS),ITEM_CD,ITEM_NM,ITEM_SPEC,"
	UNIValue(0,0)=UNIValue(0,0) & "  dbo.ufn_GetCodeName('Y1007' ,R_GRADE ),dbo.ufn_GetCodeName('Y1007' ,T_GRADE ),dbo.ufn_GetCodeName('Y1007' ,P_GRADE ),dbo.ufn_GetCodeName('Y1007' ,Q_GRADE ),END_DT,TRANS_DT,REMARK"
	if trim(Request("txtItem_cd"))="" then
		UNIValue(0,1)= ""
	else
		UNIValue(0,1)= " AND ITEM_CD=" & item_id
	end if
	
	UNIValue(0,2)= req_id
	UNIValue(0,3)= item_spec

	UNIValue(0,4)="'"&uniConvDate(Request("txtFromReqDt"))&"' AND '"&uniConvDate(Request("txtTOReqDt"))&"' " '�Ƿ����� 
	
	UNIValue(0,4)="'"&uniConvDate(Request("txtFromReqDt"))&"' AND '"&uniConvDate(Request("txtTOReqDt"))&"' " '�Ƿ����� 
	
	IF trim(Request("txtFromEndDt"))="" AND trim(Request("txtToEndDt"))="" then
		UNIValue(0,5)=" "

	else
		UNIValue(0,5)="AND CONVERT(CHAR(8) ,ISNULL(END_DT,''),112) BETWEEN '"&uniConvDate(Request("txtFromEndDt"))&"' AND '"&uniConvDate(Request("txtToEndDt"))&"' " '�Ϸ�Ⱓ 

	end if
	
	
	if status="'*'" then
		UNIValue(0,6)="" 'status
	else 
		UNIValue(0,6)="AND STATUS IN ("&status&") "
	end if

	UNIValue(0,7)="ORDER BY A.INSRT_DT DESC " 'order by 

    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

End Sub

%>
































<OBJECT RUNAT=server PROGID="prjPublic.cCtlTake" id=lgADF></OBJECT>

