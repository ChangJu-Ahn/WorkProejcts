<%
'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : A_RECEIPT
'*  3. Program ID        : f5122mb2
'*  4. Program �̸�      : ���������̵�ó�� 
'*  5. Program ����      : ���������̵�ó�� 
'*  6. Comproxy ����Ʈ   : f5122ma1
'*  7. ���� �ۼ������   : 2000/10/16
'*  8. ���� ���������   : 2002/02/15
'*  9. ���� �ۼ���       : ������ 
'* 10. ���� �ۼ���       :  
'* 11. ��ü comment      :
'* 12. ���� Coding Guide : this mark(��) means that "Do not change"
'*                         this mark(��) Means that "may  change"
'*                         this mark(��) Means that "must change"
'* 13. History           :
'*                         
'**********************************************************************************************
'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>

<!-- #Include file="../../inc/IncServer.asp"  -->

<%					

'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next															' ��: 
Err.Clear 

Dim lgADF																		'�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg																	'�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0									'�� : DBAgent Parameter ���� 


Call HideStatusWnd																'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

'Dim LngMaxRow																	' ���� �׸����� �ִ�Row
Dim LngRow

Dim arrVal, arrTemp																'��: Spread Sheet �� ���� ���� Array ���� 
Dim strStatus																	'��: Sheet �� ���� Row�� ���� (Create/Update/Delete)
Dim	lGrpCnt																		'��: Group Count

Dim strMode						'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim StrNextKeyNoteNo			' NoteNO ���� �� 
Dim StrNextKeyGlNo				' GLNO ���� �� 
Dim lgStrPrevKeyNoteNo			' Note NO ���� �� 
Dim lgStrPrevKeyGlNo

Dim strNoteNo,strFrBizCd 
Dim strWhere0
Dim strMsgCd, strMsg1, strMsg2

Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Const GroupCount = 30

strMode = Request("txtMode")	'�� : ���� ���¸� ���� 

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)

	lgStrPrevKeyNoteNo = "" & UCase(Trim(Request("lgStrPrevKeyNoteNo")))
	lgStrPrevKeyGlNo = "" & UCase(Trim(Request("lgStrPrevKeyGlNo")))
		
Call TrimData()
Call FixUNISQLDATA()
Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Function FixUNISQLData()
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim intI
    Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
		    '--------------- ������ coding part(�������,Start)----------------------------------------------------
                   
    UNISqlId(0) = "F5122MA101"
    UNISqlId(1) = "A_GETBIZ"

    Redim UNIValue(1,1)
	
	UNIValue(0,0) = strWhere0
	UNIValue(1,0) = FilterVar(strFrBizCd, "''", "S")
		
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Function QueryData()
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
		    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
    iStr = Split(lgstrRetMsg,gColSep)
		  
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Response.End
    End If
	
	If rs1.EOF And rs1.BOF Then
		If strMsgCd = "" And strBizCd <> "" Then			
			strMsgCd = "970000"
			strMsg1 = Request("txtToBizCd_Alt")
			If strMsgCd <> "" Then
				Call DisplayMsgBox(strMsgCd, vbInformation, strMsg1, strMsg2, I_MKSCRIPT)
				Response.End 
			End If	
		End If
	Else
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtFrBizCd.value = "<%=ConvSPChars(Trim(rs1(0)))%>"			
			.txtFrBizNm.value = "<%=ConvSPChars(Trim(rs1(1)))%>"			
		End With
		</Script>
<%
	End If
	
    If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close:		Set rs0 = Nothing
		Set lgADF = Nothing
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	Else
		Call  MakeSpreadSheetData()
    End If				
	
    Call ReleaseObj()
End Sub
'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()  

	strFromDt	 = UNIConvDate(Request("txtFromDt"))		'���۸����� 
	strToDt		 = UNIConvDate(Request("txtToDt"))			'���Ḹ���� 
	strFrBizCd   = UCase(Request("txtFrBizCd"))             '���ʻ���� 
	strBpCd		 = UCase(Request("txtBpCd"))	            '�ŷ�ó 
	strFrDeptCd  = UCase(Request("txtFrDeptCd"))			'���ʺμ�		 
	strNoteNo    = Request("txtNoteNo")                     '������ȣ 
	'gChangeOrgId
	
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))	

	strWhere0 = ""		
	strWhere0 = strWhere0 &     " JOIN b_biz_area    B  ON A.biz_area_cd = B.biz_area_cd "
    strWhere0 = strWhere0 & " LEFT OUTER JOIN b_biz_partner BP ON A.bp_cd   = BP.bp_cd "
    strWhere0 = strWhere0 & " LEFT OUTER JOIN b_acct_dept   D  ON A.dept_cd = D.dept_cd and A.org_change_id = D.org_change_id "
    strWhere0 = strWhere0 & " LEFT OUTER JOIN b_minor e on a.note_sts=e.minor_cd and e.major_cd= " & FilterVar("F1008","''","S") & " " 
    strWhere0 = strWhere0 & " where A.ISSUE_DT between  " & FilterVar(strFromDt, "''", "S") & " and  " & FilterVar(strToDt, "''", "S") & " "			    
	strWhere0 = strWhere0 & "   and A.biz_area_cd LIKE  " & FilterVar(strFrBizCd, "''", "S") & ""	
	strWhere0 = strWhere0 & "   and A.note_sts in (" & FilterVar("OC", "''", "S") & "," & FilterVar("MV", "''", "S") & ")"
	strWhere0 = strWhere0 & "   and A.note_fg = " & FilterVar("D1", "''", "S") & " "
	
	if strBpCd <> "" Then 		
		strWhere0 = strWhere0 & " and A.bp_cd LIKE  " & FilterVar(strBpCd, "''", "S") & ""
	end if

	if strFrDeptCd <> "" Then 		
		strWhere0 = strWhere0 & " and A.dept_cd LIKE  " & FilterVar(strFrDeptCd, "''", "S") & " and A.org_change_id =  " & FilterVar(gChangeOrgId, "''", "S") & ""
	end if	
	
	If strNoteNo <> "" Then 
		strWhere0 = strWhere0 & " and A.NOTE_NO >= " & Filtervar(strNoteNo, "''", "S")		
	End If

	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL		= " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL	= " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		lgAuthUsrIDAuthSQL		= " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If
	
	strWhere0 = strWhere0 & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL

	If lgStrPrevKeyNoteNo <> "" Then strWhere0 = strWhere0 & " and A.NOTE_NO >= " & Filtervar(lgStrPrevKeyNoteNo, "''", "S")

	strWhere0 = strWhere0 & " Order by A.DEPT_CD, A.NOTE_NO "
	

End Sub

'----------------------------------------------------------------------------------------------------------
' Set MakeSpreadSheetData
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Dim intLoopCnt
%>
<Script Language=vbscript>
Option Explicit

	Dim LngMaxRow       
	Dim strData
	Const C_SHEETMAXROWS_D = 100
	
	LngMaxRow = parent.frm1.vspdData.MaxRows										'Save previous Maxrow                                         	
<%
	
	If rs0.recordcount > GroupCount Then
		intLoopCnt = GroupCount
	Else
		intLoopCnt = rs0.recordcount
	End If
			
	For LngRow = 1 To intLoopCnt

%>		
		strData = strData & Chr(11) & 0
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("DEPT_CD"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("DEPT_NM"))%>"		
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("NOTE_NO"))%>"
		strData = strData & Chr(11) & "<%=UNINumClientFormat(rs0("NOTE_AMT"), ggAmtOfMoney.DecPoint, 0)%>"	
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("NOTE_STS"))%>"	
		strData = strData & Chr(11) & ""  '��� 
		strData = strData & Chr(11) & ""  '�̵��μ� 
		strData = strData & Chr(11) & ""  '�̵��μ��˾� 
		strData = strData & Chr(11) & ""  '�̵��μ���	
'		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_CD"))%>" 
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_NM"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("ISSUE_DT"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("DUE_DT"))%>"		
		strData = strData & Chr(11) & LngMaxRow + <%=LngRow%> 
		strData = strData & Chr(11) & Chr(12)		
<%      
		rs0.MoveNext
    Next
		    
    If Not rs0.EOF Then
%>    
		parent.lgStrPrevKeyNoteNo = "<%=ConvSPChars(rs0("NOTE_NO"))%>"
		parent.lgStrPrevKeyGlNo = ""
<%	Else	%>
		parent.lgStrPrevKeyNoteNo = ""
		parent.lgStrPrevKeyGlNo = ""
<%	End If	%>
		
	With parent
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowData strData
				
		If .frm1.vspdData.MaxRows < C_SHEETMAXROWS_D And .lgStrPrevKeyNoteNo <> "" Then
			.DbQuery					
		Else
			'.frm1.hProcFg.value		= "<%=ConvSPChars(Request("cboProcFg"))%>"
			'.frm1.hNoteFg1.value	= "<%=ConvSPChars(Request("cboNoteFg"))%>"
			'.frm1.hNoteSts.value	= "<%=ConvSPChars(Request("cboNoteSts"))%>"								
			'.frm1.hDueDtStart.value	= "<%=Request("txtDueDtStart")%>"
			'.frm1.hDueDtEnd.value	= "<%=Request("txtDueDtEnd")%>"									
			.frm1.hToBizAreaCd.value = "<%=ConvSPChars(Request("txtFrBizCd"))%>"  ''hidden ���ʻ�����ڵ�		
			.DbQueryOK
		End If
	End With
					
</script>
<%      
End Sub
		
Sub ReleaseObj()			
	Set rs0 = Nothing
	Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
End Sub			
		
%>		
		


