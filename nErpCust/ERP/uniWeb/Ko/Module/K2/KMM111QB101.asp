<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : MM111QA1
'*  4. Program Name         : ��Ƽ���۴ϸ�����ȸ 
'*  5. Program Desc         : ��Ƽ���۴ϸ�����ȸ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/01/14
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Oh Chang Won
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%	
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")
    Dim lgOpModeCRUD
    
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '�� : DBAgent Parameter ���� 
    Dim rs1, rs2, rs3, rs4,rs5
	Dim istrData

	
	Dim iStrIvNo

	Dim StrNextKey		' ���� �� 
	Dim lgStrPrevKey	' ���� �� 
	Dim iLngMaxRow		' ���� �׸����� �ִ�Row
	Dim iLngRow
	Dim GroupCount  
	Dim lgCurrency        
	Dim index,Count     ' ���� �� Return ���� ���� ������ ���� ����     
    Dim lgDataExist
    Dim lgPageNo
    Dim sRow
    Dim lglngHiddenRows
    Dim lgStrPrevKeyM
    DIM MaxRow2
    Dim MaxCount
    Dim iStrVatIncFlg
    
    Dim arrRsVal(11)
	Const C_SHEETMAXROWS_D  = 100
 
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call HideStatusWnd                                                               '��: Hide Processing message
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    lgOpModeCRUD  = Request("txtMode") 

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
             Call  SubBizQueryMulti()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    On Error Resume Next
	lgPageNo       = UNICInt(Trim(Request("lgPageNo1")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgStrPrevKeyM  = UNICInt(Trim(Request("lgStrPrevKeyM")),0)
	lgDataExist    = "No"
	iLngMaxRow     = CLng(Request("txtMaxRows"))
	lgStrPrevKey   = Request("lgStrPrevKey")
	sRow           = CLng(Request("lRow"))
	lglngHiddenRows = CLng(Request("lglngHiddenRows"))

	Call FixUNISQLData()
	Call QueryData()	
	
End Sub    

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query�ϱ� ����  DB Agent �迭�� �̿��Ͽ� Query���� ����� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(0,0)                                                 '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                        '    parameter�� ���� ���� ������ 
    UNISqlId(0) = "MM111QA102" 											' header
 
 	      
    iStrIvNo = Trim(Request("txtIvNo"))
                  
	UNIValue(0,0) = " '"& FilterVar(Trim(UCase(Request("txtIvNo"))), " " , "SNM") & "' "
	
	    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO�� Record Set�̿��Ͽ� Query�� �ϰ� Record Set�� �Ѱܼ� MakeSpreadSheetData()���� Spreadsheet�� �����͸� 
' �Ѹ� 
' ADO ��ü�� �����Ҷ� prjPublic.dll������ �̿��Ѵ�.(�󼼳����� vb�� �ۼ��� prjPublic.dll �ҽ� ����)
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 

    Dim FalsechkFlg
    
    FalsechkFlg = False    

    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
'		Call DisplayMsgBox("172400", vbOKOnly, iStrPrNo, "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.End
    Else    
        Call  MakeSpreadSheetData()
    End If

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr
	Response.Write "	.ggoSpread.Source       = .frm1.vspdData2 "			& vbCr
    Response.Write "	.ggoSpread.SSShowData     """ & istrData	 & """" & vbCr	
    Response.Write "	.lgPageNo1				=  """ & lgPageNo	 & """" & vbCr	
    
    Response.Write "	.lgStrPrevKeyM(" & sRow - 1 & ") = """ & lgStrPrevKeyM & """" & vbCr
    Response.Write "    .lglngHiddenRows(" & sRow - 1 & ") = """ & MaxRow2 & """" & vbCr  
    Response.Write "    .DbQueryOk2(" & MaxCount & ")" & vbCr
    Response.Write "End With"		& vbCr
    Response.Write "</Script>"		& vbCr        

End Sub


'----------------------------------------------------------------------------------------------------------
'QueryData()�� ���ؼ� Query�� �Ǹ� MakeSpreadSheetData()�� ���ؼ� �����͸� ���������Ʈ�� �ѷ��ִ� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100            
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    DIM i

	Const M_IV_DTL_BILL_NO			= 0
	Const M_IV_DTL_BILL_SEQ_NO      = 1
	Const M_IV_DTL_ITEM_CD			= 2
	Const B_ITEM_ITEM_NM            = 3
	Const B_ITEM_SPEC               = 4
	Const M_IV_DTL_BILL_QTY         = 5
	Const M_IV_DTL_BILL_UNIT        = 6
	Const M_IV_DTL_BILL_PRC         = 7
	Const M_IV_DTL_BILL_DOC_AMT     = 8
	Const M_IV_DTL_VAT_DOC_AMT      = 9
	Const M_IV_DTL_VAT_TYPE         = 10
	Const M_IV_DTL_VAT_TYPE_NM      = 11
	Const M_IV_DTL_VAT_RATE         = 12
	Const M_IV_DTL_VAT_INC          = 13

    lgDataExist    = "Yes"
	MaxRow2 = 0	
    iLoopCount = 0
    i = 0
	
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
		MaxRow2     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    

        
	'----- ���ڵ�� Į�� ���� ----------
	'A.BILL_NO, A.BILL_SEQ_NO, A.CUST_ITEM_CD, B.ITEM_NM, B.SPEC, A.BILL_QTY, A.BILL_UNIT, A.BILL_PRC, 
	'A.BILL_DOC_AMT, A.VAT_DOC_AMT, A.VAT_TYPE, (SELECT MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'B9001' AND MINOR_CD = A.VAT_TYPE) VAT_TYPE_NM,
	'A.VAT_RATE, A.VAT_INC
	'-----------------------------------    
	
   Do while Not (rs0.EOF Or rs0.BOF)

        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_BILL_NO))	 									'
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_BILL_SEQ_NO))                                    '
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_ITEM_CD))                                   'ǰ��          

		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(B_ITEM_ITEM_NM))                                                      'ǰ���        
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(B_ITEM_SPEC))                                                         'ǰ��԰�      
		
		
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_IV_DTL_BILL_QTY),ggQty.DecPoint,0)               '����          
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_BILL_UNIT))                                      '����          
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_IV_DTL_BILL_PRC), ggAmtOfMoney.DecPoint,0)       '�ܰ�           
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_IV_DTL_BILL_DOC_AMT), ggAmtOfMoney.DecPoint,0)   '�ݾ�                   
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_IV_DTL_VAT_DOC_AMT), ggAmtOfMoney.DecPoint,0)     '�ΰ����ݾ�             
		
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_VAT_TYPE))                                       '�ΰ�������    
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_VAT_TYPE_NM))                                    '�ΰ���������  
		
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_IV_DTL_VAT_RATE), ggExchRate.DecPoint,6)         '�ΰ�����      
		
'		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_VAT_INC))                                        '�ΰ������Ա��� 
		IF rs0(M_IV_DTL_VAT_INC) = "1" Then
			iStrVatIncFlg = "����"
		Else
			iStrVatIncFlg = "����"			
		End if	
        iRowStr = iRowStr & Chr(11) & ConvSPChars(iStrVatIncFlg)												'�ΰ������Ա��� 
        iRowStr = iRowStr & Chr(11) & sRow
		iRowStr = iRowStr & Chr(11) & Trim(ConvSPChars(MaxRow2 + iLoopCount))
        iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount                             


        If iLoopCount - 1 < C_SHEETMAXROWS_D Then
           istrData = istrData & iRowStr & Chr(11) & Chr(12)
        Else
           'lgStrPrevKeyM = lgStrPrevKeyM + 1
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        rs0.MoveNext
        i = i + 1
   Loop

    If iLoopCount-1 < C_SHEETMAXROWS_D Then                                      '��: Check if next data exists
		lgPageNo = ""
       'lgStrPrevKeyM = ""
    End If

    MaxRow2 = MaxRow2 + iLoopCount 
    MaxCount = iLoopCount
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF

End Sub


'==============================================================================
' Function : SheetFocus
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	
	If Trim(lRow) = "" Then Exit Function
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

%>