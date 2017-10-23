<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%

   Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2     '�� : DBAgent Parameter ���� 
   Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
   Dim lgTailList
   Dim lgSelectList
   Dim lgSelectListDT
   Dim lgDataExist
   Dim lgPageNo
   Dim iTotstrData
   
   Dim strMvmtType
   Dim strItemNm
   Dim iFrPoint
   iFrPoint=0
   
    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")
     
    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgDataExist      = "No"

    Call  FixUNISQLData()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
    call  QueryData()                                                    '�� : DB-Agent�� ���� ADO query

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    Const C_SHEETMAXROWS_D = 100
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = C_SHEETMAXROWS_D * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))

		Next
 
        If iLoopCount < C_SHEETMAXROWS_D Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
           PvArr(iLoopCount) = lgstrData	
			lgstrData = ""
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop
	
	iTotstrData = Join(PvArr, "")

    If iLoopCount < C_SHEETMAXROWS_D Then                                      '��: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF

End Sub
'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
    On Error Resume Next
    
    SetConditionData = FALSE
    
    If Not(rs1.EOF Or rs1.BOF) Then
       strMvmtType =  rs1(1)
    End If   

    Set rs1 = Nothing 

    If Not(rs2.EOF Or rs2.BOF) Then
       strItemNm =  rs2(1)
    End If   

    Set rs2 = Nothing 
	
	SetConditionData = TRUE
	
End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	Dim arrVal(2)
    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(2,2)                                                 '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                        '    parameter�� ���� ���� ������ 
     UNISqlId(0) = "M4111QA001" 										' main query(spread sheet�� �ѷ����� query statement)
     UNISqlId(1) = "M4111QA503"											'�԰����� 
     UNISqlId(2) = "s0000qa001"											'ǰ��� 
     

	'--- 2004-08-20 by Byun Jee Hyun for UNICODE	
     UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 


	If Len(Trim(Request("txtMvmtType"))) Then				'�԰����� 
		strVal = " AND a.IO_TYPE_CD = " & FilterVar(Trim(UCase(Request("txtMvmtType"))), " " , "S") & " " & chr(13)
	End if
	arrVal(0) = FilterVar(Trim(UCase(Request("txtMvmtType"))), " " , "S")

	If Len(Trim(Request("txtMVFrDt"))) Then					'�԰���(����)
		strVal = strVal & " AND A.MVMT_DT >='" & UNIConvDate(Request("txtMVFrDt")) & "'" & chr(13)
	End If

	If Len(Trim(Request("txtMVToDt"))) Then					'�԰���(����)
		strVal = strVal & " AND A.MVMT_DT <='" & UNIConvDate(Request("txtMVToDt")) & "'" & chr(13)
	End If

	If Len(Trim(Request("txtItem"))) Then					'ǰ�� 
		strVal = strVal & " AND A.ITEM_CD = " & FilterVar(Trim(UCase(Request("txtItem"))), " " , "S") & " " & chr(13)
	End If
	arrVal(1) = FilterVar(Trim(UCase(Request("txtItem"))), " " , "S")

	If Len(Trim(Request("txtBeneficiary"))) Then			'������ 
		strVal = strVal & " AND A.BP_CD = " & FilterVar(Trim(UCase(Request("txtBeneficiary"))), " " , "S") & " " & chr(13)
	End If

'	If Len(Trim(Request("txtPayTerms"))) Then				'������� 
'		strVal = strVal & " AND A.PAY_METH ='" & Trim(Request("txtPayTerms")) & "'" & chr(13)
'	End If
	
	If Len(Trim(Request("txtPurGrp"))) Then					'���ű׷� 
		strVal = strVal & " AND D.PUR_GRP = " & FilterVar(Trim(UCase(Request("txtPurGrp"))), " " , "S") & " " & chr(13)
	End If

	If Len(Trim(Request("txtCurrency"))) Then				'ȭ�� 
		strVal = strVal & " AND D.PO_CUR = " & FilterVar(Trim(UCase(Request("txtCurrency"))), " " , "S") & " " & chr(13)
	End If

	If Len(Trim(Request("txtPONo"))) Then					'���ֹ�ȣ 
		strVal = strVal & " AND A.PO_NO = " & FilterVar(Trim(UCase(Request("txtPONo"))), " " , "S") & " " & chr(13)
	End If
	
	'2003.07 TrackingNo �߰� 
    If Len(Request("txtTrackingNo")) Then
		strVal = strVal & " AND A.TRACKING_NO = " & FilterVar(Trim(UCase(Request("txtTrackingNo"))), " " , "S") & "  "		
	End If
	
    UNIValue(0,1) = strVal   
    UNIValue(1,0) = arrVal(0)
    UNIValue(2,0) = arrVal(1)

'================================================================================================================   
   ' UNIValue(0,UBound(UNIValue,2)) = " ORDER BY C.MVMT_NO DESC"
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    Set lgADF   = Nothing
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
	IF SetConditionData() = FALSE THEN EXIT SUB
	 
    If rs0.EOF And rs0.BOF Then
       Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
       rs0.Close
       Set rs0 = Nothing
       Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If
    
End Sub

%>
<Script Language=vbscript>
With parent
    .frm1.txtMvmtTypeNm.Value	= "<%=ConvSPChars(strMvmtType)%>"
    .frm1.txtItemNm.Value		= "<%=ConvSPChars(strItemNm)%>"

    If "<%=lgDataExist%>" = "Yes" Then
       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.frm1.txtHMvmtType.value		= "<%=ConvSPChars(Request("txtMvmtType"))%>"		<%'�԰����� %>
			.frm1.txtHMVFrDt.value		= "<%=Request("txtMVFrDt")%>"		<%'�԰���(����)%>
			.frm1.txtHMVToDt.value		= "<%=Request("txtMVToDt")%>"		<%'�԰���(����)%>
			.frm1.txtHItem.value			= "<%=ConvSPChars(Request("txtItem"))%>"			<%'ǰ�� %>
			.frm1.txtHBeneficiary.value	= "<%=ConvSPChars(Request("txtBeneficiary"))%>"	<%'������ %>
			'.frm1.txtHPayTerms.value		= "<%=ConvSPChars(Request("txtPayTerms"))%>"		<%'������� %>
			.frm1.txtHPurGrp.value		= "<%=ConvSPChars(Request("txtPurGrp"))%>"		<%'���ű׷� %>
			.frm1.txtHCurrency.value		= "<%=ConvSPChars(Request("txtCurrency"))%>"		<%'ȭ�� %>
			.frm1.txtHPONo.value			= "<%=ConvSPChars(Request("txtPONo"))%>"			<%'���ֹ�ȣ %>
       End If
       'Show multi spreadsheet data from this line
       
       .ggoSpread.Source  = .frm1.vspdData
       Parent.frm1.vspdData.Redraw = False
       .ggoSpread.SSShowData "<%=iTotstrData%>","F"          '�� : Display data

		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,"<%=ConvSPChars(Request("txtCurrency"))%>",.GetKeyPos("A",9),"C","I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,"<%=ConvSPChars(Request("txtCurrency"))%>",.GetKeyPos("A",10),"A","I","X","X")

       .lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
 
       .DbQueryOk
       Parent.frm1.vspdData.Redraw = True
    End If   
End With    
</Script>	
