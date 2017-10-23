<%Option Explicit%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->


<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")       

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1                         '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim lgPlantnm
Dim lgWorkStepNm
Dim lgItemAcctNm
Dim lgItemGroupNm
Dim lgItemNm
Dim lgMFCSum
Dim lgMATSum
Dim lgSEMISum
Dim lgSum

'--------------- ������ coding part(��������,Start)--------------------------------------------------------

'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 


    lgPageNo       = Trim(Request("lgPageNo"))                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = Trim(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"
	lgSum = 0
	

    
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    lgDataExist    = "Yes"
    lgstrData      = ""

    If UniConvNumStringToDouble(lgPageNo,0) > 0 Then
       rs0.Move     = UniConvNumStringToDouble(lgMaxCount,0) * UniConvNumStringToDouble(lgPageNo,0)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < UniConvNumStringToDouble(lgMaxCount,0) Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = Cstr(UniConvNumStringToDouble(lgPageNo,0) + 1)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < UniConvNumStringToDouble(lgMaxCount,0) Then                                            '��: Check if next data exists
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
    Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(1,7)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

	' ��ȿ�� üũ 

  	
    IF Trim(Request("txtPlantCd")) <> "" Then 
		Call CommonQueryRs("PLANT_NM","B_PLANT","PLANT_CD = " & FilterVar(Request("txtPlantCd"), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		If Trim(Replace(lgF0,Chr(11),"")) = "X" then
		Else
			lgPlantNm = Trim(Replace(lgF0,Chr(11),""))
		End if
	END IF
	
  	Call CommonQueryRs("A.MINOR_NM","B_MINOR A,B_CONFIGURATION B","A.MAJOR_CD = " & FilterVar("C2000", "''", "S") & "  AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = " & FilterVar(Request("txtWorkStepCd"), "''", "S")  & _
  				" AND A.MINOR_CD = B.MINOR_CD and  B.SEQ_NO = 4 and B.REFERENCE = " & FilterVar("Y", "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
	If Trim(Replace(lgF0,Chr(11),"")) = "X" then
	Else
		lgWorkStepNm = Trim(Replace(lgF0,Chr(11),""))
	End if
	

    IF Trim(Request("txtItemAcctCd")) <> "" Then 
		Call CommonQueryRs("MINOR_NM","B_MINOR","MAJOR_CD =" & FilterVar("P1001", "''", "S") & "  AND MINOR_CD = " & FilterVar(Request("txtItemAcctCd"), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		If Trim(Replace(lgF0,Chr(11),"")) = "X" then
		Else
			lgItemAcctNm = Trim(Replace(lgF0,Chr(11),""))
		End if
	END IF	

    IF Trim(Request("txtItemGroupCd")) <> "" Then 
		Call CommonQueryRs("ITEM_GROUP_NM","B_ITEM_GROUP","ITEM_GROUP_CD = " & FilterVar(Request("txtItemGroupCd"), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		If Trim(Replace(lgF0,Chr(11),"")) = "X" then

		Else
			lgItemGroupNm = Trim(Replace(lgF0,Chr(11),""))
		End if
	END IF	
	
	IF Trim(Request("txtItemCd")) <> "" Then 
  		Call CommonQueryRs("A.ITEM_NM","B_ITEM A,B_ITEM_BY_PLANT B","B.PLANT_CD = " & FilterVar(Request("txtPlantCd"), "''", "S") & _
  		 " AND B.ITEM_CD = " & FilterVar(Request("txtItemCd"), "''", "S") & " AND  A.ITEM_CD = B.ITEM_CD" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		If Trim(Replace(lgF0,Chr(11),"")) = "X" then
		Else
			lgItemNm = Replace(Trim(Replace(lgF0,Chr(11),"")),"""","")
		End if
	END IF

	
    UNISqlId(0) = "C3603MA101"
	UNISqlId(1) = "C3603MA102"
	
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	UNIValue(0,1)  = FilterVar(Request("txtYYYYMM"), "''", "S")               
	UNIValue(1,0)  = FilterVar(Request("txtYYYYMM"), "''", "S")               
	
	
	IF Trim(Request("txtPlantCd")) <> "" Then
		UNIValue(0,2)  = FilterVar(Trim(Request("txtPlantCd"))   , "''", "S")
		UNIValue(1,1)  = FilterVar(Trim(Request("txtPlantCd"))   , "''", "S")
	ELSE
	    UNIValue(0,2)  = "" & FilterVar("%", "''", "S") & ""
	    UNIValue(1,1)  = "" & FilterVar("%", "''", "S") & ""
	END IF

	IF Trim(Request("txtWorkStepCd")) <> "" Then
		UNIValue(0,3)  = FilterVar(Trim(Request("txtWorkStepCd"))   , "''", "S")
		UNIValue(1,2)  = FilterVar(Trim(Request("txtWorkStepCd"))   , "''", "S")
	ELSE
	    UNIValue(0,3)  = "" & FilterVar("%", "''", "S") & ""
	    UNIValue(1,2)  = "" & FilterVar("%", "''", "S") & ""
	END IF

	IF Trim(Request("txtItemAcctCd")) <> "" Then
		UNIValue(0,4)  = FilterVar(Trim(Request("txtItemAcctCd"))   , "''", "S")
		UNIValue(1,3)  = FilterVar(Trim(Request("txtItemAcctCd"))   , "''", "S")
	ELSE
	    UNIValue(0,4)  = "" & FilterVar("%", "''", "S") & ""
	    UNIValue(1,3)  = "" & FilterVar("%", "''", "S") & ""
	END IF

	IF Trim(Request("txtItemGroupCd")) <> "" Then
		UNIValue(0,5)  = FilterVar(Trim(Request("txtItemGroupCd"))   , "''", "S")
		UNIValue(1,4)  = FilterVar(Trim(Request("txtItemGroupCd"))   , "''", "S")
	ELSE
	    UNIValue(0,5)  = "" & FilterVar("%", "''", "S") & ""
	    UNIValue(1,4)  = "" & FilterVar("%", "''", "S") & ""
	END IF
	
	IF Trim(Request("txtItemCd")) <> "" Then
		UNIValue(0,6)  = FilterVar(Trim(Request("txtItemCd"))   , "''", "S")
		UNIValue(1,5)  = FilterVar(Trim(Request("txtItemCd"))   , "''", "S")
	ELSE
	    UNIValue(0,6)  = "" & FilterVar("%", "''", "S") & ""
	    UNIValue(1,5)  = "" & FilterVar("%", "''", "S") & ""
	END IF
	

	
	
		                 
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
    Dim iStr
    Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    'If iStr(0) <> "0" Then
    '    Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    'End If    
        
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("234600", vbOKOnly, "", "", I_MKSCRIPT)		'�۾��ܰ躰 ���� ����  �����Ͱ� �������� �ʾ� ���忡 �����߽��ϴ�.
        rs0.Close
        rs1.Close
        Set rs0 = Nothing
        Set rs1 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If

	lgMFCSum = rs1(0)
	lgMATSum = rs1(1)
	lgSEMISum = rs1(2)
    lgSum = rs1(3)
    rs1.Close
    Set rs1 = Nothing
End Sub

%>

<Script Language=vbscript>
    With Parent
		.Frm1.txtPlantNm.Value = "<%=ConvSPChars(lgPlantNm)%>"
		.Frm1.txtWorkStepNm.Value = "<%=ConvSPChars(lgWorkStepNm)%>"
		.Frm1.txtItemAcctNm.Value = "<%=ConvSPChars(lgItemAcctNm)%>"
		.Frm1.txtItemGroupNm.Value = "<%=ConvSPChars(lgItemGroupNm)%>"
		.Frm1.txtItemNm.Value = "<%=ConvSPChars(lgItemNm)%>"

    If "<%=lgDataExist%>" = "Yes" Then

    
       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" or "<%=lgPageNo%>"= "" Then   ' "1" means that this query is first and next data exists
			.frm1.hYyyymm.value = "<%=Request("txtYYYYMM")%>"
			.frm1.hWorkStepCd.value = "<%=Request("txtWorkStepCd")%>"
			.frm1.hPlantCd.value = "<%=Request("txtPlantCd")%>"
			.frm1.hItemAcctCd.value = "<%=Request("txtItemAcctCd")%>"
			.frm1.hItemGroupCd.value = "<%=Request("txtItemGroupCd")%>"
			.frm1.hItemCd.value = "<%=Request("txtItemCd")%>"
			.Frm1.txtMFCSum.text = "<%=UniNumClientFormat(lgMFCSum,ggAmtofMoney.Decpoint,0)%>" 
			.Frm1.txtMATSum.text = "<%=UniNumClientFormat(lgMATSum,ggAmtofMoney.Decpoint,0)%>" 
			.Frm1.txtSEMISum.text = "<%=UniNumClientFormat(lgSEMISum,ggAmtofMoney.Decpoint,0)%>" 
			.Frm1.txtSum.text = "<%=UniNumClientFormat(lgSum,ggAmtofMoney.Decpoint,0)%>" 
       End If
       
       'Show multi spreadsheet data from this line
       
			.ggoSpread.Source  = Parent.frm1.vspdData
			.ggoSpread.SSShowData "<%=lgstrData%>"                  '�� : Display data
			.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
			.DbQueryOk
    End If   
    End With
</Script>	

