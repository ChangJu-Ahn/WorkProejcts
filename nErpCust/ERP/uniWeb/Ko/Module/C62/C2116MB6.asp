<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

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

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1                             '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgDataExist
Dim lgPageNo
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim lgPlantCd
Dim lgItemCd
Dim lgErrorStatus
Dim lgPrevKey
Dim lgSelectListDT

Dim	DMI_CO		'��������(����)
Dim DMO_CO		'��������(�ܺ�)
Dim IMI_CO		'��������(����)
Dim IMO_CO		'��������(�ܺ�)
Dim DLI_CO		'�����빫��(����)
Dim DLO_CO		'�����빫��(�ܺ�)
Dim ILI_CO		'�����빫��(����)
Dim ILO_CO		'�����빫��(�ܺ�)
Dim DEI_CO		'�������(����)
Dim DEO_CO		'�������(�ܺ�)
Dim IEI_CO		'�������(����)
Dim IEO_CO		'�������(�ܺ�)

Dim lgItemNm			'0
Dim	lgBasicUnit		'1
Dim LngRow
 
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 
	

    lgPageNo       = Trim(Request("lgPageNo"))                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
'    lgMaxCount     = Trim(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgDataExist    = "No"
    lgSelectListDT = Split(Request("lgSelectListDT"),  gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    
    lgPlantCd	   = Trim(Request("txtPlantCd"))
    lgItemCd	   = Trim(Request("txtItemCd"))
	lgPrevKey	   = Trim(Request("lgPrevKey2"))
	LngRow			= UniCInt(Trim(Request("MaxRow")),0)
	
	
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData1(byval iOption)
    Dim  RecordCnt
    Dim  ColCnt

    lgItemNm = ""
    lgBasicUnit = ""

	DMI_CO = 0 : DMO_CO = 0 : IMI_CO = 0 : IMO_CO = 0 : DLI_CO = 0 : DLO_CO = 0
	ILI_CO = 0 : ILO_CO = 0 : DEI_CO = 0 : DEO_CO = 0 : IEI_CO = 0 : IEO_CO = 0
    
    
    IF iOption = 1 Then
		
		
		lgItemNm = Trim(rs0(0))
		lgBasicUnit = Trim(rs0(1))
		
	    Do while Not (rs0.EOF Or rs0.BOF)
			IF UCase(Trim(rs0(2))) = "D" AND UCase(Trim(rs0(3))) = "M" Then
				DMI_CO = DMI_CO + CDbl(rs0(4))
				DMO_CO = DMO_CO + CDbl(rs0(5))
			ELSEIF UCase(Trim(rs0(2))) = "I" AND UCase(Trim(rs0(3))) = "M" Then
				IMI_CO = IMI_CO + CDbl(rs0(4))
				IMO_CO = IMO_CO + CDbl(rs0(5))
			ELSEIF UCase(Trim(rs0(2))) = "D" AND UCase(Trim(rs0(3))) = "L" Then
				DLI_CO = DLI_CO + CDbl(rs0(4))
				DLO_CO = DLO_CO + CDbl(rs0(5))
			ELSEIF UCase(Trim(rs0(2))) = "I" AND UCase(Trim(rs0(3))) = "L" Then
				ILI_CO = ILI_CO + CDbl(rs0(4))
				ILO_CO = ILO_CO + CDbl(rs0(5))
			ELSEIF UCase(Trim(rs0(2))) = "D" AND UCase(Trim(rs0(3))) = "E" Then
				DEI_CO = DEI_CO + CDbl(rs0(4))
				DEO_CO = DEO_CO + CDbl(rs0(5))
			ELSE
				IEI_CO = IEI_CO + CDbl(rs0(4))
				IEO_CO = IEO_CO + CDbl(rs0(5))
			END IF
			rs0.MoveNext
		Loop

		rs0.Close
		Set rs0 = Nothing     
    END IF
End Sub

Sub MakeSpreadSheetData2()
    Dim  iLoopCount, iLoopCount2
    lgstrData = ""

    lgDataExist    = "Yes"

    IF lgPrevKey <> "" Then
		Do while Not (rs1.EOF Or rs1.BOF)
			 IF Trim(rs1(0)) = lgPrevKey  Then
				Exit Do	
			 END IF
		     rs1.MoveNext
		Loop
	END IF
	
    Const C_SHEETMAXROWS_D  = 100 
	
	lgMaxCount = CInt(C_SHEETMAXROWS_D)       
    
    iLoopCount = -1
    iLoopCount2 = 0
    lgstrData = ""
    
    Do while Not (rs1.EOF Or rs1.BOF)

		If  iLoopCount < UniConvNumStringToDouble(lgMaxCount,0)  Then
			lgstrData = lgstrData & Chr(11) & ConvSPChars(FormatRsString(lgSelectListDT(0),Trim(rs1(0))))			'ǰ���ڵ� 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(FormatRsString(lgSelectListDT(1),Trim(rs1(1))))		'ǰ����� 
			lgstrData = lgstrData & Chr(11) & FormatRsString(lgSelectListDT(2),Trim(rs1(2)))	'����(����)
			lgstrData = lgstrData & Chr(11) & iLoopCount2 + LngRow + 1
			lgstrData = lgstrData & Chr(11) & Chr(12)
		Else
		    lgPrevKey = Trim(rs1(0))
		    Exit Do
		END IF
       rs1.MoveNext
       iLoopCount = iLoopCount + 1
	   iLoopCount2 = iLoopCount2 +1
	Loop


    If  iLoopCount < UniConvNumStringToDouble(lgMaxCount,0) Then                                            '��: Check if next data exists
        lgPrevKey = ""
    End If
  	
  	
	rs1.Close
    Set rs1 = Nothing 
    
    
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(1,1)

    UNISqlId(0) = "C2110MA102"
    UNISqlId(1) = "C2110MA103"
    
    UNIValue(0,0) = FilterVar(lgPlantCd, "''", "S")				'�����ڵ� 
    UNIValue(0,1) = FilterVar(lgItemCd, "''", "S")				'ǰ���ڵ� 
    UNIValue(1,0) = FilterVar(lgPlantCd, "''", "S")				'�����ڵ� 
    UNIValue(1,1) = FilterVar(lgItemCd, "''", "S")				'ǰ���ڵ� 
    

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 

	IF lgErrorStatus = "YES" Then
		Exit Sub
	END IF
	
	
	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
		
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    'If iStr(0) <> "0" Then
    '    Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    'End If    
        
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Call  MakeSpreadSheetData1(0)
    Else 
		
		Call  MakeSpreadSheetData1(1)
    End If

    
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
    Else 
		Call  MakeSpreadSheetData2()
    End If

	
	
End Sub





'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '��: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>

<Script Language=vbscript>
    
With Parent
    
    
       'Set condition data to hidden area
		.frm1.txtDi_Mcost.text  =  "<%=UNINumClientFormat(DMI_CO+DMO_CO, ggUnitCost.DecPoint, 0)%>"
		.frm1.txtDi_Lcost.text =   "<%=UNINumClientFormat(DLI_CO+DLO_CO, ggUnitCost.DecPoint, 0)%>"
		.frm1.txtDi_Ecost.text =   "<%=UNINumClientFormat(DEI_CO+DEO_CO, ggUnitCost.DecPoint, 0)%>"
		
		.frm1.txtInd_Mcost.text =  "<%=UNINumClientFormat(IMI_CO+IMO_CO, ggUnitCost.DecPoint, 0)%>"
		.frm1.txtInd_Lcost.text =  "<%=UNINumClientFormat(ILI_CO+ILO_CO, ggUnitCost.DecPoint, 0)%>"
		.frm1.txtInd_Ecost.text =  "<%=UNINumClientFormat(IEI_CO+IEO_CO, ggUnitCost.DecPoint, 0)%>"
		
			'���ο��� ������ �� 
		.frm1.txtInDi_Sum.text =  "<%=UNINumClientFormat(DMI_CO+DLI_CO+DEI_CO, ggUnitCost.DecPoint, 0)%>"
			'���ο��� ������ �� 
		.frm1.txtInInd_Sum.text = "<%=UNINumClientFormat(IMI_CO+ILI_CO+IEI_CO, ggUnitCost.DecPoint, 0)%>"

		   'ǰ��� 
		.frm1.txtItemNmDesc.value = "<%=ConvSPChars(lgItemNm)%>"
			'�԰� 
		.frm1.txtItemUnt.value = "<%=ConvSPChars(lgBasicUnit)%>"
		
   		'�ܺο��� ���� �� 
		.frm1.txtOutDi_Sum.text =  "<%=UNINumClientFormat(DMO_CO+DLO_CO+DEO_CO, ggUnitCost.DecPoint, 0)%>"
   		'�ܺο��� ���� �� 
		.frm1.txtOutInd_Sum.text =  "<%=UNINumClientFormat(IMO_CO+ILO_CO+IEO_CO, ggUnitCost.DecPoint, 0)%>"
		
		
		'������ �� 
		.frm1.txtDi_Sum.text   = "<%=UNINumClientFormat(DMI_CO+DMO_CO+DLI_CO+DLO_CO+DEI_CO+DEO_CO, ggUnitCost.DecPoint, 0)%>"
		'������ �� 
		.frm1.txtInd_Sum.text  = "<%=UNINumClientFormat(IMI_CO+IMO_CO+ILI_CO+ILO_Co+IEI_CO+IEO_CO, ggUnitCost.DecPoint, 0)%>"
       
    
    If "<%=lgDataExist%>" = "Yes" Then    
       .ggoSpread.Source  = .frm1.vspdData2
       .ggoSpread.SSShowData "<%=lgstrData%>"            '�� : Display data
       .lgPrevKey2      =  "<%=lgPrevKey%>"               '�� : Next next data tag
       .DbQueryOk()
    End If   
END WITH
</Script>	
