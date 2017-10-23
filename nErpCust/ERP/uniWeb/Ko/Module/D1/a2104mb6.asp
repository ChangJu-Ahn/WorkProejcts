<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<% 
	Call LoadBasisGlobalInf()

	On Error Resume Next
	Err.Clear

	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1                       '�� : DBAgent Parameter ���� 
	Dim lgstrData                                                              '�� : data for spreadsheet data
	Dim lgStrPrevKey                                                           '�� : ���� �� 
	Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
	Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo

	Dim lgClassType
	Dim	lgClassNm
	Dim txtClassType


  
	Call HideStatusWnd 

	lgDataExist    = "No"

	txtClassType	= Trim(Request("txtClassType"))

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

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To 18
            iRowStr = iRowStr & Chr(11) & RTrim(rs0(ColCnt))
		Next
 
        lgstrData      = lgstrData  & iRowStr & Chr(11) & iLoopCount + 1 & Chr(11) & Chr(12)

        rs0.MoveNext
	Loop

	rs0.Close
    Set rs0 = Nothing 

    If Not( rs1.EOF OR rs1.BOF) Then
   		lgClassType = rs1(0)
		lgClassNm = rs1(1)
    End IF
    rs1.Close
    Set rs1= Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(1,1)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    UNISqlId(0) = "A2104QA101"
    UNISqlId(1) = "A_CLSTYPE"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = FilterVar(txtClassType, "''", "S")                                          '��: Select list
    UNIValue(0,1) = FilterVar(txtClassType, "''", "S") 
    UNIValue(1,0) = FilterVar(txtClassType, "''", "S")     
    '--------------- ������ coding part(�������,End)------------------------------------------------------
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

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If

    If  rs0.EOF And rs0.BOF Then

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
    If "<%=lgDataExist%>" = "Yes" Then
		With Parent
		.Frm1.txtClassTypeNm.Value		  = "<%=ConvSPChars(lgClassNm)%>" 
        .ggoSpread.Source  = .frm1.vspdData
        .ggoSpread.SSShowData "<%=ConvSPChars(lgstrData)%>"
        .DbQueryOk
       End With
    End If
</Script>

