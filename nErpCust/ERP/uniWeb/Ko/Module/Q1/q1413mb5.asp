<%@LANGUAGE = VBScript%> 
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->	
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1413MB5
'*  4. Program Name         : �跮 ������ n,k���� 
'*  5. Program Desc         : 
'*  6. Component List       : PQBG120
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/07/30
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Koh Jae Woo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- ChartFX�� ����� ����ϱ� ���� Include ���� -->
<!-- #include file="../../inc/CfxIE.inc" -->
<%													
On Error Resume Next

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("I", "*", "NOCOOKIE", "QB")

Dim strP1
Dim strP2
Dim strAlpha
Dim strBeta
	 
Dim strSTDack
Dim strInsCri

Dim strUpperBound
Dim strLowerBound

Dim VarSamlpesize
Dim VarAcceptValue
'/* Issue: ��� ���� �ȳ��� - �������� �߸��� - START */
Dim Z_alpha
'/* Issue: ��� ���� �ȳ��� - �������� �߸��� - END */
Dim Z_beta
Dim Z_p1
Dim Z_p2

Dim sample
Dim k
Dim k1
Dim k2
  
strP1 = Request("txtP1")
strP2 = Request("txtP2")
strAlpha = Request("txtAlpha")
strBeta = Request("txtBeta")

strSTDack = Request("txtSTDack")
strInsCri = Request("txtInsCri")

strUpperBound = Request("txtUpperBound")
strLowerBound = Request("txtLowerBound")

Z_alpha = Z(CDbl(UNIConvNum(strAlpha, 0) * 0.01))
Z_beta = Z(CDbl(UNIConvNum(strBeta, 0) * 0.01))
Z_p1 = Z(CDbl(UNIConvNum(strP1, 0) * 0.01))
Z_p2 = Z(CDbl(UNIConvNum(strP2, 0) * 0.01))

If strSTDack = "0" Then	
	sample = ((Z_alpha + Z_beta) / (Z_p1 - Z_p2)) ^  2
	
	VarSamlpesize = Cint(sample)
		
	k1 = - Z_p1 + ( Z_alpha  / sqr(VarSamlpesize))
	k2 = - Z_p2 - (Z_beta  / sqr(VarSamlpesize))
	k = (k1 + k2) / 2 
	k = Round(k, 3)
Else
	k = -((Z_alpha * Z_p2) + (Z_beta * Z_p1)) / (Z_alpha + Z_beta)
	k = Round(k, 3)

	sample = (1 + (( k ^ 2) / 2)) * ((Z_alpha + Z_beta) / (Z_p1 - Z_p2 )) ^ 2 
	VarSamlpesize = Cint(sample)
End If
%>
<Script Language=vbscript>
With Parent.frm1
	.txtSampleSize.Text = "<%=UniNumClientFormat(VarSamlpesize, ggQty.DecPoint ,0)%>"
	.txtAcceptSize.Text = "<%=UniNumClientFormat(k, 4 ,0)%>"
End with
</Script>	

<%
	'/* Issue: ���Ժ��� Ȯ���� ���ϴ� ���� ���� - START */
	'++++++++++++++++++++++++++++++++++++++++++  2.5.1 ���Ժ��� ����Լ� +++++++++++++++++++++++++++++++++++++++
	'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	Const PI = 3.141592653589793		'3.14     '3.1415926535897932384626433832D
    Const SQRT_2PI_INVERSE = 0.398942280401433		
	
    Const P  = 0.2316419
    Const C1  = 0.31938153
    Const C2  = -0.356563782
    Const C3  = 1.78147937
    Const C4  = -1.821255978
    Const C5  = 1.330274429
  
	'************************************************************************************************************************
    '                                                             ����Ȯ�� Q(u)
    '       Q(u) = 1 - ��(u)
    '       ��(u) = �� ��(u) du
    '       ��(u) = 1/ ��(2* PI) * Exp(-u^2 /2)
    '************************************************************************************************************************
    Private Function Q(ByVal pvz)
        Dim iDblQ
        Dim iDecPSI

        iDecPSI = SQRT_2PI_INVERSE * Exp(-(pvz ^ 2) / 2)
		
		If Abs(pvz) <= 2.2 Then
            iDblQ = ShentonExpansion(pvz, iDecPSI)
        Else
            iDblQ = LaplaceExpansion(pvz, iDecPSI)
        End If
		
		If pvz < 0 Then
            iDblQ = 1 - iDblQ
        End If
		
		Q = iDblQ
    End Function

    'Shenton ���м� ������ 
    Private Function ShentonExpansion(ByVal pvz , ByVal pviDecPSI) 
        Dim iDblQ
        Dim ABS_Z

        iDblQ = 0
        ABS_Z = Abs(pvz)
        Dim k 

        For k = 12 To 1 Step -1
            iDblQ = k * (pvz ^ 2) / ((2 * k) + 1 + (((-1) ^ (k - 1)) * iDblQ))
        Next

        iDblQ = (1 / 2) - ((pviDecPSI * ABS_Z) / (1 - iDblQ))

        ShentonExpansion = iDblQ

    End Function

    'Laplace ���м� ������ 
    Private Function LaplaceExpansion(ByVal pvz, ByVal pviDecPSI) 
        Dim iDblQ 
        Dim ABS_Z 
        Dim k 

        iDblQ = 0
        ABS_Z = Abs(pvz)

        For k = 12 To 1 Step -1
            iDblQ = k / (ABS_Z + iDblQ)
        Next

        iDblQ = (pviDecPSI * 1) / (ABS_Z + iDblQ)

        LaplaceExpansion = iDblQ
    End Function

 
	'************************************************************************************************************************
    '                                                             ���Ժ��� ����Ȯ���� 
    '       Q(z) = 1 - ��(z)
    '       ��(z) = �� ��(z) du
    '       ��(z) = 1/ ��(2* PI) * Exp(-z^2 /2)
    '************************************************************************************************************************
    Function z(ByVal pvQ )

        Dim iDecQ 
        Dim iDecPSI
        Dim iDecPSI_Prime
        Dim k 
        Dim y 
        Dim iDblu_2Q 
        Dim Q_for_Distinction 
        Dim Value_for_Distinction_Formula 
        Dim Distinction_Ratio 
		
		'1���� ū ����Ȯ������ ���� �� ����, 1�� ���� -���Ѵ뿡�� +���Ѵ��� ���� ����Ȯ�� ���̴� 
        If pvQ >= 1 Then
            Exit Function
        End If
		
		If pvQ = 0.5 Then
            z = 0
        End If
        
        If pvQ < 0.5 Then
            iDecQ = pvQ
        Else
            iDecQ = 1 - pvQ
        End If
		
		y = -Log(4 * iDecQ * (1 - iDecQ))
		
        iDblu_2Q = Sqr(y * (2.0611786 - 5.7262204 / (y + 11.640595)))
        
        k = 0
		
        Do
            iDecPSI = SQRT_2PI_INVERSE * Exp(-iDblu_2Q ^ 2 / 2)
            iDecPSI_Prime = -iDblu_2Q * iDecPSI

            Q_for_Distinction = Q(iDblu_2Q)
           
            Value_for_Distinction_Formula = iDecPSI ^ 2 - (iDecQ - Q_for_Distinction) * iDecPSI_Prime * 2

            If Value_for_Distinction_Formula > 0 Then
                Distinction_Ratio = (iDecQ - Q_for_Distinction) * 2 / (-iDecPSI - Sqr(Value_for_Distinction_Formula))
            Else
                Distinction_Ratio = -iDecPSI / iDecPSI_Prime
            End If

            iDblu_2Q = iDblu_2Q + Distinction_Ratio

            k = k + 1
			
        Loop While (k < 10 And Abs(Distinction_Ratio) > 10 ^ (-4))
		
		If pvQ > 0.5 Then
            iDblu_2Q = -iDblu_2Q
        End If
		
		'z = iDblu_2Q		'�� ���α׷������� �Ʒ��� ���� ��� 
        z = -iDblu_2Q

    End Function
	
	'/* Issue: ���Ժ��� Ȯ���� ���ϴ� ���� ���� - END */    
%>
