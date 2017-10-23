//======================================================================================================
//	INC Name	: ToolbarScript.js
//	Description	: ToolBar 속성 전환시 사용 
//======================================================================================================
// 켜진 상태 

	tbExplorern = new Image;
	tbExplorern.src = "../CShared/image/tQuery_on.gif";

	tbQueryn = new Image;
	tbQueryn.src = "../CShared/image/tQuery_on.gif";

	tbNewn = new Image;
	tbNewn.src = "../CShared/image/tNew_on.gif";

	tbDeleten = new Image;
	tbDeleten.src = "../CShared/image/tDelete_on.gif";

	tbInsertRown = new Image;
	tbInsertRown.src = "../CShared/image/tInsertRow_on.gif";

	tbDeleteRown = new Image;
	tbDeleteRown.src = "../CShared/image/tDeleteRow_on.gif";

	tbSaven = new Image;
	tbSaven.src = "../CShared/image/tSave_on.gif";

	tbPrevn = new Image;
	tbPrevn.src = "../CShared/image/tPrev_on.gif";

	tbNextn = new Image;
	tbNextn.src = "../CShared/image/tNext_on.gif";

	tbCopyn = new Image;
	tbCopyn.src = "../CShared/image/tCopy_on.gif";

	tbCanceln = new Image;
	tbCanceln.src = "../CShared/image/tCancel_on.gif";

	tbExceln = new Image;
	tbExceln.src = "../CShared/image/tExcel_on.gif";

	tbPrintn = new Image;
	tbPrintn.src = "../CShared/image/tPrint_on.gif";

	tbExitn = new Image;
	tbExitn.src = "../CShared/image/tExit_on.gif";

	tbHelpn = new Image;
	tbHelpn.src = "../CShared/image/tHelp_on.gif";

	tbFindn = new Image;
	tbFindn.src = "../CShared/image/tFind_on.gif";

	//tbSaveGridn = new Image;
	//tbSaveGridn.src = "../../CShared/image/tool16_on.gif";

// 그래이된 상태 by Shin hyoung jae 2001/3/21
	tbExplorerg = new Image;
	tbExplorerg.src = "../CShared/image/tQuery_gr.gif";

	tbQueryg = new Image;
	tbQueryg.src = "../CShared/image/tQuery_gr.gif";

	tbNewg = new Image;
	tbNewg.src = "../CShared/image/tNew_gr.gif";

	tbDeleteg = new Image;
	tbDeleteg.src = "../CShared/image/tDelete_gr.gif";

	tbInsertRowg = new Image;
	tbInsertRowg.src = "../CShared/image/tInsertRow_gr.gif";

	tbDeleteRowg = new Image;
	tbDeleteRowg.src = "../CShared/image/tDeleteRow_gr.gif";

	tbSaveg = new Image;
	tbSaveg.src = "../CShared/image/tSave_gr.gif";

	tbPrevg = new Image;
	tbPrevg.src = "../CShared/image/tPrev_gr.gif";

	tbNextg = new Image;
	tbNextg.src = "../CShared/image/tNext_gr.gif";

	tbCopyg = new Image;
	tbCopyg.src = "../CShared/image/tCopy_gr.gif";

	tbCancelg = new Image;
	tbCancelg.src = "../CShared/image/tCancel_gr.gif";

	tbExcelg = new Image;
	tbExcelg.src = "../CShared/image/tExcel_gr.gif";

	tbPrintg = new Image;
	tbPrintg.src = "../CShared/image/tPrint_gr.gif";

	tbExitg = new Image;
	tbExitg.src = "../CShared/image/tExit_gr.gif";

	tbHelpg = new Image;
	tbHelpg.src = "../CShared/image/tHelp_gr.gif";

	tbFindg = new Image;
	tbFindg.src = "../CShared/image/tFind_gr.gif";


	// 커진 상태 
	tbExplorerd = new Image;
	tbExplorerd.src = "../CShared/image/tQuery_off.gif";

	tbQueryd = new Image;
	tbQueryd.src = "../CShared/image/tQuery_off.gif";

	tbNewd = new Image;
	tbNewd.src = "../CShared/image/tNew_off.gif";

	tbDeleted = new Image;
	tbDeleted.src = "../CShared/image/tDelete_off.gif";

	tbInsertRowd = new Image;
	tbInsertRowd.src = "../CShared/image/tInsertRow_off.gif";

	tbDeleteRowd = new Image;
	tbDeleteRowd.src = "../CShared/image/tDeleteRow_off.gif";

	tbSaved = new Image;
	tbSaved.src = "../CShared/image/tSave_off.gif";

	tbPrevd = new Image;
	tbPrevd.src = "../CShared/image/tPrev_off.gif";

	tbNextd = new Image;
	tbNextd.src = "../CShared/image/tNext_off.gif";

	tbCopyd = new Image;
	tbCopyd.src = "../CShared/image/tCopy_off.gif";

	tbCanceld = new Image;
	tbCanceld.src = "../CShared/image/tCancel_off.gif";

	tbExceld = new Image;
	tbExceld.src = "../CShared/image/tExcel_off.gif";

	tbPrintd = new Image;
	tbPrintd.src = "../CShared/image/tPrint_off.gif";

	tbExitd = new Image;
	tbExitd.src = "../CShared/image/tExit_off.gif";

	tbHelpd = new Image;
	tbHelpd.src = "../CShared/image/tHelp_off.gif";

	//tbSaveGridd = new Image;
	//tbSaveGridd.src = "../../CShared/image/tool16_off.gif";

	tbFindd = new Image;
	tbFindd.src = "../CShared/image/tFind_off.gif";


	function ChgEnabImg(imgName) {			//Enable Image로 전환 
		imgOff = eval(imgName + "n.src");
		document[imgName].src = imgOff;
	}

	function ChgDisImg(imgName) {			//Disable Image로 전환 
		imgDis = eval(imgName + "d.src");
		document[imgName].src = imgDis;
	}

	function ChgGryImg(imgName) {			//Gray Image로 전환 by Shin hyoung jae 2001/3/21
		imgGry = eval(imgName + "g.src");
		document[imgName].src = imgGry;
	}
