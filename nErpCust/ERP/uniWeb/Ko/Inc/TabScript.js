var gImgFolder = ""

function ChangeTabs(pPageNo) {

	var panel = TabDiv	//document.all.tags("div")	// Div 를 Array로 선언 
	var myTabs = MyTab

	if (gImgFolder == "") {
		var strLoc = myTabs[pPageNo-1].rows[0].cells(0).background
		var iLoc = 1
		var strTemp
		
		iLoc = strLoc.lastIndexOf("/");
		
		gImgFolder = strLoc.substr(0, iLoc+1);
		
	}

	// "../../image/table/tab_up_bg.gif"
	
	for ( var i=0; i < myTabs.length; i++){
		if (i != pPageNo - 1) {
//			myTabs[i].rows[0].cells(0).background = gImgFolder + "tab_up_bg.gif"; //IE 진행바 에러를 없애기 위해서 불필요한 코드를 막음 
			myTabs[i].rows[0].cells(1).background = gImgFolder + "tab_up_bg.gif";
//			myTabs[i].rows[0].cells(2).background = gImgFolder + "tab_up_bg.gif"; //IE 진행바 에러를 없애기 위해서 불필요한 코드를 막음 

			myTabs[i].rows[0].cells(0).children(0).src = gImgFolder + "tab_up_left.gif";
			myTabs[i].rows[0].cells(2).children(0).src = gImgFolder + "tab_up_right.gif";
			panel(i).style.display = "none";
		}
		
	}

	// 각각의 Tab 속성을 Default, Display None으로 설정 
//	  myTabs[pPageNo-1].rows[0].cells(0).background = gImgFolder + "seltab_up_bg.gif"; //IE 진행바 에러를 없애기 위해서 불필요한 코드를 막음 
 	  myTabs[pPageNo-1].rows[0].cells(1).background = gImgFolder + "seltab_up_bg.gif";
//	  myTabs[pPageNo-1].rows[0].cells(2).background = gImgFolder + "seltab_up_bg.gif"; //IE 진행바 에러를 없애기 위해서 불필요한 코드를 막음 
 
	  myTabs[pPageNo-1].rows[0].cells(0).children(0).src = gImgFolder + "seltab_up_left.gif";
	  myTabs[pPageNo-1].rows[0].cells(2).children(0).src = gImgFolder + "seltab_up_right.gif";
	  panel(pPageNo-1).style.display = "";
	// 해당 Tab의 속성을 Enable로 설정 
       gPageNo     = pPageNo ;

}

function ResizeTabs() {
	var panel = document.all.tags("div")	// Div 를 Array로 선언 
	var myTabs = MyTab

	for (var i=0; i < panel.length; i++) {
		myTabs[i].parentElement.width = eval(myTabs[i].rows[0].cells(0).offsetWidth) + eval(myTabs[i].rows[0].cells(1).offsetWidth) + eval(myTabs[i].rows[0].cells(2).offsetWidth);
	}
}