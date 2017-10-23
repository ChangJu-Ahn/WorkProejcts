//
//	  Copyright (C) 2003-2014 GrapeCity Inc.	All rights reserved.
//

var FarPoint;(function(a){var b;(function(b){var c;(function(b){var e=function(){function d(a,b){this.resizeHandlers=[];this.eventHandlers=[];this.virtualUpdated=false;this.ignoreNextClick=0;this.chartSpread=a;this.container=b;this.spreadView=a.getSpreadView()}d.prototype.init=function(){var i=this;this.xmlData=typeof this.chartSpread.getXmlDataObject=="function"?this.chartSpread.getXmlDataObject():null;if(this.xmlData==null)this.tableChartData=document.getElementById(this.chartSpread.id+"_XMLDATA");this.chartID=this.container.getAttribute("serverID");this.chart=this.getChart();var f=this;this.container.getChartObj=function(){return f};this.buildResizeCorner();this.chartStyleInfo=new c(this);var e=this.getchartInfoByID();if(this.chart.offsetParent!=null){var b=this.getChartLocation();this.setChartLocation(e,b.x,b.y);var d=$$(this.chart),h=d.innerWidth(),g=d.innerHeight();this.setChartSize(e,h,g)}this.isActived=this.container.getAttribute("isActive")=="true";this.isActived&&this.activeChart();var a=this.chartSpread;if(a.virtualPaging=="true"&&a.virtualTop>0&&this.virtualUpdated!=true){this.container.style.top=parseInt(this.container.style.top)+a.virtualTop+"px";this.virtualUpdated=true}a.preventScrollView&&this.container.removeAttribute("tabindex");this.updateTouchAction();if($$.browser.mb){this.focusDiv=document.createElement("DIV");this.focusDiv.id=this.chartID+"_focusDiv";this.focusDiv.tabIndex=-1;this.focusDiv.style.lineHeight="0px";this.focusDiv.style.height="0px";this.focusDiv.style.outline="none";if(this.spreadView.IsMultiBrowserMode())$$.j(this.focusDiv).insertBefore(this.container.firstChild);else this.container.insertBefore(this.focusDiv,this.container.firstChild);this.container.focus=function(){i.Focus()}}this.registerChartEvents()};d.prototype.updateTouchAction=function(a){if(a==null){var b=document.getElementById(this.chartSpread.id+"_view");if(b)a=$$.browser.mb?$$.j(b).css("touch-action"):b.style.msTouchAction}this.container.style.msTouchAction=a;this.container.style.touchAction=a};d.prototype.getIsDrag=function(){return this.isMoved};d.prototype.getChartLocation=function(){var a=$$(this.chart).clientLocation();a.x+=parseInt(this.container.style.left);a.y+=parseInt(this.container.style.top);if(this.isActived&&!this.hasCustomHighlightCss()){var b=$$(this.container);a.x+=b.border("left");a.y+=b.border("top")}return a};d.prototype.addChartInfo=function(){var a=this.getchartInfoByID();if(a!=null)return a;var d=this.getChartInfo();if(document.all!=null)if(this.xmlData!=null){var c=this.xmlData;a=c.createElement("chart");a.setAttribute("id",this.chartID)}else{a=this.tableChartData.createNode("element","chart","");var b=this.tableChartData.createNode("attribute","id","");b.text=this.chartID;a.attributes.setNamedItem(b)}else{a=document.createElement("chart");a.setAttribute("id",this.chartID)}d.appendChild(a);return a};d.prototype.getchartInfoByID=function(){var c=this.getChartInfo(),b;if(this.xmlData!=null){var a=this.xmlData;a.switchNodeContext(c);var b=a.selectSingleNode("chart[id='"+this.chartID+"']");a.resetNodeContext()}else b=c.selectSingleNode("./chart[@id='"+this.chartID+"']");return b};d.prototype.getChart=function(){for(var c=null,b=this.container.getElementsByTagName("img"),a=0;a<b.length;a++)if(b[a].build==null){c=b[a];break}return c};d.prototype.getChartInfo=function(){var a=this.xmlData;return a!=null?a.selectSingleNode("chartinfo"):typeof this.tableChartData.documentElement=="undefined"?this.tableChartData.getElementsByTagName("chartinfo")[0]:this.tableChartData.documentElement.selectSingleNode("//chartinfo")};d.prototype.getEventTarget=function(a){return a.target==document&&a.currentTarget!=null?a.currentTarget:a.target!=null?a.target:a.srcElement};d.prototype.isResizeHandler=function(a){return this.resizeHandlers.indexOf(a,0)>-1};d.prototype.registerChartEvents=function(){this.attachEvent(this.container,"mousedown",this.containerMousedown);this.attachEvent(this.container,"click",this.containerMouseClick);if(this.container.getAttribute("moveChart")!="false"){var d=$$.browser.mb;this.attachEvent(document,"mouseup",this.endDragChart,d);this.attachEvent(this.container,"mouseup",this.endDragChart);this.attachEvent(this.container,"keydown",this.chartKeyDown);$$.browser.isTouchEventModel&&this.attachEvent(this.container,$$.TouchEvents.TouchCancel,this.endDragChart)}if(this.container.getAttribute("selectChart")!="false"){this.attachEvent(this.container,"contextmenu",this.chartContextMenu);$$.browser.mb&&this.attachEvent(this.container,b.EventManipulator.CONTEXT_MENU,this.chartContextMenu)}if(this.resizeHandlers!=null)for(var a=0;a<this.resizeHandlers.length;a++){var c=this.resizeHandlers[a];this.attachEvent(c,"click",this.handleMouseClick);$$.browser.isTouchEventModel&&this.attachEvent(c,$$.TouchEvents.TouchCancel,this.endResizeChart)}};d.prototype.canSelect=function(){return this.container.getAttribute("selectChart")!="false"};d.prototype.onGestureTap=function(a){if(this.isActived){this.containerMouseClick(a);this.scrollChartIntoView()}else{a.sender=this;this.containerMousedown(a,true);this.chartSpread.UpdatePostbackData();this.container.getAttribute("moveChart")!="false"&&this.endDragChart(a);if(this.ignoreNextClick==null)this.ignoreNextClick=1;else this.ignoreNextClick++}};d.prototype.activeChart=function(){if(this.container.getAttribute("selectChart")!="false"){var b=this.getChartInfo(),a=b.attributes.getNamedItem("activechart"),d=$$.browser;if(d.ie&&this.xmlData==null){a=this.tableChartData.createNode("attribute","activechart","");a.text=this.chartID;b.attributes.setNamedItem(a)}else b.setAttribute("activechart",this.chartID);this.container.setAttribute("isActive","true");this.chartSpread.initialized&&this.spreadView.GetActiveSpreadView(false,true).Spread.ClearSelection(false,null,false);this.chartSpread.SetActiveChartObj(this);this.highlightChart();var c=this;if(!this.chartSpread.preventScrollView)this.__focusTimeoutId=setTimeout(function(){c.Focus()},100);this.buildResizeCorner();$$.browser.webkit&&this.spreadView.EditModePermanent&&this.spreadView.GetSpreadContext().switchHiddenTextBox(true);(this.spreadView.IsInTouchMode()||this.spreadView.IsMultiBrowserMode())&&this.scrollChartIntoView();this.updateTouchAction("none")}};d.prototype.Focus=function(){if(!this.spreadView.IsMultiBrowserMode())this.container.focus();else this.focusDiv.focus()};d.prototype.inactiveChart=function(){var b=this.chartSpread.GetActiveChartObj();if(b!=null){if(b.__focusTimeoutId!=null){clearTimeout(b.__focusTimeoutId);b.__focusTimeoutId=null}b.normalChart();var d=this.getChartInfo();d.removeAttribute("activechart");d.removeAttribute("tabindex");var c=b.chartSpread.getSpreadView();c.IsInTouchMode()&&!c.Touch.closedTouchStripRecently&&c.Touch.IsTouchStripOpening&&a.Web.Spread.TouchController.ActiveTouchStrip.Area==3&&c.Touch.HideTouchStrip()}};d.prototype.scrollChartIntoView=function(){if(this.chartSpread.getAttribute("ClientAutoSize")=="True")return;var a=this.container.parentElement;if(a==null)return;this.stopVirtualPaging();var b=this.container,d=a.scrollTop,c=a.scrollLeft,e=document.getElementById(this.chartSpread.id+"_scrollvp"),i=a.clientWidth,h=a.clientHeight,f=(this.getResizeHandlerWidth()+this.container.clientTop)/2,g=null;if(this.chartSpread.getViewportSize)g=this.chartSpread.getViewportSize();if(g!=null&&g.height<a.offsetHeight)d=0;else{var m=b.offsetTop<a.scrollTop,j=b.offsetTop+b.offsetHeight+f>a.scrollTop+a.clientHeight;if(this.containerMoveUp){if(m)d=b.offsetTop;else if(j||b.offsetTop>this.chartSpread.getViewport().offsetTop+this.chartSpread.getViewport().offsetHeight)d=b.offsetTop+b.offsetHeight-h+f}else if(j)d=b.offsetTop+b.offsetHeight-h+f;else if(m)d=b.offsetTop-(this.spreadView.IsMultiBrowserMode()?f:0)}if(g!=null&&g.width<a.clientWidth)c=0;else{var l=b.offsetLeft<a.scrollLeft,k=b.offsetLeft+b.offsetWidth+f>a.scrollLeft+a.clientWidth;if(this.containerMoveLeft){if(l)c=b.offsetLeft;else if(k)c=b.offsetLeft+b.offsetWidth-i+f}else if(k)c=b.offsetLeft+b.offsetWidth-i+f;else if(l)c=b.offsetLeft}if(d<0)d=0;if(c<0)c=0;!this.isActived&&this.Focus();if(e!=null){var n=e.onscroll;if(e.scrollLeft!=c&&e.scrollTop!=d)e.onscroll=null;if(a.scrollLeft!=c||$$.browser.ieversion<10)a.scrollLeft=c;if(a.scrollTop!=d||$$.browser.ieversion<10)a.scrollTop=d;if(e.scrollLeft!=c||$$.browser.ieversion<10){e.scrollLeft=c;e.scrollLeft=c}e.onscroll=n;if(e.scrollTop!=d||$$.browser.ieversion<10){e.scrollTop=d;e.scrollTop=d}}else{if(c!=null)a.scrollLeft=c;if(d!=null)a.scrollTop=d}this.startVirtualPaging()};d.prototype.stopVirtualPaging=function(){this.chartSpread.stopVirtualPaging(this.chartSpread)};d.prototype.startVirtualPaging=function(){var a=Function.CreateDelegate(this,this.startVirtualPaging2);setTimeout(a,100)};d.prototype.startVirtualPaging2=function(){this.chartSpread.startVirtualPaging(this.chartSpread)};d.prototype.cancelDefault=function(a){if(a.preventDefault)a.preventDefault();else a.returnValue=false;if(a.stopPropagation)a.stopPropagation();else a.cancelBubble=true;return false};d.prototype.preventDefault=function(a){if(a.preventDefault!=null)a.preventDefault();else a.returnValue=false;return false};d.prototype.stopPropagation=function(a){if(a.stopPropagation!=null)a.stopPropagation();else a.cancelBubble=true;return false};d.prototype.getResizeHandler=function(e,h,i,g){for(var b=this.resizeHandlers,a=null,c=0;c<b.length;c++)if(b[c].getAttribute("rstype")==e){a=b[c];break}if(a==null){var d=this.chartSpread!=null?this.chartSpread.id+"_"+this.chartID+"_"+e:null;if(d!=null)a=document.getElementById(d);if(a==null){a=document.createElement("IMG");a.build=true;a.style.position="absolute";a.style.display="none";a.style.cursor=g;a.setAttribute("rstype",e);a.id=d;this.container.appendChild(a)}this.attachEvent(a,"mousedown",this.startResizeChart);this.attachEvent(a,"mouseup",this.endResizeChart);this.attachEvent(a,"click",this.endResizeChart);if(window.navigator.pointerEnabled)this.attachEvent(a,"pointerdown",this.onHandlerPointerDown);else this.attachEvent(a,"MSPointerDown",this.onHandlerPointerDown);b.push(a)}a.style.left=h;a.style.top=i;var f=this.spreadView.IsInTouchMode()?"chartTouchCornerImg":"chartConerImg",j=this.chartSpread.getAttribute(f);a.src=j};d.prototype.getResizeHandlerWidth=function(){var a=this.spreadView.GetSpreadContext();return this.spreadView.IsInTouchMode()||a.isInTouchAjax&&a.isInTouchAjax()?15:7};d.prototype.buildResizeCorner=function(){if(this.container.getAttribute("sizeChart")=="false")return;var k=this.isBuildCorner;if(!k){if(this.container.clientHeight==0){var f=this.container;while(f!=null&&f!=document.body){f.displaySetting=f.style.display;f.style.display="";f=f.parentElement}}this.container.style.height=this.container.clientHeight+"px";this.container.style.width=this.container.clientWidth+"px"}var g=this.container.clientTop/2,j=parseInt(this.container.style.width),i=parseInt(this.container.style.height),l=this.getResizeHandlerWidth(),a=l/2,d="nw-resize",c=-g-a+"px",e=-g-a+"px",b="nw-resize",h=this.getResizeHandler(d,c,e,b);d="ne-resize";c=j+g-a+"px";e=-g-a+"px";b="ne-resize";h=this.getResizeHandler(d,c,e,b);d="sw-resize";c=-g-a+"px";e=i+g-a+"px";b="sw-resize";h=this.getResizeHandler(d,c,e,b);d="se-resize";c=j+g-a+"px";e=i+g-a+"px";b="se-resize";h=this.getResizeHandler(d,c,e,b);d="n-resize";c=j/2-a+"px";e=-g-a+"px";b="s-resize";h=this.getResizeHandler(d,c,e,b);d="s-resize";c=j/2-a+"px";e=i+g-a+"px";b="s-resize";h=this.getResizeHandler(d,c,e,b);d="w-resize";c=-g-a+"px";e=i/2-a+"px";b="w-resize";h=this.getResizeHandler(d,c,e,b);d="e-resize";c=j+g-a+"px";e=i/2-a+"px";b="e-resize";h=this.getResizeHandler(d,c,e,b);if(!k){var f=this.container;while(f!=null&&f!=document.body){if(f.displaySetting!=null){f.style.display=f.displaySetting;f.displaySetting=null}else break;f=f.parentElement}this.isProcessingResize=false;this.isBuildCorner=true}};d.prototype.onHandlerPointerDown=function(b){if(!$$.utils.isTouchEvent(b)){var a=this.chartSpread.getSpreadView();if(a.IsInTouchMode()){a.Touch.ExitFromTouchMode();$$.utils.cancelBubble(b,false)}}};d.prototype.hasCustomHighlightCss=function(){var a=this.container.getAttribute("selectedCssClass");return a!=null&&a.length>0};d.prototype.highlightChart=function(){if(this.hasCustomHighlightCss())this.container.className=this.container.getAttribute("selectedCssClass");else{this.container.style.border="double 1px red";this.container.style.left=parseInt(this.container.style.left)-1+"px";this.container.style.top=parseInt(this.container.style.top)-1+"px"}for(var b=0;b<this.resizeHandlers.length;b++){var c=this.resizeHandlers[b];c.style.display=""}this.container.getAttribute("sizeChart")!="false"&&this.buildResizeCorner();this.isActived=true;if(this.chartSpread.getSpreadView().IsLayoutInitialized()){var a=this.spreadView.GetSpreadContext();if(a.sizeSpread)a.sizeSpread(true);else a.SizeSpread(this.chartSpread,true)}};d.prototype.normalChart=function(){if(!this.isActived)return;this.updateTouchAction();if(this.hasCustomHighlightCss())this.container.className="";else{this.container.style.left=parseInt(this.container.style.left)+1+"px";this.container.style.top=parseInt(this.container.style.top)+1+"px";this.container.style.border="none"}var b=this.spreadView.GetSpreadContext();if($$.browser.ie&&$$.browser.ieversion>10){var e=this.chartSpread!=null?this.chartSpread.id+"_"+this.chartID+"_s-resize":null,d=document.getElementById(e);if(d){var g=d.getBoundingClientRect(),a=this.container.parentElement,h=a.getBoundingClientRect();if(a.scrollHeight>a.clientHeight&&a.scrollHeight-a.clientHeight==a.scrollTop&&b.getViewport().getBoundingClientRect().bottom<this.container.getBoundingClientRect().top)a.scrollTop=Math.max(0,a.scrollTop-Math.ceil(this.getResizeHandlerWidth()/2))}}for(var c=0;c<this.resizeHandlers.length;c++){var f=this.resizeHandlers[c];f.style.display="none"}this.container.removeAttribute("isActive");this.isActived=false;if(this.chartSpread.getSpreadView().IsLayoutInitialized())if(b.sizeSpread)b.sizeSpread(true);else b.SizeSpread(this.chartSpread,true)};d.prototype.setChartLocation=function(c,e,b){if(c==null)return;if(this.chartSpread.virtualPaging=="true"&&this.chartSpread.virtualTop>0)b=b-this.chartSpread.virtualTop;var a=c.getElementsByTagName("location")[0];if(this.xmlData==null){if(a==null)a=this.tableChartData.createNode("element","location","");a.text=e+","+b}else{var d=this.xmlData;if(a==null)a=d.createElement("location");d.textContent(a,e+","+b)}c.appendChild(a);this.chartSpread.UpdatePostbackData()};d.prototype.setChartSize=function(b,e,d){if(b==null)return;var a=b.getElementsByTagName("size")[0],f=$$.browser;if(this.xmlData==null){if(a==null)a=this.tableChartData.createNode("element","size","");a.text=e+","+d}else{var c=this.xmlData;if(a==null)a=c.createElement("size");c.textContent(a,e+","+d)}b.appendChild(a)};d.prototype.getChartSize=function(d){if(d!=null){var b=d.getElementsByTagName("size")[0];if(b!=null){var a;if(this.xmlData!=null)a=this.xmlData.textContent(b);if(!a)b.text||b.textContent||b.innerHTML;if(a){var c=a.indexOf(",");if(c>0){var f=a.substr(0,c),e=a.substr(c+1,a.length-c-1);return{width:parseInt(f),height:parseInt(e)}}}}}return null};d.prototype.attachEvent=function(c,e,d,h,g){if(c==null||e==null||d==null)return;var b;if(c.getAttribute)b=c.getAttribute("rstype");if(!b)b=c.id;if(b==null&&c.nodeName=="#document")b="document"+(g!=null?""+g:"");var f=this.chartID+":"+b+":"+e,a=this.eventHandlers[f];if(a==null){a={};this.eventHandlers[f]=a}if(a[d.toString()]==null){a[d.toString()]=Function.CreateDelegate(this,d);$$.utils.attachEvent(c,e,a[d.toString()],h)}};d.prototype.detachEvent=function(a,d,e,i,g){if(a==null||d==null||e==null||this.eventHandlers==null)return;var b=a.id;if(a.getAttribute)b=a.getAttribute("rstype");if(!b)b=a.id;if(b==null&&a.nodeName=="#document")b="document"+(g!=null?""+g:"");var h=this.chartID+":"+b+":"+d,c=this.eventHandlers[h];if(c==null)return;var f=c[e.toString()];if(f!=null){$$.utils.detachEvent(a,d,f,i);c[e.toString()]=null}};d.prototype.fireEvent=function(b,c,f,g,e,d){if(b==null)return;var a;if(document.createEvent){a=document.createEvent("Events");a.initEvent(c,true,false)}else if(document.createEventObject)a=document.createEventObject();else return;a.chartObj=this;a.top=g;a.left=f;a.width=e;a.height=d;if(b.dispatchEvent)b.dispatchEvent(a);else b.fireEvent&&b.fireEvent("on"+c,a)};d.prototype.updateChartImgitem=function(){if(this.container.getAttribute("imageDirty")=="true"){var a=this.getChart();if(this.chartSpread.inDesign)a.src=a.src;else a.src=a.src+"#"+Math.random()}};d.prototype.isLeftButtonClicked=function(a){var b=a.sender;return $$.browser.ie&&$$.browser.version<11?a.button<2:$$.browser.ie||$$.browser.mozilla?a.buttons==1:a.which==1};d.prototype.dispose=function(){var d=$$.browser.mb;this.detachEvent(this.container,"mousedown",this.containerMousedown);this.detachEvent(this.container,"click",this.containerMouseClick);if(this.container.getAttribute("moveChart")!="false"){this.detachEvent(document,"mouseup",this.endDragChart,d);this.detachEvent(this.container,"mouseup",this.endDragChart);$$.browser.isTouchEventModel&&this.detachEvent(this.container,$$.TouchEvents.TouchCancel,this.endDragChart);this.detachEvent(this.container,"keydown",this.chartKeyDown);this.detachEvent(document,"selectstart",this.cancelDefault);this.detachEvent(this.container,"mousemove",this.dragChart);this.detachEvent(document,"mousemove",this.dragChart,d)}this.detachEvent(this.container,"contextmenu",this.chartContextMenu);$$.browser.mb&&this.detachEvent(this.container,b.EventManipulator.CONTEXT_MENU,this.chartContextMenu);if(this.resizeHandlers!=null)for(var c=0;c<this.resizeHandlers.length;c++){var a=this.resizeHandlers[c];this.detachEvent(a,"mousedown",this.startResizeChart);this.detachEvent(a,"mouseup",this.endResizeChart);this.detachEvent(a,"click",this.endResizeChart);if(window.navigator.pointerEnabled)this.detachEvent(a,"pointerdown",this.onHandlerPointerDown);else this.detachEvent(a,"MSPointerDown",this.onHandlerPointerDown);this.detachEvent(a,"click",this.handleMouseClick);$$.browser.isTouchEventModel&&this.detachEvent(a,$$.TouchEvents.TouchCancel,this.endResizeChart)}};d.prototype.containerMousedown=function(a,c){if(this.ignoreNextClick!=null&&this.ignoreNextClick>0)this.ignoreNextClick--;this.lastMouseDownEvent=a;if(this.container.getAttribute("selectChart")=="false")return;this.chartSpread.SetActiveChart(this.chartID);if(!c&&!this.isLeftButtonClicked(a)||this.containerIsResize==true)return;this.spreadView.OnSpreadChartClick(a);this.container.getAttribute("moveChart")!="false"&&this.startDragChart(a);this.detachEvent(this.container,"click",this.containerMouseClick);this.fireEvent(this.container,"click");this.attachEvent(this.container,"click",this.containerMouseClick);this.buildResizeCorner();if(!this.spreadView.IsMultiBrowserMode())if($$.browser.ieversion>9&&(this.spreadView.FrozenRowCount>0||this.spreadView.FrozenColumnCount>0)){var b=this;setTimeout(function(){b.isActived&&document.activeElement!=b.container&&b.container.focus()},0)}else document.activeElement!=this.container&&this.container.focus();else this.Focus()};d.prototype.processClickForTouchStrip=function(){if(this.spreadView.IsInTouchMode()&&!this.spreadView.Touch.closedTouchStripRecently)if(this.spreadView.Touch.IsTouchStripOpened)this.spreadView.Touch.HideTouchStrip();else if(this.isActived&&this.lastMouseDownEvent!=null){var a=this.lastMouseDownEvent.target;if(a.tagName=="AREA"||a.tagName=="MAP")a=this.container;var b=this.calculateTouchStripOffset(this.lastMouseDownEvent);this.spreadView.Touch.ShowTouchStrip(a,b.X,b.Y)}};d.prototype.containerMouseClick=function(){if(this.endResizeRecently){this.endResizeRecently=false;this.Focus();return}if(this.ignoreNextClick==null||this.ignoreNextClick==0)this.processClickForTouchStrip();else this.ignoreNextClick--;if(this.spreadView.Touch)this.spreadView.Touch.closedTouchStripRecently=false};d.prototype.handleMouseClick=function(){if(this.endResizeRecently){this.endResizeRecently=false;this.Focus();return}this.processClickForTouchStrip();if(this.ignoreNextClick!=null&&this.ignoreNextClick>0)this.ignoreNextClick--;if(this.spreadView.Touch)this.spreadView.Touch.closedTouchStripRecently=false};d.prototype.startResizeChart=function(c){var b=$$.browser.mb;this.endResizeRecently=false;var d=c.target||c.srcElement;this.isDrag=false;this.containerIsResize=true;try{var a=window,e=0;do{var f=a.document;this.detachEvent(f,"mouseup",this.endResizeChart,b,e);this.attachEvent(f,"mouseup",this.endResizeChart,b,e);if(e<$$.browser.getAccessParentLevel()&&a.parent!=null&&a!=a.parent){a=a.parent;e+=1}else break}while(true)}catch(g){$$.DEBUG.Log(g,"startResizeChart")}this.attachEvent(document,"click",this.endResizeChart,b);this.attachEvent(document,"selectstart",this.cancelDefault);this.cornerIsResize=true;this.resizeStartX=c.clientX;this.resizeStartY=c.clientY;(!b||$$.browser.ie&&this.spreadView.IsInTouchMode())&&typeof d.setCapture!="undefined"&&d.setCapture();this.attachEvent(d,"mousemove",this.resizeChart);this.attachEvent(document,"mousemove",this.resizeChart,b);this.currentResizeCorner=d;this.preventDefault(c)};d.prototype.resizeChart=function(c){if(!this.isLeftButtonClicked(c))return;var e=this.currentResizeCorner;if(e==null)return;var b=e.getAttribute("rstype"),f=this.chartStyleInfo,h=Math.max(f.getStyleWidth(),18),g=Math.max(f.getStyleHeight(),18);if(e==null)return;if(this.isProcessingResize==true)return;this.isProcessingResize=true;this.container.style.cursor=e.style.cursor;var d={left:parseInt(this.container.style.left),top:parseInt(this.container.style.top),width:this.container.clientWidth,height:this.container.clientHeight},a={left:d.left,top:d.top,width:d.width,height:d.height};if(b=="e-resize"||b=="ne-resize"||b=="se-resize")a.width=Math.max(h,parseInt(this.container.style.width)+c.clientX-this.resizeStartX);if(b=="s-resize"||b=="sw-resize"||b=="se-resize")a.height=Math.max(g,parseInt(this.container.style.height)+c.clientY-this.resizeStartY);if(b=="n-resize"||b=="nw-resize"||b=="ne-resize"){var j=Math.max(0,parseInt(this.container.style.top)+c.clientY-this.resizeStartY);a.height+=d.top-j;a.top=j}if(b=="w-resize"||b=="sw-resize"||b=="nw-resize"){var i=Math.max(0,parseInt(this.container.style.left)+c.clientX-this.resizeStartX);a.width+=d.left-i;a.left=i}if(a.height<=g)a.height=g+2;if(a.width<=h)a.width=h+2;if(a){this.container.style.left=a.left+"px";this.container.style.top=a.top+"px";this.container.style.width=a.width+"px";this.container.style.height=a.height+"px";f.setInnerSize(a.width,a.height)}this.buildResizeCorner();this.resizeStartX=c.clientX;this.resizeStartY=c.clientY;this.cancelDefault(c);var k=this.chartSpread.getSpreadView();if(!k.IsMultiBrowserMode())this.chartSpread.ResumeLayout(true);else{k.RelayoutViewport();$$.utils.cancelBubble(c)}this.isProcessingResize=false};d.prototype.endResizeChart=function(l){var c=$$.browser.mb?true:false;if(typeof this.currentResizeCorner=="undefined")return;try{var a=window,f=0;do{var m=a.document;this.detachEvent(m,"mouseup",this.endResizeChart,c,f);if(f<$$.browser.getAccessParentLevel()&&a.parent!=null&&a!=a.parent){a=a.parent;f+=1}else break}while(true)}catch(n){$$.DEBUG.Log(n,"startResizeChart")}this.detachEvent(document,"mousemove",this.resizeChart,c);this.detachEvent(document,"click",this.endResizeChart,c);this.detachEvent(document,"selectstart",this.cancelDefault);this.containerIsResize=false;this.container.style.cursor="default";var d=this.currentResizeCorner;if(d!=null&&this.cornerIsResize!=null&&this.cornerIsResize==true){var b=this.getchartInfoByID(),e=b!=null?this.getChartSize(b):null;this.endResizeRecently=e==null||e.width!=this.chart.width||e.height!=this.chart.height;this.cornerIsResize=false;this.detachEvent(d,"mousemove",this.resizeChart);(!c||$$.browser.ie&&this.spreadView.IsInTouchMode())&&typeof d.releaseCapture!="undefined"&&d.releaseCapture();b=this.addChartInfo();var g=this.getChartLocation(),j=g.x,k=g.y;this.setChartLocation(b,j,k);var i=parseInt(""+this.chart.width),h=parseInt(""+this.chart.height);this.setChartSize(b,i,h);this.chartSpread.RefreshChart();this.fireEvent(this.container,"resize",j,k,i,h)}if($$.browser.ie&&$$.browser.ieversion<11)this.chartSpread.ResumeLayout(true);else this.chartSpread.sizeSpread(true);this.currentResizeCorner=null;this.cancelDefault(l)};d.prototype.startDragChart=function(a){var b=$$.browser.mb;this.container.style.cursor="move";this.isDrag=true;this.isMoved=false;this.oldLocation=this.getChartLocation();this.containerX=this.moveStartX=a.clientX;this.containerY=this.moveStartY=a.clientY;this.attachEvent(this.container,"mousemove",this.dragChart);(!b||$$.browser.ie&&this.spreadView.IsInTouchMode())&&typeof this.container.setCapture!="undefined"&&this.container.setCapture();!$$.browser.ie&&this.container.setAttribute("tabindex","0");this.attachEvent(document,"mousemove",this.dragChart,b);this.attachEvent(document,"selectstart",this.cancelDefault);this.cancelDefault(a)};d.prototype.dragChart=function(a){if(!this.isLeftButtonClicked(a)){this.endDragChart(a);return}if(this.isDrag!=null&&this.isDrag==true){this.container.style.left=parseInt(this.container.style.left)+a.clientX-this.moveStartX+"px";this.container.style.top=parseInt(this.container.style.top)+a.clientY-this.moveStartY+"px";this.moveStartX=a.clientX;this.moveStartY=a.clientY}var b=this.chartSpread.getSpreadView();if(!b.IsMultiBrowserMode())this.chartSpread.ResumeLayout(true);else{b.RelayoutViewport();$$.utils.cancelBubble(a)}};d.prototype.endDragChart=function(a){if(!$$.utils.contains(document.body,this.container)){this.dispose();return}if(this.isDrag!=null&&this.isDrag==true){this.container.style.cursor="default";this.isDrag=false;var d=$$.browser.mb;(!d||$$.browser.ie&&this.spreadView.IsInTouchMode())&&typeof this.container.releaseCapture!="undefined"&&this.container.releaseCapture();this.detachEvent(this.container,"mousemove",this.dragChart);this.detachEvent(document,"mousemove",this.dragChart,d);this.detachEvent(document,"selectstart",this.cancelDefault);this.cancelDefault(a);var h=this.addChartInfo(),b=this.getChartLocation(),f=b.x,g=b.y;this.setChartLocation(h,f,g);this.fireEvent(this.container,"move",f,g);if(this.oldLocation.x!=b.x||this.oldLocation.y!=b.y){if(this.ignoreNextClick==null)this.ignoreNextClick=1;else this.ignoreNextClick++;this.isMoved=true}this.oldLocation=null;this.stopVirtualPaging();if($$.browser.ie&&$$.browser.ieversion<11)this.chartSpread.ResumeLayout(true);else this.chartSpread.sizeSpread(true);this.startVirtualPaging();if(this.chartSpread.getViewportSize){var c=false,e=this.chartSpread.getViewportSize();if(e.height<this.container.parentElement.offsetHeight&&this.container.parentElement.scrollTop>0)c=true;if(e.width<this.container.parentElement.offsetWidth&&this.container.parentElement.scrollLeft>0)c=true;if(c&&$$.browser.ie){this.scrollChartIntoView();this.chartSpread.ResumeLayout(true)}}if(this.containerX==a.clientX||this.containerY==a.clientY){this.scrollChartIntoView();if(this.containerX==a.clientX)this.containerMoveLeft=false;if(this.containerY==a.clientY)this.containerMoveUp=false}else{this.containerMoveLeft=this.containerX>a.clientX;this.containerMoveUp=this.containerY>a.clientY;this.moveStartX=a.clientX;this.moveStartY=a.clientY}}};d.prototype.calculateTouchStripOffset=function(a){var b=0,c=0;if($$.browser.mb&&(a.target.tagName=="AREA"||a.target.tagName=="MAP")){var d=this.container.getBoundingClientRect();b=a.clientX-d.left;c=a.clientY-d.top}else{b=a.offsetX;c=a.offsetY}return{X:b,Y:c}};d.prototype.chartContextMenu=function(c){var b=this.spreadView;if(b!=null){if(b.IsInTouchMode()){var e,d;if(b.Touch.IgnoreContextMenuEvent){b.Touch.IgnoreContextMenuEvent=false;d=true}else if(this.isActived){var f=this.calculateTouchStripOffset(c);e=b.Touch.ShowTouchStrip(c&&$$.utils.contains(this.container,c.target)?c.target:this.container,Math.max(0,f.X),Math.max(0,f.Y));if(e){this.containerIsResize=false;this.isDrag=false}}if(e||b.Touch.IsTouchStripOpened){d=true;if(this.__focusTimeoutId!=null){clearTimeout(this.__focusTimeoutId);this.__focusTimeoutId=null}a.Web.Spread.TouchController.ActiveTouchStrip&&a.Web.Spread.TouchController.ActiveTouchStrip.Focus()}d&&$$.utils.cancelBubble(c,true)}b.ContextMenus!=null&&!$$.browser.isTouchEventModel&&$$.utils.cancelBubble(c,false)}};d.prototype.chartKeyDown=function(a){if(a.keyCode==38)this.container.style.top=parseInt(this.container.style.top+0)-1+"px";else if(a.keyCode==40)this.container.style.top=parseInt(this.container.style.top+0)+1+"px";else if(a.keyCode==37)this.container.style.left=parseInt(this.container.style.left+0)-1+"px";else if(a.keyCode==39)this.container.style.left=parseInt(this.container.style.left+0)+1+"px";var c=this.addChartInfo(),b=this.getChartLocation();this.setChartLocation(c,b.x,b.y);this.cancelDefault(a);this.stopVirtualPaging();if($$.browser.ie&&$$.browser.ieversion<11)this.chartSpread.ResumeLayout(true);else this.chartSpread.sizeSpread(true);this.startVirtualPaging();return false};return d}();b.SpreadChart=e;var d=function(){function a(a){this.container=a.container;this.init()}a.prototype.init=function(){var a=this.container,c=this.getCornerWidth()/2,b=this.getBorderWidth();this.Left=a.offsetLeft;this.Right=a.offsetLeft+a.clientWidth;this.Top=a.offsetTop;this.Bottom=a.offsetTop+a.clientHeight;this.X=this.Left;this.Y=this.Top;this.Width=Math.max(this.Right-this.Left,0);this.Height=Math.max(this.Bottom-this.Top,0)};a.prototype.getCornerWidth=function(){return 7};a.prototype.getBorderWidth=function(){return 1};a.prototype.focus=function(){this.container!=null&&this.container.focus()};a.prototype.select=function(){this.container!=null&&this.container.select()};a.prototype.isSelect=function(){return this.container.select};return a}();b.SpreadChartContainer=d;var c=function(){function a(a){this.spreadChart=a;this.init()}a.prototype.init=function(){var b=this.getChartImage(),a=b.style;this.width=b.width;this.height=b.height;this.border={top:a.borderTopWidth!=""?parseInt(a.borderTopWidth):0,right:a.borderRightWidth!=""?parseInt(a.borderRightWidth):0,bottom:a.borderBottomWidth!=""?parseInt(a.borderBottomWidth):0,left:a.borderLeftWidth!=""?parseInt(a.borderLeftWidth):0};this.margin={top:a.marginTop!=""?parseInt(a.marginTop):0,right:a.marginRight!=""?parseInt(a.marginRight):0,bottom:a.marginBottom!=""?parseInt(a.marginBottom):0,left:a.marginLeft!=""?parseInt(a.marginLeft):0};this.padding={top:a.paddingTop!=""?parseInt(a.paddingTop):0,right:a.paddingRight!=""?parseInt(a.paddingRight):0,bottom:a.paddingBottom!=""?parseInt(a.paddingBottom):0,left:a.paddingLeft!=""?parseInt(a.paddingLeft):0}};a.prototype.getChartImage=function(){return this.spreadChart.getChart()};a.prototype.getStyleWidth=function(){return this.border.left+this.border.right+(this.margin.left+this.margin.right)+(this.padding.left+this.padding.right)};a.prototype.getTotalWidth=function(){return this.width+this.getStyleWidth()};a.prototype.getStyleHeight=function(){return this.border.top+this.border.bottom+(this.margin.top+this.margin.bottom)+(this.padding.top+this.padding.bottom)};a.prototype.getTotalHeight=function(){return this.height+this.getStyleHeight()};a.prototype.setInnerSize=function(c,b){var a=this.getChartImage();a.style.width=Math.max(c-this.getStyleWidth(),1)+"px";a.style.height=Math.max(b-this.getStyleHeight(),1)+"px";this.init()};return a}();b.SpreadChartStyleInfo=c;(function(a){a[a.OUTSIDE=-1]="OUTSIDE";a[a.INSIDE=0]="INSIDE";a[a.INTERSECT=1]="INTERSECT"})(b.ChartPosition||(b.ChartPosition={}));var g=b.ChartPosition,f=function(){function a(){}a.prototype.isIE=function(){return $$.browser.ie};a.prototype.getViewportChartIntersection=function(a,c,h,f,b,d,g,i){var e=-1;if(b<a&&b+i>a||b>a&&b<a+h)e=1;else if(c>d&&c<d+g||d>c&&d<c+f)e=1;if(e==1)if(a>b&&a+h<b+i&&c>d&&c+f<d+g||b>a&&b+g<a+f&&d>c&&d+g<c+f)e=0;return e};a.prototype.getChartsIdCollection=function(f){for(var b=[],c=0,d=f.getElementsByTagName("DIV"),a=0;a<d.length;a++){var e=d[a];if(e.id.indexOf("_chartContainer")>0){b[c]=e.id;c++}}return b};a.prototype.getChartElements=function(e){for(var b=[],c=e.getElementsByTagName("DIV"),a=0;a<c.length;a++){var d=c[a];d.id.indexOf("_chartContainer")>0&&b.push(d)}return b};a.prototype.IsActiveChartExisted=function(d){for(var b=this.getChartElements(d),a=0;a<b.length;a++){var c=b[a];if(c&&c.getAttribute("isActive")=="true")return true}return false};a.prototype.getUnionChartsBounds=function(e){var f=e.getSpreadView?e.getSpreadView():null,d=f!=null&&f.IsMultiBrowserMode()?f.GetCharts():this.getChartsIdCollection(e);if(d.length==0)return null;for(var b={top:0,left:0,right:0,bottom:0},g=0;g<d.length;g++){var i=d[g],a=document.getElementById(i);if(a&&typeof a.getChartObj=="function"){var h=a.getChartObj(),c=h.isActived?h.getResizeHandlerWidth()/2:0,m=a.offsetTop-c,k=a.offsetLeft-c,l=a.offsetLeft+a.offsetWidth+c,j=a.offsetTop+a.offsetHeight+c;b.top=Math.min(m,b.top);b.left=Math.min(k,b.left);b.right=Math.max(l,b.right);b.bottom=Math.max(j,b.bottom)}}return b};a.prototype.doMapAreaPostback=function(c,e,d){var b=document.getElementById(c.id);if(b!=null){if(typeof b.parentElement.parentElement.getChartObj=="undefined")return;var a=b.parentElement.parentElement.getChartObj();typeof a!="undefined"&&a!=null&&!a.getIsDrag()&&__doPostBack(e,d)}};return a}();b.ChartUtil=f;window.FpChartUtil=new a.Web.Spread.ChartUtil})(c=b.Spread||(b.Spread={}))})(b=a.Web||(a.Web={}))})(FarPoint||(FarPoint={}))