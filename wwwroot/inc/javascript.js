<!--
window.defaultStatus = "© Skateology";

//highlights de rij waar je muiscursor overheen gaat
function highlightRow(Row, Action)
{
	//Row moet this zijn, Action moet 'over' of 'out' zijn
	var Cells = null;
	var newColor = null;
	Cells = Row.getElementsByTagName('td');
	if (Action == 'over') {
    	newColor = '#EFF5FB';
	}
	if (Action == 'out') {
		newColor = '';
	}
	var c = null;
    var rowCellsCnt	= Cells.length;
	for (c = 0; c < rowCellsCnt; c++) {
		Cells[c].setAttribute('bgcolor', newColor, 0);
	}
}

function newImage(arg) {
	if (document.images) {
		rslt = new Image();
		rslt.src = arg;
		return rslt;
	}
}

function changeImages() {
	if (document.images && (preloadFlag == true)) {
		for (var i=0; i<changeImages.arguments.length; i+=2) {
			document[changeImages.arguments[i]].src = changeImages.arguments[i+1];
		}
	}
}

var preloadFlag = false;
function preloadImages() {
	if (document.images) {
		menu_1_mouseover = newImage("images/menu_1_mouseover.gif");
		menu_2_mouseover = newImage("images/menu_2_mouseover.gif");
		menu_3_mouseover = newImage("images/menu_3_mouseover.gif");
		menu_4_mouseover = newImage("images/menu_4_mouseover.gif");
		menu_5_mouseover = newImage("images/menu_5_mouseover.gif");
		menu_6_mouseover = newImage("images/menu_6_mouseover.gif");
		menu_7_mouseover = newImage("images/menu_7_mouseover.gif");
		preloadFlag = true;
	}
}
// -->

function mailto(prefix,domein)
{
  if (domein == null || domein == "") domein = 'skateology';
  var deel1 = 'mailto:';
  var deel2 = '%40';
  var deel3 = '.';
  var deel4 = 'nl';
  document.location = deel1 + prefix + deel2 + domein + deel3 + deel4;
}