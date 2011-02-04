	var DHTML_ARROW_MIN = 37;
	var DHTML_ARROW_MAX = 40;
	var DHTML_LEFT_ARROW = 37; 	// left-arrow
	var DHTML_UP_ARROW = 38; 	// up-arrow
	var DHTML_RIGHT_ARROW = 39;  // right-arrow
	var DHTML_DOWN_ARROW = 40;   // down-arrow



function tabClick( nTab )
{
	event.cancelBubble = true;
	el = event.srcElement;

	for (i = 0; i < tabs.length; i++)
	{
		tabs[i].className = "clsTab";
		newsContent[i].style.display = "none";
	}
	newsContent[nTab].style.display = "block";
	tabs[nTab].className = "clsTabSelected";
}


function ToggleDisplay(oButton, oItems)
{

	if ((oItems.style.display == "") || (oItems.style.display == "none"))	{
		oItems.style.display = "block";
		oButton.src = "/msdn-online/start/images/minus.gif";
	}	else {
		oItems.style.display = "none";
		oButton.src = "/msdn-online/start/images/plus.gif";
	}
	return false;
}

function leftnav_keyup()
{
	var iKey = window.event.keyCode;
	
	// BUGBUG: IE4 returns BODY instead of element with the focus. Use event object instead
	//var oActive = document.activeElement;
	var oActive = window.event.srcElement;
	
	if( DHTML_LEFT_ARROW == iKey || DHTML_RIGHT_ARROW == iKey )
	{
		if ('clsTocHead' == oActive.className)
		{
			// handle headings that expand/collapse
			HandleKeyForHeading(oActive, iKey);
		} 
		else if( "A" == oActive.tagName )
		{
			MoveFocus( oActive, iKey );
		}
	}
	
	return;
}


function MoveFocus( oActive, iKey )
{
	iSrcIndex = oActive.sourceIndex;
	
	if( iKey == DHTML_RIGHT_ARROW)
	{
		while( oItem = document.all[ ++iSrcIndex ] )
		{
			if( !leftNavTable.contains( oItem ) ) return;
			if( "A" == oItem.tagName )
			{
				oItem.focus();
				break;
			}

		}
	}
	else
	{
		while( oItem = document.all[ --iSrcIndex ] )
		{
			if( ( "clsTocHead" == oItem.className || "clsTocHead" == oItem.parentElement.className ) && "A" == oItem.tagName )
			{
				oItem.focus();
				break;
			}

		}
	}
}


// Handle keyboard action on a section
function HandleKeyForHeading(oActive, iKey)
{
	
	sActiveId = oActive.id;
	oItem = document.all[ sActiveId + "Items" ];
	oBtn = document.all[ sActiveId + "Btn" ];

	if( ( "block" != oItem.style.display ) ^ ( DHTML_LEFT_ARROW == iKey ) )
	{
		ToggleDisplay( oBtn ,oItem );
	}
	else
	{
		MoveFocus( oActive, iKey );		
	}
}


function handleMouseover() {
	eSrc = window.event.srcElement;
	eSrcTag=eSrc.tagName.toUpperCase();
	if (eSrcTag=="DIV" && eSrc.className.toUpperCase()=="CLSTOCHEAD")	eSrc.style.textDecoration = "underline";
	if (eSrcTag=="LABEL") eSrc.style.color="#003399";
}

function handleMouseout() {
	eSrc = window.event.srcElement;
	eSrcTag=eSrc.tagName.toUpperCase();
	if (eSrcTag=="DIV" && eSrc.className.toUpperCase()=="CLSTOCHEAD")	eSrc.style.textDecoration = "";
	if (eSrcTag=="LABEL") eSrc.style.color="";
}


document.onmouseover=handleMouseover;
document.onmouseout=handleMouseout;