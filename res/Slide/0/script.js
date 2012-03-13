var vfHaveLoaded=false;
//<LMP>
function RemoveStyles( sld, tagName ) 
{
    var tags = sld.getElementsByTagName( tagName );
    for( var i=0; i<tags.length; i++ )
    {
        var tag = tags[i];
        tag.style.position = "absolute";
        tag.style.left = 0; 
        tag.style.top = 0;
        tag.style.width = "100%";
        tag.style.height = "100%";
    }
}
function AddTextPositionStyles(sld)
{
     var tags = sld.getElementsByTagName( "SPAN" );
    for(var i = 0;i < tags.length;i ++)
    {
        var tag = tags[i];
        var parent = tag.parentNode;
        if(parent.style.position == "absolute" &&
           tag.style.position != "absolute" &&
           tag.lang != "")
        {
            tag.style.position = "relative";
        }
    }
}
function RemoveImgPopupToolbar(sld)
{ 
    var tags = sld.getElementsByTagName( "img" );
    for(var i = 0;i < tags.length;i ++)
    {
        var tag = tags[i];
        tag.galleryImg=false;
    }
}
function CancelKeyDown()
{
    event.cancelBubble = true; 
    return false;	
}
function UnloadSld()
{
document.body.onkeypress = null;
document.body.onclick = null;
document.body.onload = null;
document.body.onkeypresstemp = null;
document.body.onresize = null;
document.body.onunload = null;
}
//</LMP>
function LoadSld()
{
	var sld=GetObj("SlideObj")
	sld.onkeydown = CancelKeyDown;
document.body.onunload = UnloadSld;
if( !g_supportsPPTHTML ) {
		    //<LMP>
    g_origW = sld.offsetWidth;
    g_origH = sld.offsetHeight;
    g_origSz = 16;
    // to fix the image bug on non-IE browser
    RemoveStyles( sld, "v:shape" );
    RemoveStyles( sld, "v:rect" );
    AddTextPositionStyles(sld);
    _RSW();
    //</LMP>
	sld.style.visibility="visible"
		return
	}
	    else
    {
        RemoveImgPopupToolbar(sld);    }
if( MakeNotesVis() ) return
	runAnimations = _InitAnimations();
	if( IsWin("PPTSld") )
		SldUpdated(GetSldId())
	g_origSz=parseInt(SlideObj.style.fontSize)
	g_origH=sld.style.posHeight
	g_origW=sld.style.posWidth
	g_scaleHyperlinks=(document.all.tags("AREA").length>0)
	if( g_scaleHyperlinks )
		InitHLinkArray()
	if( g_scaleInFrame||(IsWin("PPTSld") && IsFullScrMode() ) )
		document.body.scroll="no"
	vfHaveLoaded=true;
	_RSW()
	if( IsWin("PPTSld") && IsFullScrMode() )
		FullScrInit();
	
	MakeSldVis();
	ChkAutoAdv()

	if( runAnimations )
	{
		if( document.all("NSPlay") )
			document.all("NSPlay").autoStart = false;}
}
function MakeSldVis()
{
	var fTrans=g_showAnimation && SldHasTrans()
	if( fTrans )
	{
		if( g_bgSound ) {
			idx=g_bgSound.indexOf(",");
			pptSound.src=g_bgSound.substr( 0, idx );
			pptSound.loop= -(parseInt(g_bgSound.substr(idx+1)));
		}
		SlideObj.filters.revealtrans.Apply()
    }
	SlideObj.style.visibility="visible"}
function MakeNotesVis()
{
	if( !IsNts() ) return false
	SlideObj.style.display="none"
	nObj = document.all.item("NotesObj")
	SetHasNts(0)
	if( nObj ) {
		nObj.style.display=""
		SetHasNts(1)
	}
	return 1
}
function ChkAutoAdv()
{
	if(SldHasTrans())
		SlideObj.onfilterchange=AutoAdv
	else
		AutoAdv()
}
function AutoAdv()
{
	if(!IsWin("PPTSld") || !gUseSldTimings )return
	var sld=GetCurSld()
	if( (sld.mAdvDelay>0) && !IsFramesMode() )
		setTimeout("GoToNextSld()",sld.mAdvDelay)
}
function GetObj(id)
{
	if(g_supportsPPTHTML) return document.all(id);
	else return document.getElementById(id);
}
function SldHasTrans() { return false; } // { return SlideObj.style.filter != ""; }
function GetSldId() { return sId=location.href.substring(location.href.lastIndexOf('/')+1) }
function HideMenu() { if( frames["PPTSld"] && PPTSld.document.all.item("ctxtmenu") && PPTSld.ctxtmenu.style.display!="none" ) { PPTSld.ctxtmenu.style.display='none'; return true } return false }
function IsWin( name ) { return window.name == name }
function IsNts() { return IsWin("PPTNts") }
function IsSldOrNts() { return( IsWin("PPTSld")||IsWin("PPTNts") ) }
function SupportsPPTAnimation() { return( navigator.platform == "Win32" && navigator.appVersion.indexOf("Windows")>0 ) }
function SupportsPPTHTML()
{
	var appVer=navigator.appVersion, msie=appVer.indexOf("MSIE "), ver=0
	if( msie >= 0 )
		ver=parseFloat( appVer.substring( msie+5, appVer.indexOf(";",msie) ) )
	else
		ver=parseInt(appVer)
	return( ver >= 4 && msie >= 0 )
}
function _RSW()
{
	    //<LMP>
    if( navigator.appName == "Microsoft Internet Explorer" )
    {
        _RSW_IE();
    }
    else
    {
        _RSW_NON_IE();
    }
}

function _RSW_NON_IE()
{
    var cltWidth= window.innerWidth; //document.body.clientWidth;
    var cltHeight= window.innerHeight; //document.body.clientHeight;
    var factor=(1.0*cltWidth)/g_origW;
    if( cltHeight < g_origH*factor )
    factor=(1.0*cltHeight)/g_origH;

    newSize = g_origSz * factor;
    if( newSize < 1 ) newSize=1;
        var sld = document.getElementById( "SlideObj" );
    sld.style.fontSize=newSize+"px";
    sld.style.width=g_origW*factor;
    sld.style.height=g_origH*factor;
    sld.style.left=(cltWidth-sld.offsetWidth)/2;
    sld.style.top=(cltHeight-sld.offsetHeight)/2;

    var localhref=window.location.href;
    var specificName=localhref.replace( /[^a-zA-Z0-9]/g, function( $0 ){ return "_" + $0.charCodeAt( 0 ); } );
    var newWindowName = specificName + factor.toString();
    if(newWindowName!=window.name)
    {
        var newWindow = window.open( window.location.href, "_self" );
        newWindow.name = newWindowName;
    }
}

function _RSW_IE()
{
    if (!vfHaveLoaded) return;
    //</LMP>

if( !g_supportsPPTHTML || IsNts() ||
	  ( !g_scaleInFrame && (!IsWin("PPTSld") || !IsFullScrMode()) ) )
		return
        var padding=0;
        if( IsWin("PPTSld") && IsFramesMode() ) padding=6
	cltWidth=document.body.clientWidth-padding
	cltHeight=document.body.clientHeight-padding
	factor=(1.0*cltWidth)/g_origW
	if( cltHeight < g_origH*factor )
		factor=(1.0*cltHeight)/g_origH
	newSize = g_origSz * factor
	if( newSize < 1 ) newSize=1
	s=SlideObj.style
	s.fontSize=newSize+"px"
	s.posWidth=g_origW*factor
	s.posHeight=g_origH*factor
	s.posLeft=(cltWidth-s.posWidth+padding)/2
	s.posTop=(cltHeight-s.posHeight+padding)/2
	if( g_scaleHyperlinks )
		ScaleHyperlinks( factor )
}
function _InitAnimations()
{
	animRuntimeInstalled = ''+document.body.localTime != 'undefined';
	isFullScreen = (window.name == "PPTSld") && !IsFramesMode();
	g_animUseRuntime = g_showAnimation && animRuntimeInstalled && !(isFullScreen && IsSldVisited());
	if( g_animUseRuntime ) {
		collSeq = document.all.tags("seq");
		if( collSeq != null ) {
			for(ii=0;ii<collSeq.length;ii++) {
				if( collSeq[ii].getAttribute( "p:nodeType" ) == "mainSeq" ) {
					g_animMainSequence = collSeq[ii];
					break;
				}
			}
		}
		if( g_animItemsToHide && document.body.playAnimations != false ) {
			for(jj = 0; jj < g_animItemsToHide.length; jj++) {
				if( hideObj = GetObj(g_animItemsToHide[jj]) )
					hideObj.runtimeStyle.visibility="hidden";
			}
		}
		if( g_animInteractiveItems ){
			for(jj = 0; jj < g_animInteractiveItems.length; jj++) {
				if( triggerObj = GetObj(g_animInteractiveItems[jj]) )
					triggerObj.runtimeStyle.cursor="hand";
			}
		}
		if( gUseSldTimings && ''+g_animSlideTime != 'undefined' ) {
			adjustedTime = document.body.calculateAutoAdvanceTimes( g_animSlideTime, g_animEffectTimings );
			if( IsWin("PPTSld") && adjustedTime != g_animSlideTime ) {
			   var sld = GetCurSld();
			   sld.mAdvDelay = adjustedTime * 1000;
			}
		}
	}
	return g_animUseRuntime;
}
gSldJump = 0, gSldJumpTrack = 0, gSldJumpIdx = "";
function _KPH()
{
	if( IsNts() ) return;
	if( !IsFramesMode() && event.keyCode == 27 && !HideMenu() )
		window.close( self );
	else if( event.keyCode == 32 ) {
		if( window.name == "PPTSld" )
			PPTSld.DocumentOnClick();
		else
			GoToNextSld();
	}
	CatchNumKeys( parent, event );
}
function CatchNumKeys( win, event ) {
	if( win.IsFullScrMode() && (48<=event.keyCode) && (event.keyCode<=57) ) {
		gSldJump = 1;
		gSldJumpIdx += (event.keyCode-48).toString();
	}
	if( win.IsFullScrMode() && gSldJump && event.keyCode == 13 ) {
		var numSlds = GetSldList().mList.length
		if ( gSldJumpIdx > numSlds )
			gSldJumpIdx = numSlds;
		if ( gSldJumpIdx >= 0 ) {
			if ( gSldJumpIdx == 0 )
				gSldJumpIdx = 1;
			var jumpTo = parseInt(gSldJumpIdx);
			gSldJump = 0; gSldJumpIdx = "";
			win.GoToSld( GetSldList().mList[jumpTo-1].mSldHref )
		}
	}
}
function _KDH()
{
	if( event.keyCode == 8 ) {
		event.returnValue = 0;
		GoToPrevSld();
	}
}
function DocumentOnClick()
{
	event.cancelBubble = true; 
return;
if( IsNts() || HideMenu() ) return;
	if( ( g_allowAdvOnClick && !IsFramesMode() ) ||
	    (event && (event.keyCode==32) ) )
		GoToNextSld();
}

var g_supportsPPTHTML = SupportsPPTHTML(), g_scaleInFrame = 1, gId="", g_bgSound="",
    g_scaleHyperlinks = false, g_allowAdvOnClick = 1, g_showInBrowser = 1, gLoopCont = 0, gUseSldTimings = 1;
var g_showAnimation = g_supportsPPTHTML && SupportsPPTAnimation() && ( (window.name=="PPTSld" && !IsFramesMode()) || g_showInBrowser );var g_animManager = null;
var g_animUseRuntime = false;
var g_animItemsToHide, g_animInteractiveItems, g_animSlideTime;
var g_animMainSequence = null;
var ENDSHOW_MESG="&#69;&#110;&#100;&#32;&#111;&#102;&#32;&#115;&#108;&#105;&#100;&#101;&#32;&#115;&#104;&#111;&#119;&#44;&#32;&#99;&#108;&#105;&#99;&#107;&#32;&#116;&#111;&#32;&#101;&#120;&#105;&#116;&#46;", SCREEN_MODE="Frames", gIsEndShow=0, NUM_VIS_SLDS=6, SCRIPT_HREF="script.js", FULLSCR_HREF="fullscreen.htm";
var gCurSld = gPrevSld = 1, g_offset = 0, gNtsOpen = gHasNts = gOtlTxtExp = 0, gHasNarration = 0, gOtlOpen = true
window.gPPTHTML=SupportsPPTHTML()
var gMainDoc=new Array(new hrefList("slide0001.htm",1,-1,1),new hrefList("slide0002.htm",1,-1,1),new hrefList("slide0034.htm",1,-1,1),new hrefList("slide0037.htm",1,-1,1),new hrefList("slide0035.htm",1,-1,1),new hrefList("slide0028.htm",1,-1,1));
function FullScrInit()
{
	g_allowAdvOnClick = GetCurSld().mAdvOnClk
	document.body.style.backgroundColor="black"
	document.oncontextmenu=_CM;
	document.onkeydown = _KDH;
	document.ondragstart=Cancel
	document.onselectstart=Cancel
	self.focus()
}
function Redirect( frmId )
{
	var str=document.location.hash,idx=str.indexOf('#'), sId=GetSldId()
	if(idx>=0) str=str.substr(1);
	if( window.name != frmId && ( sId != str) ) {
		obj = GetObj("Main-File")
		window.location.href=obj.href+"#"+sId
		return 1
	}
	return 0
}
var MHTMLPrefix = CalculateMHTMLPrefix();
function CalculateMHTMLPrefix()
{
	if ( document.location.protocol == 'mhtml:') {
		href=new String(document.location.href)
		Start=href.indexOf('!')+1
		End=href.lastIndexOf('/')+1
		if (End < Start)
			return href.substring(0, Start)
		else
		return href.substring(0, End)
	}
	return '';
}
function GetTags(base,tag)
{
	if(g_supportsPPTHTML) return base.all.tags(tag);
	else return base.getElementsByTagName(tag);
}
function UpdNtsPane(){ if(frames["PPTNts"]) PPTNts.location.replace( MHTMLPrefix+GetHrefObj( gCurSld ).mNtsHref ) }
function UpdNavPane( sldIndex ){ if(gNavLoaded) PPTNav.UpdNav() }
function UpdOtNavPane(){ if(gOtlNavLoaded) PPTOtlNav.UpdOtlNav() }
function UpdOtlPane(){ if(gOtlLoaded) PPTOtl.UpdOtl() }
function SetHasNts( fVal )
{
	if( gHasNts != fVal ) {
		gHasNts=fVal
		UpdNavPane()
	}
}
function ToggleOtlText()
{
	gOtlTxtExp=!gOtlTxtExp
	UpdOtlPane()
}
function ClearMedia()
{
	if (PPTSld.pptSound) PPTSld.pptSound.loop = 0;
}
function FullScreen()
{
	if ( PPTSld.g_animUseRuntime )
		PPTSld.document.body.pause();
	ClearMedia();
	var href = ( document.location.protocol == 'mhtml:') ? FULLSCR_HREF : FULLSCR_HREF+"#"+GetHrefObj(gCurSld).mSldHref;
	if(PPTNav.event.ctrlKey) {
		var w = (window.screen.availWidth * 1.0) / 2.0
		var h = w * (PPTSld.g_origH * 1.0) / PPTSld.g_origW
		win = window.open( MHTMLPrefix+href,null,"toolbar=0,resizable=1,top=0,left=0," + "width="+ w + ",height=" + h );
		if( win.document.body && PPTSld.g_animUseRuntime )
			win.document.body.PPTSldFrameset=window;
	}
	else
	{
		win = window.open( MHTMLPrefix+href,null,"fullscreen=yes" );
		if( win.document.body && PPTSld.g_animUseRuntime )
			win.document.body.PPTSldFrameset=window;
	}
}
function ToggleVNarration()
{
	rObj=PPTSld.document.all("NSPlay")
	if( rObj && !PPTSld.g_animUseRuntime ) {
		if( (rObj.playState == 1)||(rObj.playState == 0) )
			rObj.Play()
		else if( rObj.playState == 2 )
			rObj.Pause()
		else
			return;
	}
	else if( PPTSld.g_animUseRuntime )
	{
		narObj = PPTSld.document.all("narrationID")
		if( narObj )
			narObj.togglePause()
	}
}
function GetCurSldNum()
{
	obj=GetHrefObj(gCurSld)
	if( obj.mOrigVis == 1 )
		return obj.mSldIdx
	else
		return gCurSld
}
function GetNumSlds()
{
	if( GetHrefObj(gCurSld).mOrigVis == 1 )
		return GetSldList().mNumVisSlds;
	else
		return GetSldList().mList.length
}
function GetSldNum( href )
{
	for(ii=0; ii<GetSldList().mList.length; ii++) {
		if ( GetSldList().mList[ii].mSldHref == href )
			return ii+1
	}
	return 1
}
function GetHrefObj( sldIdx ){ return GetSldList().mList[sldIdx-1] }
function IsFramesMode(){ return ( SCREEN_MODE == "Frames" ) }
function IsFullScrMode(){ return ( SCREEN_MODE == "FullScreen" ) }
function GoToNextSld()
{
	ii=gCurSld + 1
	if( GetHrefObj( ii-1 ).mOrigVis == 0 ) {
		if( ii<=GetSldList().mList.length ) {
			obj=GetHrefObj(ii)
			obj.mVis=1
			GoToSldNum(ii)
			return
		}
	}
	else {
		obj=GetHrefObj( ii )
		while ( obj && ( obj.mOrigVis == 0 ) )
			obj=GetHrefObj(++ii)
		if( obj && obj.mOrigVis ) {
			GoToSldNum(ii)
			return
		}
	}
	if( gSldStack.length > 1 )
		PopSldList();
	else if( !IsFramesMode() ) {
		if( gLoopCont )
			GoToFirst()
		else
			EndShow()
	}
}
function GoToPrevSld()
{
	ii=gCurSld-1
	if( ii > 0 ) {
		obj=GetHrefObj(ii)
		while ( obj && ( obj.mVis == 0 ) && ( ii>0 ) )
			obj=GetHrefObj(--ii)
		if( ii == 0 ) ii=1
			GoToSldNum(ii)
	}
}
function GoToFirst(){ GoToSld( GetHrefObj(1).mSldHref ) }
function GoToLast()
{
	ii=GetSldList().mList.length
	if( ii != gCurSld )
		GoToSld( GetHrefObj(ii).mSldHref )
}
function GoToSldNum( num )
{
	if( PPTSld.event ) PPTSld.event.cancelBubble=true
	obj = GetHrefObj( num )
	obj.mVis=1
	gPrevSld=gCurSld
	gCurSld = num;
	PPTSld.location.replace(MHTMLPrefix+obj.mSldHref)
	if( IsFramesMode() ) {
		UpdNavPane(); UpdOtlPane(); UpdNtsPane()
	}
}
function GoToSld( href )
{
	if( PPTSld.event ) PPTSld.event.cancelBubble=true
	GetHrefObj( GetSldNum(href) ).mVis=1
	PPTSld.location.replace(MHTMLPrefix+href)
}
function SldUpdated( id )
{
	if( id == GetHrefObj(gCurSld).mSldHref ) return
	gPrevSld=gCurSld
	gCurSld=GetSldNum(id)
	if( IsFramesMode() ) {
		UpdNavPane(); UpdOtlPane(); UpdNtsPane()
	}
}
function PrevSldViewed(){ GoToSld( GetHrefObj(gPrevSld).mSldHref ) }
function HasPrevSld() { return ( gIsEndShow || ( gCurSld != 1 && GetHrefObj( gCurSld-1 ).mVis == 1 )||( GetCurSldNum() > 1 ) ) }
function HasNextSld() { return (GetCurSldNum() != GetNumSlds()) }
function CloseWindow() {
	if( HideMenu() ) return;
	var event = PPTSld.event;
	if( !IsFramesMode() && event && (event.keyCode==27 || event.keyCode==32 || event.type=="click" ) )
		window.close( self );
	CatchNumKeys( self, event );
}
function Unload() { gIsEndShow=0; }
function SetupEndShow() {
	gIsEndShow=1;
	PPTSld.document.body.scroll="no";
	PPTSld.document.onkeypress=CloseWindow;
	PPTSld.document.onclick=CloseWindow;
	PPTSld.document.oncontextmenu=_CM;
}
function EndShow()
{
	if( IsFramesMode() ) return
	if( PPTSld.event ) PPTSld.event.cancelBubble=true
	doc=PPTSld.document
	var dir = doc.body.dir
	if( dir != "rtl" ) dir = "ltr";
	doc.open()
	doc.writeln('<html><body dir=' + dir + ' bgcolor=black onload=SetupEndShow() onunload=Unload()><center><p><font face=Tahoma color=white size=2><br><b>' + ENDSHOW_MESG + '</b></font></p></center></body></html>')
	doc.close()
}
function SetSldVisited(){ GetSldList().mList[gCurSld-1].mVisited=true }
function IsSldVisited(){ return GetSldList().mList[gCurSld-1].mVisited }
function hrefList( sldHref, visible, advDelay, advClk )
{
	this.mSldHref= this.mNtsHref = sldHref
	this.mOrigVis= this.mVis = visible
	this.mVisited= false
	this.mAdvDelay= advDelay
	this.mAdvOnClk= advClk
}
function SldList(arr,curSld,fEnd)
{
	this.mCurSld = curSld;
	this.mList = new Array();
	var idx = 1;
	for(ii=0;ii<arr.length;ii++) {
		this.mList[ii] = new hrefList( arr[ii].mSldHref, arr[ii].mOrigVis, arr[ii].mAdvDelay, arr[ii].mAdvOnClk );
		if( arr[ii].mOrigVis )
			this.mList[ii].mSldIdx = idx++;
	}
	this.mNumVisSlds = idx-1;
	this.fEndShow = fEnd;
}
function GetSldList() {	return gSldStack[gSldStack.length-1] }
function GetCurSld() { return GetSldList().mList[gCurSld - 1] }
gSldStack = new Array();
gSldStack[0] = new SldList(gMainDoc,gCurSld,1)
function ToggleOtlPane()
{
	frmset=document.all("PPTHorizAdjust")
	frm=document.all("PPTOtl")
	if( gOtlOpen )
		frmset.cols="*,100%"
	else
		frmset.cols="25%,*"
	gOtlOpen=!gOtlOpen
	frm.noResize=!frm.noResize
	UpdOtNavPane()
}
function ToggleNtsPane()
{
	frmset=document.all("PPTVertAdjust")
	frm=document.all("PPTNts")
	if( gNtsOpen )
		frmset.rows="100%,*"
	else
		frmset.rows="*,20%"
	gNtsOpen=!gNtsOpen
	UpdNtsPane()
}
function ViewCustomShow(idx,fEnd)
{
	if( !IsFullScrMode() ) return;
	var sldList = new Array();
	var custShow = custShowList[idx-1];
	var jj = 0;
	for( ii=0;ii<custShow.length;ii++ ) {
		if( custShow[ii] <= gMainDoc.length )
			sldList[jj++] = gMainDoc[custShow[ii]-1];
	}
	if (sldList.length > 0) {
	PushSldList(sldList,fEnd);
	gCurSld = 1;
	}
	else
		if( PPTSld.event ) PPTSld.event.cancelBubble=true
}
function PushSldList(arr,fEnd) {
	var ii = gSldStack.length;
	gSldStack[ii] = new SldList(arr,gCurSld,fEnd);
	GoToSld( gSldStack[ii].mList[0].mSldHref );
}
function PopSldList() {
	if (gSldStack[gSldStack.length-1].fEndShow)
		EndShow()
	else {
	gCurSld = gSldStack[gSldStack.length-1].mCurSld;
	gSldStack[gSldStack.length-1] = null;
	gSldStack.length--;
	var sldList = gSldStack[gSldStack.length-1];
	GoToSld( sldList.mList[gCurSld - 1].mSldHref );
	}
}
var custShowList=new Array();
function ImgBtn( oId,bId,w,action )
{
	var t=this
	t.Perform    = _IBP
	t.SetActive  = _IBSetA
	t.SetInactive= _IBSetI
	t.SetPressed = _IBSetP
	t.SetDisabled= _IBSetD
	t.Enabled    = _IBSetE
	t.ChangeIcon = null
	t.UserAction = action
	t.ChgState   = _IBUI
	t.mObjId   = oId
	t.mBorderId= bId
	t.mWidth   = w
	t.mIsOn    = t.mCurState = 0
}
function _IBSetA()
{
	if( this.mIsOn ) {
		obj=this.ChgState( gHiliteClr,gShadowClr,2 )
		obj.style.posTop=0
	}
}
function _IBSetI()
{
	if( this.mIsOn ) {
		obj=this.ChgState( gFaceClr,gFaceClr,1 )
		obj.style.posTop=0
	}
}
function _IBSetP()
{
	if( this.mIsOn ) {
		obj=this.ChgState( gShadowClr,gHiliteClr,2 )
		obj.style.posLeft+=1; obj.style.posTop+=1
	}
}
function _IBSetD()
{
	obj=this.ChgState( gFaceClr,gFaceClr,0 )
	obj.style.posTop=0
}
function _IBSetE( state )
{
	var t=this
	GetObj( t.mBorderId ).style.visibility="visible"
	if( state != t.mIsOn ) {
		t.mIsOn=state
		if( state )
			t.SetInactive()
		else
			t.SetDisabled()
	}
}
function _IBP()
{
	var t=this
	if( t.mIsOn ) {
		if( t.UserAction != null )
			t.UserAction()
		if( t.ChangeIcon ) {
			obj=GetObj(t.mObjId)
			if( t.ChangeIcon() )
				obj.style.posLeft=obj.style.posLeft+(t.mCurState-4)*t.mWidth
			else
				obj.style.posLeft=obj.style.posLeft+(t.mCurState-0)*t.mWidth
		}
		t.SetActive()
	}
}
function _IBUI( clr1,clr2,nextState )
{
	var t=this
	SetBorder( GetObj( t.mBorderId ),clr1,clr2 )
	obj=GetObj( t.mObjId )
	obj.style.posLeft=obj.style.posLeft+(t.mCurState-nextState)*t.mWidth-obj.style.posTop
	t.mCurState=nextState
	return obj
}
function TxtBtn( oId,oeId,action,chkState )
{
	var t=this
	t.Perform    = _TBP
	t.SetActive  = _TBSetA
	t.SetInactive= _TBSetI
	t.SetPressed = _TBSetP
	t.SetDisabled= _TBSetD
	t.SetEnabled = _TBSetE
	t.GetState   = chkState
	t.UserAction = action
	t.ChgState   = _TBUI
	t.mObjId      = oId
	t.m_elementsId= oeId
	t.mIsOn       = 1
}
function _TBSetA()
{
	var t=this
	if( t.mIsOn && !t.GetState() )
		t.ChgState( gHiliteClr,gShadowClr,0,0 )
}
function _TBSetI()
{
	var t=this
	if( t.mIsOn && !t.GetState() )
		t.ChgState( gFaceClr,gFaceClr,0,0 )
}
function _TBSetP()
{
	if( this.mIsOn )
		this.ChgState( gShadowClr,gHiliteClr,1,1 )
}
function _TBSetD()
{
	this.ChgState( gFaceClr,gFaceClr,0,0 )
	this.mIsOn = 0
}
function _TBSetE()
{
	var t=this
	if( !t.GetState() )
		t.ChgState( gFaceClr,gFaceClr,0,0 )
	else
		t.ChgState( gShadowClr,gHiliteClr,1,1 )
	t.mIsOn = 1
}
function _TBP()
{
	var t=this
	if( t.mIsOn ) {
		if( t.UserAction != null )
			t.UserAction()
        if( !t.GetState )
            return
		if( t.GetState() )
			t.SetPressed()
		else
			t.SetActive()
	}
}
function _TBUI( clr1,clr2,lOffset,tOffset )
{
	SetBorder( GetObj( this.mObjId ),clr1,clr2 )
	Offset( GetObj( this.m_elementsId ),lOffset,tOffset )
}
function Offset( obj, top, left ){ obj.style.top=top; obj.style.left=left }
function SetBorder( obj, upperLeft, lowerRight )
{
	s=obj.style;
	s.borderStyle      = "solid"
	s.borderWidth      = 1
	s.borderLeftColor  = s.borderTopColor = upperLeft
	s.borderBottomColor= s.borderRightColor = lowerRight
}
function GetBtnObj(){ return gBtnArr[window.event.srcElement.id] }
function BtnOnOver(){ b=GetBtnObj(); if( b != null ) b.SetActive() }
function BtnOnDown(){ b=GetBtnObj(); if( b != null ) b.SetPressed() }
function BtnOnOut(){ b=GetBtnObj(); if( b != null ) b.SetInactive() }
function BtnOnUp()
{
	b=GetBtnObj()
	if( b != null )
		b.Perform()
	else
		Upd()
}
function GetNtsState(){ return gNtsOpen }
function GetOtlState(){ return gOtlOpen }
function GetOtlTxtState(){ return gOtlTxtExp }
function NtsBtnSetFlag( fVal )
{
	s=document.all.item( this.m_flagId ).style
	s.display="none"
	if( fVal )
		s.display=""
	else
		s.display="none"
}
function _BSetA_Border(){ b = gBtnArr[this.mObjId]; if( b != null ) b.SetActive() }
function _BSetI_Border(){ b = gBtnArr[this.mObjId]; if( b != null ) b.SetInactive() }
function _BSetP_Border(){ b = gBtnArr[this.mObjId]; if( b != null ) b.SetPressed() }
function _BSetA_BorderImg()
{
	b = gBtnArr[this.mBorderId]
	if( b != null && this.mIsOn && !b.GetState() ) {
		obj=this.ChgState( gHiliteClr,gShadowClr,2 )
		obj.style.posTop=0
	}
}
function _BSetI_BorderImg()
{
	b = gBtnArr[this.mBorderId]
	if( b != null && this.mIsOn && !b.GetState() ) {
		obj=this.ChgState( gFaceClr,gFaceClr,1 )
		obj.style.posTop=0
	}
}
var gHiliteClr="THREEDHIGHLIGHT",gShadowClr="THREEDSHADOW",gFaceClr="THREEDFACE"
var gBtnArr = new Array()
gBtnArr["nb_otl"] = new TxtBtn( "nb_otl","nb_otlElem",ToggleOtlPane,GetOtlState )
gBtnArr["nb_otlElem"] = new TxtBtn( "nb_otl","nb_otlElem",ToggleOtlPane,GetOtlState )
gBtnArr["nb_nts"] = new ImgBtn( "nb_nts","nb_ntsBorder",10,ToggleNtsPane )
gBtnArr["nb_nts"].SetActive = _BSetA_BorderImg;
gBtnArr["nb_nts"].SetInactive = _BSetI_BorderImg;
gBtnArr["nb_ntsBorder"] = new TxtBtn( "nb_ntsBorder","nb_ntsElem",ToggleNtsPane,GetNtsState )
gBtnArr["nb_ntsElem"] = new TxtBtn( "nb_ntsBorder","nb_ntsElem",ToggleNtsPane,GetNtsState )
gBtnArr["nb_prevBorder"] = gBtnArr["nb_prev"]= new ImgBtn( "nb_prev","nb_prevBorder",30,GoToPrevSld )
gBtnArr["nb_nextBorder"] = gBtnArr["nb_next"]= new ImgBtn( "nb_next","nb_nextBorder",30,GoToNextSld )
gBtnArr["nb_sldshw"]= new ImgBtn( "nb_sldshw","nb_sldshwBorder",18,FullScreen )
gBtnArr["nb_sldshwBorder"] = new TxtBtn( "nb_sldshw","nb_sldshwBorder",FullScreen,null )
gBtnArr["nb_sldshwBorder"].SetActive = _BSetA_Border;
gBtnArr["nb_sldshwBorder"].SetInactive = _BSetI_Border;
gBtnArr["nb_sldshwText"] = new TxtBtn( "nb_sldshw","nb_sldshwText",FullScreen,null )
gBtnArr["nb_sldshwText"].SetActive = _BSetA_Border;
gBtnArr["nb_sldshwText"].SetInactive = _BSetI_Border;
gBtnArr["nb_voice"] = gBtnArr["nb_voiceBorder"] = new ImgBtn( "nb_voice","nb_voiceBorder",18,ToggleVNarration )
gBtnArr["nb_otlTxtBorder"] = gBtnArr["nb_otlTxt"]= new ImgBtn( "nb_otlTxt","nb_otlTxtBorder",23,ToggleOtlText )
gBtnArr["nb_ntsBorder"].m_flagId= "nb_nts"
gBtnArr["nb_ntsBorder"].SetFlag = NtsBtnSetFlag
gBtnArr["nb_otlTxt"].ChangeIcon= GetOtlTxtState
var sNext="&#78;&#101;&#120;&#116;",sPrev="&#80;&#114;&#101;&#118;&#105;&#111;&#117;&#115;",sEnd="&#69;&#110;&#100;&#32;&#83;&#104;&#111;&#119;",sFont="&#65;&#114;&#105;&#97;&#108;",sArrow="&#65;&#114;&#114;&#111;&#119;",sFreeform="&#70;&#114;&#101;&#101;&#102;&#111;&#114;&#109;",sRect="&#82;&#101;&#99;&#116;&#97;&#110;&#103;&#108;&#101;",sOval="&#79;&#118;&#97;&#108;"
function ShowMenu()
{
	BuildMenu();
	var doc=PPTSld.document.body,x=PPTSld.event.clientX+doc.scrollLeft,y=PPTSld.event.clientY+doc.scrollTop
	m = PPTSld.document.all.item("ctxtmenu")
	m.style.pixelLeft=x
	if( (x+m.scrollWidth > doc.clientWidth)&&(x-m.scrollWidth > 0) )
		m.style.pixelLeft=x-m.scrollWidth
	m.style.pixelTop=y
	if( (y+m.scrollHeight > doc.clientHeight)&&(y-m.scrollHeight > 0) )
		m.style.pixelTop=y-m.scrollHeight
	m.style.display=""
}
function _CM()
{
	if( !IsFullScrMode() ) return;
	if(!PPTSld.event.ctrlKey) {
		ShowMenu()
		return false
	} else
		HideMenu()
}
function BuildMenu()
{
	if( PPTSld.document.all.item("ctxtmenu") ) return
	var mObj=CreateItem( PPTSld.document.body )
	mObj.id="ctxtmenu"
	mObj.style.visibility="hidden"
	var s=mObj.style
	s.position="absolute"
	s.cursor="default"
	s.width="120px"
	SetCMBorder(mObj,"menu","black")
	var iObj=CreateItem( mObj )
	SetCMBorder( iObj, "threedhighlight","threedshadow" )
	iObj.style.padding=2
	CreateMenuItem( iObj,sNext,M_GoNextSld,M_True )
	CreateMenuItem( iObj,sPrev,M_GoPrevSld,M_HasPrevSld )
	CreateSeparator( iObj )
	CreateMenuItem( iObj,sEnd,M_End,M_True )
	mObj.style.visibility="visible"
}
function Cancel() { window.event.cancelBubble=true; window.event.returnValue=false }
function Highlight() { ChangeClr("activecaption","threedhighlight") }
function Deselect() { ChangeClr("threedface","menutext") }
function Perform()
{
	e=PPTSld.event.srcElement
	if( e.type=="menuitem" && e.IsActive() )
		e.Action()
	else
		PPTSld.event.cancelBubble=true
}
function ChangeClr( bg,clr )
{
	e=PPTSld.event.srcElement
	if( e.type=="menuitem" && e.IsActive() ) {
		e.style.backgroundColor=bg
		e.style.color=clr
	}
}
function M_HasPrevSld() { return( HasPrevSld() ) }
function M_GoNextSld() { if( gIsEndShow ) M_End(); else GoToNextSld() }
function M_GoPrevSld() { if( gIsEndShow ) { gIsEndShow=0; history.back(); PPTSld.event.cancelBubble=true; } else GoToPrevSld() }
function M_True() { return true }
function M_End() { window.close( self ) }
function CreateMenuItem( node,text,action,eval )
{
	var e=CreateItem( node )
	e.type="menuitem"
	e.Action=action
	e.IsActive=eval
	e.innerHTML=text
	if( !e.IsActive() )
		e.style.color="threedshadow"
	e.onclick=Perform
	e.onmouseover=Highlight
	e.onmouseout=Deselect
	s=e.style;
	s.fontFamily=sFont
	s.fontSize="9pt"
	s.paddingLeft=2
}
function CreateSeparator( node )
{
	var sObj=CreateItem( node )
	SetCMBorder(sObj,"menu","menu")
	var s=sObj.style
	s.borderTopColor="threedshadow"
	s.borderBottomColor="threedhighlight"
	s.height=1
	s.fontSize="0px"
}
function CreateItem( node )
{
	var elem=PPTSld.document.createElement("DIV")
	node.insertBefore( elem )
	return elem
}
function SetCMBorder( o,ltClr,rbClr )
{
	var s=o.style
	s.backgroundColor="menu"
	s.borderStyle="solid"
	s.borderWidth=1
	s.borderColor=ltClr+" "+rbClr+" "+rbClr+" "+ltClr
}
function WMClick(){if(g_animUseRuntime){window.document.body.pptNext();}}
function Pause(){if(g_animUseRuntime){window.document.body.pause();}}
function Play(){if(g_animUseRuntime){window.document.body.resume();}}
document.oncontextmenu = function(){ return false; } 