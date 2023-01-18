*******************************************************
* Marco Plaza, 2022 * @nfoxdev
* visualfoxpro color preferences v1.2.0
* https://github.com/nfoxdev/nfthemes
*******************************************************
#Define crlf Chr(13)+Chr(10)
#Define themeextension 'vfptheme'

_Screen.AddProperty('nfthemes',Createobject('nfThemes'))


*************************************
Define Class nfthemes As Form
*************************************

	ShowWindow  = 2
	BorderStyle = 0
	MaxButton	= .F.
	MinButton	= .F.
	AutoCenter	= .T.
	AlwaysOnTop	= .T.
	Width		= 350
	Height		= 450+40
	Caption 	= 'nfTools - ( Click=FontColor RightClick=BackColor )'
	BackColor 	= Rgb(100,100,100)
	ShowTips	= .T.
	foxreg = .F.

	Add Object savetheme As CommandButton With Caption='Save Theme',Visible = .T.,Top=4,Left=10,Height=20,Width=150
	Add Object loadtheme As CommandButton With Caption='Load Theme',Visible = .T.,Top=4,Left=190,Height=20,Width=150

	Add Object pf As PageFrame With ;
		left = 2, Top = 25,Height = 410,Width=348,;
		themes=.F.,TabStyle=0,SpecialEffect=0,TabStretch=0,;
		borderwidth=0,PageCount=10

*--------------------------
	Function Init()
*--------------------------

	Local nexttop
	Local kn
	Local ntab
	Local Array atabnames(1)

	TEXT TO colorsCsv NOSHOW PRETEXT 2+4+8
EditorNormalColor,EditorVariableColor,EditorCommentColor,EditorKeyWordColor,EditorOperatorColor,EditorConstantColor,EditorStringColor,
TraceNormalColor,TraceExecutingColor,TraceCallStackColor,
TraceBreakpointColor,TraceSelectedColor,
WatchNormalColor,WatchSelectedColor,WatchChangedColor,
LocalsNormalColor,LocalsSelectedColor,
OutputNormalColor,OutputSelectedColor,
CallstackNormalColor,CallstackSelectedColor,
BookmarkColor,ShortcutColor
	ENDTEXT

	AddProperty(This,'colorKeys(1)',.Null.)
	Alines(This.colorkeys,m.colorscsv,1,',')

	This.foxreg = Newobject('foxreg',Sys(2004)+'samples\classes\registry.prg')

	This.pf.PageCount = Alines(atabnames,'Editor,Trace,Locals,Watch,Output,Callstack,Bookmark,Shortcut',1,',')

	For ntab = 1 To This.pf.PageCount

		nexttop		= 10

		With This.pf.Pages(m.ntab)

			.Caption = atabnames(m.ntab)
			.BackColor = Rgb(100,100,100)
			.ForeColor = Rgb(255,255,255)

			For Each kn In This.colorkeys

				If m.kn = .Caption
					.AddObject(m.kn,'colorlbl',m.kn,m.nexttop)
					nexttop = m.nexttop + .&kn..Height + 3
				Endif

			Endfor

		Endwith

	Endfor

	This.Visible = .T.


*-----------------------------------------------
	Procedure savetheme.Click()
*-----------------------------------------------
	Local cfile
	Local Page
	Local pp
	Local Array albl(1)

	cfile = Putfile('Save Theme as','my preferences',themeextension)

	If !Empty(m.cfile)
		Thisform.dosavetheme(m.cfile)
	Endif


*----------------------------------------------
	Procedure dosavetheme(m.cfile)
*----------------------------------------------

	cfile = Forceext(m.cfile,themeextension)

	Strtofile('[vfp9colors]'+crlf,m.cfile)

	For Each Page In Thisform.pf.Pages As PageFrame

		Amembers(albl,m.page,2,'U')

		For Each pp In albl
			Strtofile( Proper(m.pp)+' = '+m.page.&pp..regkey+crlf,m.cfile,1)
		Endfor

	Endfor

	Messagebox('Theme saved as '+crlf+m.cfile,0,'nfThemes')


*-----------------------------------------------
	Procedure loadtheme.Click()
*-----------------------------------------------
	Local cfile
	Local oldinireg
	Local Value
	Local colorkey
	Local Page
	Local pp
	Local Array albl(1)

	cfile = Getfile(themeextension,'Load VFP colors')

	If File(m.cfile)

		oldinireg = Newobject('oldinireg',Sys(2004)+'samples\classes\registry.prg')

		For Each colorkey In Thisform.colorkeys
			Value = ''
			oldinireg.getinientry(@m.value,'vfp9colors',m.colorkey,m.cfile)
			Thisform.foxreg.setfoxoption(m.colorkey,m.value)
		Endfor


		For Each Page In Thisform.pf.Pages As PageFrame

			Amembers(albl,m.page,2,'U')

			For Each pp In albl
				m.page.&pp..setcolorfromreg()
			Endfor

		Endfor

		Sys(3056)

		Messagebox('Theme '+Justfname(m.cfile)+' loaded',0,'nfThemes')

	Endif



*****************************************
Enddefine
******************************************

*******************************************
Define Class colorlbl As Label
*******************************************

	Visible		= .T.
	FontSize	= 12
	FontName	= 'Segoe UI'
	Alignment	= 2
	foxreg		= .F.
	Height		= 50
	Left 		= 10
	SpecialEffect=2
	colorkeyname= ''
	fontkeyname=''
	regkey	= ''


*-----------------------------------------------------------
	Procedure Init( colorkeyname,Top)
*----------------------------------------------------------

	Local regval

	With This
		.Caption	= m.colorkeyname
		.colorkeyname = m.colorkeyname
		.Top		= m.top
		.Width 		= Thisform.Width-20
		.ToolTipText = m.colorkeyname
		.setcolorfromreg()
	Endwith

*-------------------------------------
	Function setcolorfromreg()
*-------------------------------------
	Local regval
	regval = ''
	Thisform.foxreg.getfoxoption(This.Caption,@regval)

	Local Array ac(1)
	If Alines(ac,Strextract(m.regval,'RGB(',')',1,1),1,',') # 6
		Dimension ac(6)
		Store '0'   To ac(1),ac(2),ac(3)
		Store '255' To ac(4),ac(5),ac(6)
	Endif

	With This
		.ForeColor	= Rgb(&ac(1),&ac(2),&ac(3))
		.BackColor	= Rgb(&ac(4),&ac(5),&ac(6))
		.regkey		= m.regval
	Endwith


*----------------------------------------------------------
	Function MouseUp(nbutton, nshift, nxcoord, nycoord)
*----------------------------------------------------------


	Local newcolor
	Local csp
	Local fred,fgreen,fblue,bred,bgreen,bblue
	Local tt,tc

	If m.nbutton=1
		tt = ' ForeColor'
		tc = This.ForeColor
	Else
		tt = ' BackColor'
		tc = This.BackColor
	Endif


	Wait  Textmerge('Pick <<m.tt>> For <<rtrim(this.colorKeyName,1,"color")>>') Window Nowait


	newcolor =  Getcolor(m.tc)

	If m.newcolor = -1
		Return
	Endif

	If m.nbutton=1
		This.ForeColor = m.newcolor
	Else
		This.BackColor = m.newcolor
	Endif

	Store '' To fred,fgreen,fblue,bred,bgreen,bblue
	This.color2rgb( This.ForeColor, @m.fred,@m.fgreen,@m.fblue)
	This.color2rgb( This.BackColor, @m.bred,@m.bgreen,@m.bblue)
	This.regkey = Textmerge('RGB(<<m.fred>>,<<m.fgreen>>,<<m.fblue>>,<<m.bred>>,<<m.bgreen>>,<<m.bblue>>), NoAuto, NoAuto')

	Thisform.foxreg.setfoxoption(This.colorkeyname,This.regkey)

	Local csp
	csp = Set('path')
	Sys(3056)
	Set Path To (m.csp)

*---------------------------------------------------
	Procedure color2rgb( tncolor, r , g, b )
*---------------------------------------------------
	r = Bitrshift(Bitand(tncolor, 0x0000ff),0)
	g = Bitrshift(Bitand(tncolor, 0x00ff00),8)
	b = Bitrshift(Bitand(tncolor, 0xff0000),16)

******************************************
Enddefine
******************************************
