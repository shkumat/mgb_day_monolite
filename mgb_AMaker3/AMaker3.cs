// Версия 2.03 от 17.09.2019г. Создание А-файлов с платежными документами
// на основе форматированных данных, полученных из электронной таблицы (CVS/DIF).
// ! Важно : в CSV-файла - разделитель должен быть точка с запятой ;
//
// v2 - добавлена поддержка измененной структуры А-файла с учетом IBAN.
//
using	MyTypes;
using	money	=	System.Decimal	;
//
class	CSepAModel
{
	readonly static int	TOTAL_FIELDS	=	CSepAFileInfo.Record_Field_Size.Length	; // количеств полей в строке A-файла
	long		NumberDoc	=	1			; // нумерация документов
	string[]	Values		= new	string[ TOTAL_FIELDS ]	; // содержимое полей строки A-файла
	string[]	Infos		= new	string[ TOTAL_FIELDS ]	; // название поля для вывода в меню
	string[]	Aliases		= new	string[ TOTAL_FIELDS ]	; // название переменной в шаблонном файле
	int[]		Kinds		= new	int[ TOTAL_FIELDS ]	;
	// Kinds - типы полей:
	//	-1 = редактируемое, обязательное для заполнения;
	//	-2 = постоянное, нередактируемое;
	//	-3 = номер документа;
	//	-4 = редактируемое, необязтельное для заполнения;
	//	иначе - номер столбца из вход. данных.
	//
	public	CSepAModel()
	{

		string		DateStr		=	CCommon.DtoC( CCommon.Today() ).Substring(2, 6);
		string		TimeStr		=	CCommon.Hour(CCommon.Now()).ToString("00") + CCommon.Minute(CCommon.Now()).ToString("00") + CCommon.Second(CCommon.Now()).ToString("00") ;
        	try {
                	NumberDoc	+=	CCommon.CLng( TimeStr ) *1000;
                } catch	( System.Exception Excpt ) {
                	NumberDoc	=	1;
                }
		Kinds[CSepAFileInfo.L_DEBITMFO    ] =	-1 ;	Values[CSepAFileInfo.L_DEBITMFO    ] =	"351629";	Infos[CSepAFileInfo.L_DEBITMFO    ] =	"МФО   Дб. сч."	;	Aliases[CSepAFileInfo.L_DEBITMFO    ]	= "MFOA_Text"	;
		Kinds[CSepAFileInfo.L_DEBITACC    ] =	-1 ;	Values[CSepAFileInfo.L_DEBITACC    ] =	""	;	Infos[CSepAFileInfo.L_DEBITACC    ] =	"Номер Дб. сч."	;	Aliases[CSepAFileInfo.L_DEBITACC    ]	= "AccountA_Text";
		Kinds[CSepAFileInfo.L_CREDITMFO   ] =	-1 ;	Values[CSepAFileInfo.L_CREDITMFO   ] =	"351629";	Infos[CSepAFileInfo.L_CREDITMFO   ] =	"МФО   Кт. сч."	;	Aliases[CSepAFileInfo.L_CREDITMFO   ]	= "MFOCl3_Text"	;
		Kinds[CSepAFileInfo.L_CREDITACC   ] =	-1 ;	Values[CSepAFileInfo.L_CREDITACC   ] =	""	;	Infos[CSepAFileInfo.L_CREDITACC   ] =	"Номер Кт. сч."	;	Aliases[CSepAFileInfo.L_CREDITACC   ]	= "AccCl3_Text"	;
		Kinds[CSepAFileInfo.L_FLAG        ] =	-2 ;	Values[CSepAFileInfo.L_FLAG        ] =	"1"	;	Infos[CSepAFileInfo.L_FLAG        ] =	null		;	Aliases[CSepAFileInfo.L_FLAG        ]	= ""		;
		Kinds[CSepAFileInfo.L_SUMA        ] =	-1 ;	Values[CSepAFileInfo.L_SUMA        ] =	""	;	Infos[CSepAFileInfo.L_SUMA        ] =	"Сумма "	;	Aliases[CSepAFileInfo.L_SUMA        ]	= ""		;
		Kinds[CSepAFileInfo.L_DTYPE       ] =	-2 ;	Values[CSepAFileInfo.L_DTYPE       ] =	"1"	;	Infos[CSepAFileInfo.L_DTYPE       ] =	null		;	Aliases[CSepAFileInfo.L_DTYPE       ]	= ""		;
		Kinds[CSepAFileInfo.L_NDOC        ] =	-3 ;	Values[CSepAFileInfo.L_NDOC        ] =	""	;	Infos[CSepAFileInfo.L_NDOC        ] =	null		;	Aliases[CSepAFileInfo.L_NDOC        ]	= ""		;
		Kinds[CSepAFileInfo.L_CURRENCY    ] =	-1 ;	Values[CSepAFileInfo.L_CURRENCY    ] =	"980"	;	Infos[CSepAFileInfo.L_CURRENCY    ] =	"Числовой код валюты";	Aliases[CSepAFileInfo.L_CURRENCY    ]	= ""		;
		Kinds[CSepAFileInfo.L_DATE1       ] =	-2 ;	Values[CSepAFileInfo.L_DATE1       ] =	DateStr ;	Infos[CSepAFileInfo.L_DATE1       ] =	null		;	Aliases[CSepAFileInfo.L_DATE1       ]	= ""		;
		Kinds[CSepAFileInfo.L_DATE2       ] =	-2 ;	Values[CSepAFileInfo.L_DATE2       ] =	DateStr	;	Infos[CSepAFileInfo.L_DATE2       ] =	null		;	Aliases[CSepAFileInfo.L_DATE2       ]	= ""		;
		Kinds[CSepAFileInfo.L_DEBITNAME   ] =	-1 ;	Values[CSepAFileInfo.L_DEBITNAME ] = "АТ \"МЕГАБАНК\"";	Infos[CSepAFileInfo.L_DEBITNAME   ] =	"Название Дб. сч.";	Aliases[CSepAFileInfo.L_DEBITNAME   ]	= "AName_Text"	;
		Kinds[CSepAFileInfo.L_CREDITNAME  ] =	-1 ;	Values[CSepAFileInfo.L_CREDITNAME  ] =	""	;	Infos[CSepAFileInfo.L_CREDITNAME  ] =	"Название Кт. сч.";	Aliases[CSepAFileInfo.L_CREDITNAME  ]	= "NameCl3_Text";
		Kinds[CSepAFileInfo.L_PURPOSE     ] =	-1 ;	Values[CSepAFileInfo.L_PURPOSE     ] =	""	;	Infos[CSepAFileInfo.L_PURPOSE     ] =	"Назначение платежа";	Aliases[CSepAFileInfo.L_PURPOSE     ]	= "Argument_Text";
		Kinds[CSepAFileInfo.L_RESERVED1   ] =	-2 ;	Values[CSepAFileInfo.L_RESERVED1   ] =	""	;	Infos[CSepAFileInfo.L_RESERVED1   ] =	null		;	Aliases[CSepAFileInfo.L_RESERVED1   ]	= ""		;
		Kinds[CSepAFileInfo.L_DEBITACC_EXT] =	-2 ;	Values[CSepAFileInfo.L_DEBITACC_EXT] =	""	;	Infos[CSepAFileInfo.L_DEBITACC_EXT] =	null		;	Aliases[CSepAFileInfo.L_DEBITACC_EXT]	= ""		;
		Kinds[CSepAFileInfo.L_CREDITACC_EXT]=	-2 ;	Values[CSepAFileInfo.L_CREDITACC_EXT] =	""	;	Infos[CSepAFileInfo.L_CREDITACC_EXT] =	null		;	Aliases[CSepAFileInfo.L_CREDITACC_EXT]	= ""		;
		Kinds[CSepAFileInfo.L_SYMBOL      ] =	-4 ;	Values[CSepAFileInfo.L_SYMBOL      ] =	""	;	Infos[CSepAFileInfo.L_SYMBOL      ] =	"Кассовый символ";	Aliases[CSepAFileInfo.L_SYMBOL      ]	= ""		;
		Kinds[CSepAFileInfo.L_RESERVED2   ] =	-2 ;	Values[CSepAFileInfo.L_RESERVED2   ] =	""	;	Infos[CSepAFileInfo.L_RESERVED2   ] =	null		;	Aliases[CSepAFileInfo.L_RESERVED2   ]	= ""		;
		Kinds[CSepAFileInfo.L_OKPO1       ] =	-1 ;	Values[CSepAFileInfo.L_OKPO1       ] =	"09804119";	Infos[CSepAFileInfo.L_OKPO1       ] =	"Идент код Дб. сч.";	Aliases[CSepAFileInfo.L_OKPO1       ]	= "OKPOA_Text"	;
		Kinds[CSepAFileInfo.L_OKPO2       ] =	-1 ; 	Values[CSepAFileInfo.L_OKPO2       ] =	""	;	Infos[CSepAFileInfo.L_OKPO2       ] =	"Идент код Кт. сч.";	Aliases[CSepAFileInfo.L_OKPO2       ]	= "OKPOCl3_Text";
		Kinds[CSepAFileInfo.L_ID          ] =	-3 ;	Values[CSepAFileInfo.L_ID          ] =	""	;	Infos[CSepAFileInfo.L_ID          ] =	null		;	Aliases[CSepAFileInfo.L_ID          ]	= ""		;
		Kinds[CSepAFileInfo.L_RESERVED3   ] =	-2 ;	Values[CSepAFileInfo.L_RESERVED3   ] =	""	;	Infos[CSepAFileInfo.L_RESERVED3   ] =	null		;	Aliases[CSepAFileInfo.L_RESERVED3   ]	= ""		;
		Kinds[CSepAFileInfo.L_DES         ] =	-2 ;	Values[CSepAFileInfo.L_DES         ] =	""	;	Infos[CSepAFileInfo.L_DES         ] =	null		;	Aliases[CSepAFileInfo.L_DES         ]	= ""		;
		Kinds[CSepAFileInfo.L_DEBITIBAN   ] =	-4 ;	Values[CSepAFileInfo.L_DEBITIBAN   ] =	""	;	Infos[CSepAFileInfo.L_DEBITIBAN   ] =	"IBAN Дб.сч."	;	Aliases[CSepAFileInfo.L_DEBITIBAN   ]	= ""		;
		Kinds[CSepAFileInfo.L_CREDITIBAN  ] =	-4 ;	Values[CSepAFileInfo.L_CREDITIBAN  ] =	""	;	Infos[CSepAFileInfo.L_CREDITIBAN  ] =	"IBAN Кт.сч."	;	Aliases[CSepAFileInfo.L_CREDITIBAN  ]	= ""		;
		Kinds[CSepAFileInfo.L_RESERVED4   ] =	-2 ;	Values[CSepAFileInfo.L_RESERVED4   ] =	""	;	Infos[CSepAFileInfo.L_RESERVED4   ] =	null		;	Aliases[CSepAFileInfo.L_RESERVED4   ]	= ""		;
		Kinds[CSepAFileInfo.L_CRLF        ] =	-2 ;	Values[CSepAFileInfo.L_CRLF        ] =	CAbc.CRLF;	Infos[CSepAFileInfo.L_CRLF        ] =	null		;       Aliases[CSepAFileInfo.L_CRLF        ]	= ""		;
	}
	public	void	IncNumber()
	{
        	NumberDoc++;
    }
	public	int	TotalFields
	{
		get	{
				return	TOTAL_FIELDS;
			}
	}
	public	int	ColOfCsv( int ColNumber )
	{
		if	( ( ColNumber<0 ) || ( ColNumber>=TOTAL_FIELDS ) )
			return	-1;
		if	( Kinds[ ColNumber ] < 0 )
			return	-1;
		else	return	Kinds[ ColNumber ];
	}
	public string this[int Index] {
		get {
			if	( ( Index < TOTAL_FIELDS ) && ( Index >= 0  ) )
                        	if	( Kinds[ Index ] == -3 )		/*  -3 = номер док  */
                                	return	NumberDoc.ToString();
                                else
					return	Values[ Index ] ;
			else
				return "" ;
		}
	}
    public	void	PrintConstValues()
	{
		string	Answer	=	"";
		CConsole.Clear();
		CCommon.Print("\tСодержимое шаблона : ");
        	for	( int CurrentField=0; CurrentField<TotalFields ; CurrentField++ )
                	if	(	( Kinds[ CurrentField ] == -1 )		/*  -1 = постоянное+редактир.  */
                		&&	( Values[ CurrentField ] != null )
				&&	( Aliases[ CurrentField ] != null )
                		)
                        	if	(	( Values[ CurrentField ].Trim() != "" )
					&&	( Aliases[ CurrentField ].Trim() != "" )
                			)
                        		CCommon.Print( CCommon.Left( Infos[ CurrentField ] , 20 ) + " = " + Values[ CurrentField ] ) ;
	}
	public	bool	LoadFromFile( string FileName )
	{
		CTextReader	TextReader	= new	CTextReader();
		string		Tmps		=	""	;
		string[]	SubTmps		;
		if	( FileName == null )
			return	false;
		if	( FileName.Trim() == "" )
			return	false;
		if	( ! TextReader.Open(  FileName , CAbc.CHARSET_WINDOWS ) )
			return	false;
		while	( TextReader.Read() ) {
			Tmps	=	TextReader.Value.Trim();
			if	( Tmps	== null )
				continue;
			if	( Tmps	== "" )
				continue;
			if	( Tmps.Substring( 0 , 1 ) == ";" )
				continue;
			if	( Tmps.IndexOf( "=" ) < 1 )
				continue;
			SubTmps	=	Tmps.Split( CCommon.Chr( 61 ) );
			if	( SubTmps.Length < 2 )
				continue;
			Tmps	=	SubTmps[0].Trim().ToUpper();
	        	for	( int CurrentField=0; CurrentField<TotalFields ; CurrentField++ )
	        		if	(	(	Aliases[ CurrentField ].Trim().ToUpper() == Tmps
	        				)
	        			&&	(	( CurrentField == CSepAFileInfo.L_PURPOSE )
	        				||	( SubTmps[1].IndexOf("/") < 0 )
	        				)
	        			)
	        			Values[ CurrentField ]	=	SubTmps[1].Trim() ;
		}
		TextReader.Close();
		return	true;
	}

	bool	SaveToFile( string FileName )
	{
		int	I	;
		if	( FileName == null )
			return	false;
		else	if	( FileName.Trim() =="" )
			return	false;
		CTextWriter	TextWriter	= new	CTextWriter();
		if	( ! TextWriter.Create(  FileName , CAbc.CHARSET_WINDOWS ) )
			return	false;
		for	( I=0 ; I<TOTAL_FIELDS ; I++ ) {
			if	(	( Kinds[ I ]	==	-1 )
				&&	( Values[ I ]	!=	null )
				&&	( Values[ I ]	!=	"" )
				&&	( Aliases[ I ]	!=	null )
				&&	( Aliases[ I ]	!=	"" )
				)
				TextWriter.Add( Aliases[ I ] + "=" +  Values[ I ].Trim() + CAbc.CRLF ) ;
		}
		TextWriter.Close();
		return	true;
	}

	public	void	AskFixedValues()
	{
		//----------------------------------------------------------
		// если из файла читается кредитовый IBAN, то кредитовый счет запрашивать у пользователя уже не будем
		if ( Kinds[ CSepAFileInfo.L_CREDITIBAN ] > 0 &&  Kinds[ CSepAFileInfo.L_CREDITACC ] == -1 )
		{
			Kinds[ CSepAFileInfo.L_CREDITACC ] = -4;
		}
		string	Answer	=	"";
        	for	( int CurrentField=0; CurrentField<TotalFields ; CurrentField++ )
                	if	( Kinds[ CurrentField ] == -1 ) 	/*  -1 = постоянное+редактир.  */
                                do {
	                        	CCommon.Write( Infos[ CurrentField ]  ) ;
        	                        if	( ( Values[ CurrentField ] != "" ) && ( Values[ CurrentField ] != null  )  )
	        	                	CCommon.Write( " ( " + Values[ CurrentField ] + " )" ) ;
					CCommon.Write( " : ");
                                	Answer	=	CCommon.Input().Trim() ;
					if        ( ( Answer!="" ) &&  ( Answer!=null ))
						Values[ CurrentField ]	=	Answer	;
				}	while	( Values[ CurrentField ] == "" )	;
		CCommon.Write( "Записать шаблон в файл ( пусто = не записывать ) : ")	;
		Answer	=	CCommon.Input();
		if	( Answer != null )
			if	( Answer.Trim() != "" )
				if	( ! SaveToFile( Answer.Trim() + ( ( CCommon.GetExtension( Answer.Trim() ) != "" ) ? "" : ".mod" ) ) )
					CCommon.Print(" Ошибка записи шаблона в файл " + Answer);
		CCommon.Write( "Нумеровать документы начиная с ( "+ NumberDoc.ToString() + " ) : ");
		Answer	=	CCommon.Input();
		if	( ( Answer!="" ) &&  ( Answer!=null ))
			try{
				NumberDoc	=	CCommon.CLng( Answer.Trim() );
			} catch	(System.Exception Excpt) {
			NumberDoc	=	1 ;
			}
	}

	public	bool	RecognizeColumn( int ColNumber )
	{
        	if	( ( ColNumber<0 ) || ( ColNumber>=TOTAL_FIELDS ) )
        		return	false;
        	int	I,MenuCount	=	0;
        	for	( I=0 ; I<TOTAL_FIELDS ; I++ ) {
        		if	( ( Kinds[ I ] == -1 ) || ( Kinds[ I ] == -4 ) )
        			MenuCount++;
        	}
		string[]        MenuItems       = new   string[ MenuCount +1 ]	;
		MenuCount	=	0;
		MenuItems[ MenuCount++ ]	=	" ( пропустить) " ;
        	for	( I=0 ; I<TOTAL_FIELDS ; I++ ) {
        		if	( ( Kinds[ I ] == -1 ) || ( Kinds[ I ] == -4 ) )
        			MenuItems[ MenuCount++ ] = Infos[ I ] ;
        	}
		int	Choice	=	CConsole.GetMenuChoice( MenuItems )	;
		if	( Choice == 0 )
			return	false;
		if	( Choice == 1 )
			return	true;
		Choice--;
		MenuCount	=	0;
        	for	( I=0 ; I<TOTAL_FIELDS ; I++ ) {
        		if	( ( Kinds[ I ] == -1 ) || ( Kinds[ I ] == -4 ) )
        			if	( ( ++MenuCount ) == Choice ) {
        				Kinds[ I ]	=	ColNumber;
        				break;
        			}
        	}
		return	true;
	}
}
//
class	CTsvWriter
{

	System.Text.StringBuilder Tsv		= new	System.Text.StringBuilder();
	CTextReader		TextReader	= new	CTextReader();
	//-----------------------------------------------------------
        //	чтение данных из системного буфера обмена
	public	void	LoadFromClipboard ()
	{
		Tsv.Length	=	0;
		Tsv.Append( CCommon.GetClipboardText() );
		return	;
	}
	//------------------------------------------------------
	//	преобразование Cvs-строки ( разделитель - ; ( точка с запятой ) ) - в Tsv-строку
	string	CommaToTab( string InpStr )
	{
		if	(  InpStr == null )
			return	"";
		else	if	( InpStr.Trim() ==""  )
			return	"";
		if	( InpStr.IndexOf(";") < 0 )
			return	InpStr;
		string		Comma		=	";"						;
		char		Quote		=	CCommon.Chr(34)					;
		string		DoubleQuote	=	Quote.ToString()+Quote.ToString()		;
		string		Result		=	InpStr.Replace( DoubleQuote , CAbc.FORM_FEED )	;
		string[]	Results		=	Result.Split(Quote)				;
		if	( Results.Length > 1 ) {
			int	I	=	0;
			for	( I=0 ; I< Results.Length ; I++)
				if	( ( I % 2 ) != 0 )
					Results[ I ]	=	Results[ I ].Replace( Comma , CAbc.CARRIAGE_RETURN );
			Result	=	System.String.Join( "" , Results );
		}
		Result		=	Result.Replace( Comma , CAbc.TAB ) ;
		Result		=	Result.Replace( CAbc.CARRIAGE_RETURN , Comma ) ;
		Result		=	Result.Replace( CAbc.FORM_FEED , "'"/*Quote.ToString()*/ ) ;
		return		Result;
	}
	//------------------------------------------------------
	//	входящие данные в ДОС-кодировке  ?
	bool	IsDosEncoding( string InpFileName )
	{
		bool	Result	=	false;
		int	Cnt	=	0;
		if	( TextReader.Open(InpFileName ,CAbc.CHARSET_DOS) ) {
			while	( TextReader.Read() ) {
				if	( ++Cnt > 29 )
					break;
				if	(	( TextReader.Value.IndexOf("а") >=0 )
					||	( TextReader.Value.IndexOf("и") >=0 )
					||	( TextReader.Value.IndexOf("е") >=0 )
					||	( TextReader.Value.IndexOf("о") >=0 )
					||	( TextReader.Value.IndexOf("б") >=0 )
					||	( TextReader.Value.IndexOf("в") >=0 )
					||	( TextReader.Value.IndexOf("д") >=0 )
					||	( TextReader.Value.IndexOf("А") >=0 )
					||	( TextReader.Value.IndexOf("И") >=0 )
					||	( TextReader.Value.IndexOf("Е") >=0 )
					||	( TextReader.Value.IndexOf("О") >=0 )
					||	( TextReader.Value.IndexOf("Б") >=0 )
					||	( TextReader.Value.IndexOf("В") >=0 )
					||	( TextReader.Value.IndexOf("Д") >=0 )
					) {
						Result	=	true;
						break	;
                                        }
                        }
			TextReader.Close();
		}
		else	return	false;
		return	Result;
	}
	//-----------------------------------------------------------
	//	чтение данных из  СSV-файла ( разделитель - ; ( точка с запятой ) )
	public	bool	LoadFromCsvFile( string InpFileName )
	{
		int		CharSet		=	CAbc.CHARSET_WINDOWS;
		Tsv.Length	=	0;
		if	(  InpFileName == null )
			return	false;
		else	if	( InpFileName.Trim() ==""  )
			return	false;
		if	( IsDosEncoding( InpFileName ) )
			CharSet	=	CAbc.CHARSET_DOS;
		if	( TextReader.Open( InpFileName , CharSet ) )
			while	( TextReader.Read() )
				Tsv.Append( CommaToTab(TextReader.Value ) + "\r\n" );
		else	return	false;
		TextReader.Close();
		return	true;
	}
	//-----------------------------------------------------------
	//	чтение данных из  DIF-файла
	public	bool	LoadFromDifFile( string InpFileName )
	{
		int		CharSet		=	CAbc.CHARSET_WINDOWS;
		string		Quote		=	CCommon.Chr(34).ToString();
		string		Value		=	""	;
                bool		WaitForAValue	=	false	;
                bool		HasDataStarted	=	false	;
		Tsv.Length	=	0;
		if	(  InpFileName == null )
			return	false;
		else	if	( InpFileName.Trim() ==""  )
			return	false;
		if	( IsDosEncoding( InpFileName ) )
			CharSet	=	CAbc.CHARSET_DOS;
		if	( TextReader.Open(InpFileName , CharSet ) )
			while	( TextReader.Read() ) {
				Value	=	TextReader.Value.Trim();
                                if	( Value.Length < 3 )
                                	continue;
				if	( CCommon.Upper( Value )=="EOD" )
					break;
				if	( CCommon.Upper( Value )=="BOT" )
                                	if	( HasDataStarted )
						Tsv.Append(  CAbc.CRLF );
                                        else	HasDataStarted	=	true;
                                if	( ! HasDataStarted )
                                	continue;
				if	( Value.Substring(0,2)=="0," )
					Tsv.Append( Value.Substring(2).Replace(",",".") + CAbc.TAB );
				if	( Value.Substring(0,2)=="1," )
					WaitForAValue = true ;
				if	( WaitForAValue == true ) {
					if	( Value.Substring(0,1)== Quote )
						Tsv.Append( Value.Substring(1,Value.Length-2).Replace(Quote+Quote,Quote) + CAbc.TAB );
                                        WaitForAValue = true;
				}
			}
		else	return	false;
		TextReader.Close();
		return	true;
	}
	//-----------------------------------------------------------
	//	запись Tsv-текста в файл
	public	bool	SaveToFile( string FileName )
	{
		int		I		;
		bool		ContainsEmpty	;
		string		CurLine		;
		string[]	Columns		;
		if	( FileName == null )
			return	false;
		if	( ( Tsv==null )  || ( Tsv.Length == 0 ) )
			return	false;
		string[]	Lines		=	Tsv.ToString().Split( CCommon.Chr( 13 ) );
		if	( Lines	==	null )
			return	false;
		if	( Lines.Length == 0 )
			return	false;
		bool		Result		=	false;
		CTextWriter	TextWriter	= new	CTextWriter();
		if	( ! TextWriter.Create(  FileName , CAbc.CHARSET_WINDOWS ) )
			return	false;

		for	( I=0 ; I<Lines.Length ; I++ ) {
			Result		=	true;
			CurLine		=	Lines[ I ].Replace( CAbc.BIG_UKR_I, "I" ).Replace( CAbc.SMALL_UKR_I, "i" ).Trim() ;
			if	( CurLine.Trim().Length==0 )
				continue;
			ContainsEmpty	= true;
			Columns	=	CurLine.Split( CCommon.Chr( 9 ) );
			if	( Columns ==  null )
				ContainsEmpty	=	true; 			// есть пустые ячейки ?
			else	if	( Columns.Length < 2 )
					ContainsEmpty	=	true;
				else	foreach	( string Item in Columns )
						if	( Item.Trim().Length != 0 )
							ContainsEmpty	=	false;
			if	( ! ContainsEmpty ) {
				if	( ! TextWriter.Add( CurLine + CAbc.CRLF ) )
					Result		=	false;
			}

		}
		TextWriter.Close();
		return	Result;
	}
}
//
class	AMaker3
{

	const	int			MAX_LINES	=	999	;	// максимальное количество строк в результ. файле
	const	int			MAX_FILES	=	100	;	// максимальное количество результ. файлов
	static	long[]			Cents		= new	long[ MAX_LINES*MAX_FILES ] ; // сумма ( в копейках ) в каждой строке
	static	int			TotalLines	=	0	;	// общее количество строк по входному файлу
	static	long			TotalCents	=	0	;	// общая сумма прводок по входному файлу ( в копейках )
	static	int			BatchNum	=	0	;	// номер результирующего файла
	static	CSepAModel		AModel		= new	CSepAModel();
	static	CSepAWriter		AFile		= new	CSepAWriter();
	static	string			ConstPartOfName	=	"!AUI"
							+	CCommon.StrY( CCommon.Year(CCommon.Now()) & 31 , 1 )
							+	CCommon.StrY( CCommon.Month(CCommon.Now()) , 1)
							+	CCommon.StrY( CCommon.Day(CCommon.Now()) , 1)
							+	CCommon.StrY( CCommon.Hour(CCommon.Now()) & 31 , 1 )
							;
	static	int			AFileExtName	=	CCommon.Second(CCommon.Now());
	static	IFileOfColumnsReader	TsvFile		= new	CCsvReader();

	//----------------------------------------------------------
	//	формирование имени результирующего файла
	static	string	GetAFileName() {
		return	ConstPartOfName
		+	"."
		+	CCommon.StrY( CCommon.Minute(CCommon.Now()) & 31 , 1 )
		+	CCommon.Right( ( BatchNum + AFileExtName ).ToString("000") , 2)
		;
	}
	//----------------------------------------------------------
	//	перевод общей сумма ( в копейках ) в основные единицы валюты
	static	money	TotalBucks
	{
		get {
			money	Result	=	TotalCents;
        		return	Result/100;
        	}
	}
	//----------------------------------------------------------
	//	общая сумма в пачке ( в копейках )
	static	long	TotalCentsInBatch()
	{
		int	I	=	0;
		long    Result	=	0;
		if	( BatchNum<MAX_FILES )
			for	( I = ( BatchNum*MAX_LINES ) ; I < ( (BatchNum+1) *MAX_LINES ) ; I++ )
				Result	+=	Cents[ I ] ;
		return	Result;
	}
        //	общее количество строк в пачке
	static	int	TotalLinesInBatch()
	{
		return	( ( TotalLines > MAX_LINES ) ? MAX_LINES : TotalLines );
	}
	//----------------------------------------------------------
	//	содержит ли строка нулевую сумму или незаполненные номера счетов
	static	bool	IsLineEmpty()
	{
		if	( GetColValue( CSepAFileInfo.L_DEBITACC ).Trim() == "" )
			return	true;
		if	( GetColValue( CSepAFileInfo.L_CREDITACC ).Trim() == ""  && GetColValue( CSepAFileInfo.L_CREDITIBAN ).Trim() == "" )
			return	true;
		long	Val	=	CCommon.CLng( GetColValue(CSepAFileInfo.L_SUMA).Trim() );
		if	( Val < 1 )
			return	true;
		else	return	false;
	}
	//----------------------------------------------------------
        //	получить значение колонки для записи в А-файл
	static	string	GetColValue( int ColNumber ) {
		if	( ColNumber<0 )
			return	"";
		string	Result;
		if	( AModel.ColOfCsv( ColNumber ) < 0 )
			Result	=	AModel[ ColNumber ];
		else	Result	=	TsvFile[ AModel.ColOfCsv( ColNumber ) ];
		if	(	( ColNumber == CSepAFileInfo.L_PURPOSE )
			||	( ColNumber == CSepAFileInfo.L_DEBITNAME )
			||	( ColNumber == CSepAFileInfo.L_CREDITNAME )
			)
			Result	=	Result.Trim().Replace( CAbc.BIG_UKR_I , "I" ).Replace( CAbc.SMALL_UKR_I , "i" ) ;
		if	( ColNumber == CSepAFileInfo.L_SUMA ) {
			Result	=	Result.Trim().Replace( ",",".");
                        money	Crncy	=	CCommon.CCur( Result ) * 100 ;
			long	Val	=	( long ) Crncy ;
			Result	=	CCommon.Right( Val.ToString() , CSepAFileInfo.Record_Field_Size[CSepAFileInfo.L_SUMA] ) ;
		}
		return	Result;
	}
	//-----------------------------------------------------------------------------------------
        //	основная программа
	public static void Main()
	{
		const	bool	DEBUG			=	false;
                const	int	MAX_COLUMNS		=	299	;	// максимальное количество столбцов
                int[]		ColWidth		= new   int[ MAX_COLUMNS ] ;	// ширины столбцов
		int		ColNumber		=	0	;
		int		ALineNumber		=	0	;
		int		AFieldNumber		=	0	;
		string		AFileName		=	""	;
		string		ModelFileName		=	CCommon.GetTempDir()+"\\"+"AMaker.mod";
		string		Now_Date_Str		=	CCommon.DtoC(CCommon.Today()).Substring(2, 6);
		string		Now_Time_Str		=	CCommon.Hour(CCommon.Now()).ToString("00")+ CCommon.Minute(CCommon.Now()).ToString("00");
		CTsvWriter	TsvWriter		= new	CTsvWriter();
		CTextReader	TextReader		= new	CTextReader();
		string		TsvFileName		=	"$"	;
		string		InpFileName		=	""	;
                int		InpColCount		=	0	;
                int		I , SourceMode		=	-2	;	// откуда читать данные : 0=ClipBoard , 1=CSV , 2=DIF
                string		Tmps			=	""	;
                string[]	SubTmps			;
		//----------------------------------------------------------
		CCommon.Print("  Программа для cоздания А-файлов с платежами на основе форматных данных," );
                CCommon.Print("  полученных из электронной таблицы (CVS/DIF). Версия 2.03 от 17.09.2019г." );
		if	( DEBUG )
			InpFileName	=	"F:\\Trash\\Kazna1.csv";
		else
			if( CCommon.ParamCount() < 2 ) {
				CCommon.Print("");
                		CCommon.Print(" Формат запуска : ");
                		CCommon.Print("      AMaker3    Имя_файла  ");
                		CCommon.Print(" где : " );
				CCommon.Print("      Имя_файла  - имя файла данных в формате CSV или DIF ");
                		CCommon.Print("");
                        	return;
                	} else
				InpFileName	=	CAbc.ParamStr[1].Trim();

		if	( InpFileName	==	"*" )
			SourceMode	=	0;
		else
			switch  ( CCommon.GetExtension( InpFileName ).ToUpper() ) {
                		case	".CSV"	: {
	                        	SourceMode	=	1;
        	                	break;
                	        }
                		case	".DIF"	: {
                        		SourceMode	=	2;
	                        	break;
        	                }
                		case	".MOD"	: {
                        		SourceMode	=	-1;
	                        	break;
        	                }
				default	: {
					CCommon.Print("Неправильный тип файла !");
                        		SourceMode	=	-2;
					break;
				}
                	}
                //----------------------------------------------------------
		// если выбран файл с шаблоном, то выводим его на экран
		if	( SourceMode == -1 ) {
			if	( AModel.LoadFromFile( InpFileName ) ) {
				AModel.PrintConstValues();
				if	( CConsole.GetBoxChoice( "Использовать теперь этот шаблон ?"," Да = Enter . Нет = Esc ." ) ) {
					if	( CCommon.FileExists( ModelFileName ) )
						CCommon.DeleteFile( ModelFileName );
					if	( CCommon.FileExists( ModelFileName ) )
						CCommon.Print( "Ошибка удаления файла " + ModelFileName );
					else	if	( ! CCommon.CopyFile( InpFileName , ModelFileName ) )
							CCommon.Print( "Ошибка записи файла  " + ModelFileName );
					return;
				}
			}
		}
		else
			if	( CCommon.FileExists( ModelFileName ) )
				AModel.LoadFromFile( ModelFileName ) ;
		if	( ( SourceMode < 0 ) || ( SourceMode > 2 ) ) {
                	CCommon.Print("Неправильная строка параметров !")  ;
                	return;
                }
                //----------------------------------------------------------
		// скидываем информацию в промежуточный Tsv-файл
		TsvFileName	=	CCommon.GetTempName();
		if	( TsvFileName == null )
			TsvFileName	=	InpFileName + ".$$$" ;
		else	if	( TsvFileName.Trim() == "" )
				TsvFileName	=	InpFileName + ".$$$" ;
		if	( SourceMode == 0 ) {
                        TsvWriter.LoadFromClipboard() ;
                	if	( ! TsvWriter.SaveToFile( TsvFileName ) ) {
				CCommon.Print("Ошибка записи в файл " + TsvFileName);
				return;
			}
		}
		if	( SourceMode > 0 )
                	if	( !CCommon.FileExists( InpFileName ) ) {
                		CCommon.Print("Не найден файл "+ InpFileName );
                		return;
                	}
		if	( SourceMode == 1 )
                        if	( TsvWriter.LoadFromCsvFile( InpFileName ) ) {
                		if	( ! TsvWriter.SaveToFile( TsvFileName ) ) {
					CCommon.Print("Ошибка записи в файл " + TsvFileName);
					return;
				}
			}
			else	{
				CCommon.Print("Ошибка чтения файла " + InpFileName  );
				return;
			}
		if	( SourceMode == 2 )
			if	( TsvWriter.LoadFromDifFile( InpFileName ) ) {
				if	( !TsvWriter.SaveToFile( TsvFileName ) ) {
					CCommon.Print("Ошибка записи в файл " + TsvFileName);
					return;
				}
			}
			else	{
				CCommon.Print("Ошибка чтения файла " + InpFileName  );
				return;
			}
                //----------------------------------------------------------
		// подсчитываем количество столбцов во входящем файле , а также ширину этих столбцов
                for( I=0 ; I<MAX_COLUMNS  ; I++)
                	ColWidth[ I ]	=	0;
                for( I=0 ; I < (MAX_LINES*MAX_FILES) ; I++)
                	Cents[ I ]	=	0;
		if	( ! TextReader.Open( TsvFileName , CAbc.CHARSET_WINDOWS ) ) {
                	CCommon.Print( "Ошибка чтения файла" + TsvFileName ) ;
                	TsvFile.Close();
                	CCommon.DeleteFile(TsvFileName);
                        return;
                }
                if	( ! TextReader.Read() )  {
       	        	CCommon.Print( "Ошибка чтения файла" + TsvFileName ) ;
                	TsvFile.Close();
                	CCommon.DeleteFile(TsvFileName);
               	        return;
                }
                for	( I = 0 ; I<20 ; I++ ) {
			Tmps	=	TextReader.Value;
                	SubTmps	=	Tmps.Split( CCommon.Chr( 9 ) );
                	if	( SubTmps != null ) {
                        	if	( SubTmps.Length > InpColCount )
					InpColCount	=	SubTmps.Length;
                                for	( ColNumber=0 ; ColNumber<SubTmps.Length ; ColNumber++ )
                                	if	( SubTmps[ ColNumber ].Length > ColWidth[ ColNumber ]  )
                                        	ColWidth[ ColNumber ]	=	SubTmps[ ColNumber ].Length    ;
                        }
                	else	InpColCount	=	0;
	                if	( ! TextReader.Read() )
                        	break;
                }
		TextReader.Close();
                if	( InpColCount == 0 ) {
                	CCommon.Print( "Не получается распознать входные данные " ) ;
			CCommon.DeleteFile( TsvFileName );
                        return;
                }
		//----------------------------------------------------------
		// выводим столбцы на экран и запрос пользователю ( помним , что в Csv - файле нумерация столбцов начинается с 1 )
		CConsole.Clear();
		if	( ! TsvFile.Open( TsvFileName , CAbc.CHARSET_WINDOWS ) ) {
			CCommon.Print("Ошибка чтения файла "+ TsvFileName );
			return;
		}
                for	( I = 0 ; I < ( System.Console.WindowHeight - 1 ) ; I++ ) {
			if	( ! TsvFile.Read() )
                        	break ;
                  	Tmps	=	""	;
			for	( ColNumber=0 ; ColNumber<InpColCount ; ColNumber++ )
				Tmps	+=	CCommon.Left( TsvFile[ ColNumber + 1 ] , ColWidth[ ColNumber ] ) + "¦";
			if	( Tmps.Length>0 )
				if	( Tmps.Length < System.Console.WindowWidth - 1  )
					CCommon.Print( Tmps ) ;
				else
					CCommon.Print( Tmps.Substring(0, System.Console.WindowWidth - 2 ) ) ;
                }
		TsvFile.Close();
		if	( ! CConsole.GetBoxChoice( "Для продолжения обработки нажмите Enter.","","Для выхода нажмите Esc. " ) ) {
			CCommon.DeleteFile( TsvFileName );
	                return;
                }
		CConsole.Clear();
		//----------------------------------------------------------
		// распознавание столбцов во входящем файле ( помним , что в Csv - файле нумерация столбцов начинается с 1 )
		for	( ColNumber=0 ; ColNumber<InpColCount ; ColNumber++ ) {
			CConsole.Clear();
			CCommon.Print("");
			if	( ! TsvFile.Open( TsvFileName , CAbc.CHARSET_WINDOWS ) ) {
				CCommon.Print("Ошибка чтения файла "+ TsvFileName );
				TsvFile.Close();
				CCommon.DeleteFile( TsvFileName );
				return;
			}
			for	( I = 0 ; I < ( System.Console.WindowHeight - 1 ) ; I++ )
				if ( ! TsvFile.Read() )
					break;
				else
					CCommon.Print( " " + TsvFile[ ColNumber + 1 ] );
			TsvFile.Close();
			if	( ! AModel.RecognizeColumn( ColNumber + 1 ) ) {
				CCommon.DeleteFile( TsvFileName );
				return;
                        }
		}
		CConsole.Clear();
		//----------------------------------------------------------
		// запрашиваем у пользователя значения постоянных полей
		AModel.AskFixedValues();
		//----------------------------------------------------------
		// подсчитываем количество строк и общую сумму по входному файлу
		TotalLines	=	0;
		TotalCents	=	0;
		if	( ! TsvFile.Open( TsvFileName , CAbc.CHARSET_WINDOWS ) ) {
			CCommon.Print("Ошибка чтения файла "+ TsvFileName );
			TsvFile.Close();
			CCommon.DeleteFile( TsvFileName );
			return;
		}
		while	( TsvFile.Read() ) {
			if	( IsLineEmpty() )
				continue;
                        Cents[ TotalLines ]	=	CCommon.CLng( GetColValue( CSepAFileInfo.L_SUMA ).Trim() ) ;
			TotalCents		+=	Cents[ TotalLines ] ;
                        TotalLines		++	;
		}
		TsvFile.Close();
		//-----------------------------------------------------------
		// запрашиваем у пользователя имя файла , в который будет записан результат
		string	ShortName	=	ConstPartOfName;
		CCommon.Write( "Краткое имя результирующего файла ( " + ShortName + " ) : ");
		ShortName	=	CCommon.Input().Trim();
		if	( ShortName.Length > 0 )
			ConstPartOfName	=	CCommon.Left( ShortName , 8 );
		//-----------------------------------------------------------
		// сверяем с пользователем общее количество строк и общую сумму
		if	( ! CConsole.GetBoxChoice(	" Всего строк : " + CCommon.Right( TotalLines.ToString() , 11 )
                                                ,       " Общая сумма : " + CCommon.StrN( TotalBucks , 11 ).Replace(",",".")
                                                ,       "_________________________________"
                                                ,       "Для продолжения нажмите Enter."
						,	"Для выхода - Esc. "
						)
			) {
			CCommon.DeleteFile( TsvFileName );
	                return;
		}
		CConsole.Clear();
		//-----------------------------------------------------------
		// записываем результатов работы программы в файлы
		if	( ! TsvFile.Open( TsvFileName , CAbc.CHARSET_WINDOWS ) ) {
			CCommon.Print("Ошибка чтения файла "+ TsvFileName );
			TsvFile.Close();
			CCommon.DeleteFile( TsvFileName );
			return;
		}
		BatchNum	=	0;
		//
		while	( TotalLines > 0 ) {
	               	AFileName	=	GetAFileName();
			if	( ! AFile.Create( AFileName , CAbc.CHARSET_DOS ) ) {
				CCommon.Print("Ошибка создания файла "+ AFileName );
				break;
			} else
				CCommon.Print( AFileName );
			//
			AFile.Head[CSepAFileInfo.H_EMPTYSTR   ]	=	"";
			AFile.Head[CSepAFileInfo.H_CRLF1      ]	=	CAbc.CRLF;
			AFile.Head[CSepAFileInfo.H_FILENAME   ]	=	CCommon.Left( AFileName , 12 );
			AFile.Head[CSepAFileInfo.H_DATE       ]	=	Now_Date_Str;
			AFile.Head[CSepAFileInfo.H_TIME       ]	=	Now_Time_Str;
			AFile.Head[CSepAFileInfo.H_STRCOUNT   ]	=	TotalLinesInBatch().ToString();
			AFile.Head[CSepAFileInfo.H_TOTALDEBET ]	=	"0";
			AFile.Head[CSepAFileInfo.H_TOTALCREDIT]	=	TotalCentsInBatch().ToString();
			AFile.Head[CSepAFileInfo.H_DES        ]	=	"0";
			AFile.Head[CSepAFileInfo.H_DES_ID     ]	=	"UIAB00";
			AFile.Head[CSepAFileInfo.H_DES_OF_HEADER]=	"";
			AFile.Head[CSepAFileInfo.H_CRLF2      ]	=	CAbc.CRLF;
			//
			if	( ! AFile.WriteHeader() )  {
				CCommon.Print("Ошибка записи файла "+ AFileName );
                                AFile.Close();
                                break;
                	}
                	//
			for	( ALineNumber = 0 ; ALineNumber < TotalLinesInBatch() ; ALineNumber++ ) {
				do
					if	( ! TsvFile.Read() )
						break	;
				while	( IsLineEmpty() ) ;
				for	( AFieldNumber=0 ; AFieldNumber < AModel.TotalFields ; AFieldNumber++  )
					AFile.Line[ AFieldNumber ]	=	GetColValue( AFieldNumber ) ;
				if	( ! AFile.WriteLine() ) {
					CCommon.Print("Ошибка записи файла "+ AFileName );
					break;
				}
        			AModel.IncNumber();
			}
			AFile.Close();
			BatchNum	++	;
			TotalLines	-=	TotalLinesInBatch() ;
		}
		TsvFile.Close();
		CCommon.DeleteFile( TsvFileName );
	}
}