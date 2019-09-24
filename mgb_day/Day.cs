// Версия 3.13 от 02.10.2018г. Оболочка открытия и закрытия дня в `Скрудже`.
// encoding=cp-1251
/*
		Параметры программы :

	-mode		Режим работы :
		open	открытие дня ;
		close	закрытие дня .
*/
using MyTypes;

class	Day {
	static	CScrooge2Config	Scrooge2Config				;
	static	CConnection	Connection						;
	static	CRecordSet	RecordSet						;
	static	CCommand	Command							;
	static	CArray		MNames		= new	CArray()	;
	static	CArray		MKinds		= new	CArray()	;
	static	CArray		MParams		= new	CArray()	;
	static	CArray		MCommands	= new	CArray()	;
	static	string		ConnectionString=	null		;
	static	string		Answer		=	null			;
	static	int			Mode		=	0				;
	static	int			DayDate		=	0				;
	static	int			NextDate	=	0				;
	static	string[]	MenuNames						;
	static	string[]	MenuKinds						;
	static	string[]	MenuCommands					;
	static	string[]	MenuParams						;
	static	string		InfNBU_OutPath=	CAbc.EMPTY		;
	static	string		UserCode	=	CCommon.GetUserName() ;
	static	int			TODAY		=	CCommon.Today()	;
	static	int			DefaultNextDate	=	( CCommon.DayOfWeek( TODAY ) == 5 ) ? ( TODAY + 3 ) : ( TODAY + 1 ) ;
	static	string		NEXT_STR	=	CCommon.StrD( DefaultNextDate , 10,10).Substring(6)
						+	CCommon.StrD( DefaultNextDate , 10,10).Substring(2,4)
						+	CCommon.StrD( DefaultNextDate , 10,10).Substring(0,2);
	// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	public static	void	Main() {
		const	bool	DEBUG		=	false		;
		CParam		Param		= new	CParam()	;
		int		Choice		=	0		;
		bool		Result		=	false		;
		byte		SavedColor	=	7		;
		string		OutputDir	=	CAbc.EMPTY	;
		string		ScroogeDir	=	CAbc.EMPTY	;
		string		ServerName	=	CAbc.EMPTY	;
		string		DataBase	=	CAbc.EMPTY	;
		string		LogFileName	=	CAbc.EMPTY	;
		string		UnsolverBillsCmdText =	CAbc.EMPTY	;
		int		UnsolvedBillsExists =	0		;
		string[]	Titles		=	{ "ОТКРЫТИЕ ДНЯ" , "ЗАКРЫТИЕ ДНЯ" } ;
		string[]	LogFileNames	=	{ "dayopen.log" , "dayclose.log" } ;
		string		NOW_STR		=	CCommon.StrD( CCommon.Today() , 10,10).Substring(6)
						+	CCommon.StrD( CCommon.Today() , 10,10).Substring(2,4)
						+	CCommon.StrD( CCommon.Today() , 10,10).Substring(0,2);
		Err.LogToConsole() ;
		CConsole.Clear();
		CCommon.Print( " Оболочка открытия и закрытия дня в `Скрудже`. Версия 3.13 от 02.10.2018г." , "" ) ;
		if	( DEBUG )
			Mode=1;
		else
			if	( CCommon.IsEmpty( Param["Mode"] ) )
				CCommon.Print( "Не указан режим работы программы !" ) ;
			else
				switch	( Param["Mode"].ToUpper() ) {
					case	"OPEN" : {
						Mode	=	1;
						break;
					}
					case	"CLOSE" : {
						Mode	=	2;
						break;
					}
					default	: {
						CCommon.Print( "Неверно указан режим работы программы !" ) ;
						break;
					}
				}
		if	( ( Mode != 1 ) && ( Mode != 2 ) )
			return;
		if	( ( Mode==2 ) && ( ! CCommon.DirExists( "Z:\\") )  ) {
			SavedColor              =       CConsole.BoxColor ;
			CConsole.BoxColor       =       CConsole.RED*16 + CConsole.WHITE;
			CConsole.ShowBox( "" , "Не подключен Z:\\" , "" );
			CConsole.ClearKeyboard();
			CConsole.ReadChar();
			CConsole.BoxColor	=	SavedColor;
			CConsole.Clear();
			CConsole.ShowCursor();
		}
		// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		//	Запрос пользователю на ввод даты дня, который открывается или закрывается
		while	( DayDate == 0 ) {
			switch	( Mode ) {
				case 1	: {
					CCommon.Write( "Какую дату открываем ? ( " + NOW_STR.Replace("/",".") +" ) ");
					break;
				}
				default	: {
					CCommon.Write( "Какую дату закрываем ? ( " + NOW_STR.Replace("/",".") +" ) ");
					break;
				}
			}
			Answer	=	CCommon.Input().Trim() ;
			if	( Answer=="" )
				DayDate		=	CCommon.Today();
			else
				DayDate		=	CCommon.GetDate( Answer );
		}
		// - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Scrooge2Config	= new	CScrooge2Config();
		if (!Scrooge2Config.IsValid) {
			CCommon.Print( Scrooge2Config.ErrInfo ) ;
			return;
		}
		ScroogeDir	=	(string)Scrooge2Config["Root"];
		OutputDir	=	(string)Scrooge2Config["Output"];
		ServerName	=	(string)Scrooge2Config["Server"];
		DataBase	=	(string)Scrooge2Config["DataBase"];
		if( ScroogeDir == null ) {
			CCommon.Print("  Не найдена переменная `Root` в настройках `Скрудж-2` ");
			return;
		}
		if( ServerName == null ) {
			CCommon.Print("  Не найдена переменная `Server` в настройках `Скрудж-2` ");
			return;
		}
		if( DataBase == null ) {
			CCommon.Print("  Не найдена переменная `Database` в настройках `Скрудж-2` ");
			return;
		}
		CCommon.Print(" Беру настройки `Скрудж-2` здесь :  " + ScroogeDir );
		// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		ConnectionString	=	"Server="	+	ServerName
					+	";Database="	+	DataBase
					+	";Integrated Security=TRUE;"  ;
		Connection = new CConnection(ConnectionString);
		if	( Connection.IsOpen()) {
			CCommon.Print(" Сервер " + ServerName  );
			CCommon.Print(" База   " + DataBase + CAbc.CRLF );
		}
		else {
			CCommon.Print( CAbc.CRLF + "  Ошибка подключения к серверу !" );
			return;
		}
		System.Console.Title="  " + Titles[Mode-1] + "             |   "+ServerName+"."+DataBase	;
		Command		= new	CCommand( Connection );
		int	IsFullAccess	=	( int ) CCommon.IsNull( Command.GetScalar(" select  dbo.Fn_IsFullAccess( DEFAULT ) " ) , (int) 0 );
		string	InfNBU_OutPath	=	( string ) CCommon.IsNull( Command.GetScalar(" exec dbo.Mega_Day_Open;5  @Mode=2 ; " ) , (string) CAbc.EMPTY );
		if	( Mode == 2 ) {
			// Записываю в историю логин пользователя, который закрывает день
			Command.Execute( " exec  dbo.Mega_Day_Close;12  @DayDate = " + DayDate.ToString() + " , @UserCode = '" + UserCode + "' " );
			// Проверка : имеются ли непроведенные документы ?
			UnsolverBillsCmdText	=	" If Exists ( select 1 from dbo.Mega_SV_WaitingBills with (NoLock) where  (DayDate=" + DayDate.ToString() +" )  and ((PermitFlag & 255 )!=255) and ((ProcessFlag & 3)!=3) ) select Convert(Integer,1) else select Convert(Integer,0) ";
			UnsolvedBillsExists	=	( int ) CCommon.IsNull( Command.GetScalar( UnsolverBillsCmdText ) , (int) 0 );
		}
		Command.Close();
		Connection.Close();
		if	( IsFullAccess < 1 ) {
			CCommon.Print( CAbc.CRLF + "  Для работы программы пользователю необходим полный доступ в Скрудже !" );
			return;
		}
		InfNBU_OutPath	=	InfNBU_OutPath.Trim();
		if	( InfNBU_OutPath.Length==0  )
			Err.Print( CCommon.Now().ToString() + "Ошибка определения выходного каталога для ОДБ." + CAbc.CRLF + CAbc.CRLF );
		if	( UnsolvedBillsExists == 1 ) {
			SavedColor		=       CConsole.BoxColor ;
			CConsole.BoxColor       =       CConsole.RED*16 + CConsole.WHITE;
			CConsole.ShowBox( "" , "Имеются непpоведенные документы","","  Пpоведите их или удалите !","");
			CConsole.ClearKeyboard();
			CConsole.ReadChar();
			CConsole.BoxColor	=	SavedColor;
			CConsole.Clear();
			CConsole.ShowCursor();
		}
		// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		if	( OutputDir !=	null ) {
			OutputDir=	ScroogeDir + "\\" + OutputDir.Trim();
			if	( ! CCommon.DirExists( OutputDir ) )
				CCommon.MkDir( OutputDir );
			if	( CCommon.DirExists( OutputDir ) ) {
				OutputDir	+=	"\\" + CCommon.StrD( DayDate , 8 , 8 ).Replace("/","").Replace(".","");
				if	( ! CCommon.DirExists( OutputDir ) )
					CCommon.MkDir( OutputDir );
				if	( ! CCommon.DirExists( OutputDir ) )
					OutputDir	=	ScroogeDir + "\\" ;
				}
			LogFileName		=	OutputDir + "\\" + LogFileNames[Mode-1] ;
		}
		else
			LogFileName		=	ScroogeDir + "\\" + LogFileNames[Mode-1] ;
		Err.LogTo( LogFileName );
		// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		//	Основной цикл программы
		do {
			System.Console.Title="  " + Titles[Mode-1] + "  " + CCommon.StrD( DayDate , 8 , 8 ) + "        |   "+ServerName+"."+DataBase	;
			string[]	FilesForMEDOC	=	CCommon.GetFileList( InfNBU_OutPath + "\\@*.*" );
			if	( FilesForMEDOC != null )
				if	( FilesForMEDOC.Length > 0 )
					CConsole.GetBoxChoice(
						""
					,	"В каталоге " + InfNBU_OutPath
					,	"Найдены файлы для налоговой  @*.*"
					,	"Не забудьте их отправить!"
					,	""
					);
			if	( ! LoadHistory( Mode ) )
				Err.Print( CCommon.Now().ToString() + "Ошибка получения с сервера истории работы программы !" );
			Choice	=	CConsole.GetMenuChoice( MenuNames )	;
			if	( Choice>0 ) {
				Err.Print( CCommon.Now().ToString() + "\t< " + MenuNames[ Choice-1 ] + "   ( " + UserCode + " ) " + CAbc.CRLF );
				if	( ! SaveHistory( Mode , Choice ) )
					Err.Print( "Ошибка сохранения на сервере истории работы программы !" );
				switch	(  MenuKinds[ Choice-1 ].Trim().ToUpper() ) {
					case	"EXE.COPY" : {
						Result=StartExeCopy(
							MacroSubstitution( MenuCommands[ Choice-1 ] )
						,	MacroSubstitution(
								DateSubstitution(
									MenuParams[ Choice-1 ]
								)
							)
						);
						break;
					}
					case	"SQL.CMD" : {
						Result=StartSqlCmd(
							DateSubstitution(
								MenuCommands[ Choice-1 ]
							)
						);
						break;
					}
					case	"SQL.RS" : {
						Result=StartSqlRS(
							DateSubstitution(
								MenuCommands[ Choice-1 ]
							)
						);
						break;
					}
					case	"SQL.ROUTINE" : {
						Result=StartSqlRoutine(
							DateSubstitution(
								MenuCommands[ Choice-1 ]
							)
						);
						break;
					}
					case	"EXC" : {
						Result=StartExc( ScroogeDir
						,	MacroSubstitution(
								DateSubstitution(
									MenuCommands[ Choice-1 ]
								)
							)
						);
						break;
					}
					case	"URL.IE" : {
						Result=StartUrl(
							DateSubstitution(
								MenuCommands[ Choice-1 ]
							)
						);
						break;
					}
					default	: {
						break;
					}
				}
				Err.Print( CCommon.Now().ToString() + "\t " + MenuNames[ Choice-1 ] + "   ( " + Result.ToString() + " ) > " + CAbc.CRLF );
			}
		} while	( Choice !=0  );
	}
	// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	//	Запуск EXC-программы с помощью Скрудж-2
	static	bool	StartExc( string ScroogeDir , string ExcName ) {
		return	CCommon.Shell( ScroogeDir + "\\EXE\\CII32.EXE" , ExcName , "open" , 1 );
	}
	// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	//	Открытие URL-адреса в программе  Internet Explorer
	static	bool	StartUrl( string Url ) {
		return	CCommon.Shell( "C:\\Program Files\\Internet Explorer\\IExplore.exe" , Url , "open" , 1 );
	}
	// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	//	Подстановка дат ( [DayDate] и [NextDate] )
	static	string DateSubstitution( string SrcStr ) {
		if	( SrcStr == null )
			return	CAbc.EMPTY;
		if	( SrcStr.Trim() == "" )
			return	SrcStr;
		string	Result	=	SrcStr.Trim().ToUpper();
		if	( ( Result.IndexOf("[NEXTDATE]") >= 0 ) && ( NextDate==0 ) ) {
			// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			//	Запрос пользователю на ввод даты дня, который закрывается
			while	( NextDate == 0 ) {
				CCommon.Write( "Какой следующий рабочий день ? ( " + NEXT_STR.Replace("/",".") +" ) ");
				Answer	=	CCommon.Input().Trim() ;
				if	( Answer=="" )
					NextDate	=	DefaultNextDate;
				else
					NextDate	=	CCommon.GetDate( Answer );
			}
		}
		Result	=	Result.Replace( "[DAYDATE]" , DayDate.ToString() );
		Result	=	Result.Replace( "[NEXTDATE]" , NextDate.ToString() );
		return		Result;
	}
	// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	//	Подстановка переменных окружения , пример MacroSubstitution("-=[%temp%]=-")
	static	string MacroSubstitution( string SrcStr ) {
		if	( SrcStr == null )
			return	CAbc.EMPTY;
		if	( SrcStr.Trim() == "" )
			return	SrcStr;
		int	Pos	=	SrcStr.IndexOf("%");
		if	( Pos < 0 )
			return	SrcStr;
		if	( SrcStr.IndexOf("%",Pos+1) < Pos )
			return	SrcStr;
		int	Pos2	=	SrcStr.IndexOf("%",Pos+1);
		if	( Pos2<(Pos+3) )
			return	SrcStr;
		string	Result	=	SrcStr.Substring(Pos+1,Pos2-Pos-1);
		Result	=	CCommon.GetEnvStr( Result );
		if	( Pos>0 )
			Result	=	SrcStr.Substring(0,Pos)+Result;
		if	( Pos2<(SrcStr.Length-1) )
			Result	=	Result + SrcStr.Substring( Pos2 + 1 );
		return	Result;
	}
	// - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	//	Вычитка списка пунктов открытия/закрытия дня
	static	bool	LoadHistory( int Mode ) {
		int	MenuItemWidth	= ( Mode == 1 ) ? 32 : 25 ;
		Connection = new CConnection(ConnectionString);
		RecordSet	= new   CRecordSet( Connection );
		bool	Result		=	false;
		if	( RecordSet.Open("exec Mega_Day_Close;10  @DayDate=" + DayDate.ToString() + " , @Mode= " + Mode.ToString()) ) {
			while	( RecordSet.Read()  ) {
				if	( CCommon.CInt( RecordSet["Cnt"] ) == 0 )
					MNames.Add( CCommon.Left( RecordSet["Name"] , MenuItemWidth ) + "*" );
				else
					MNames.Add( CCommon.Left( RecordSet["Name"] , MenuItemWidth ) + "* " + RecordSet["Cnt"].Trim());
				MCommands.Add( RecordSet["Command"] );
				MParams.Add( RecordSet["Params"] );
				MKinds.Add( RecordSet["Kind"] );
			Result	=	true;
			}
		}
		else
			Result	=	false;
		RecordSet.Close();
		int		MenuCount	=	MNames.Count;
		MenuNames			= new   string[ MenuCount ] ;
		MenuKinds			= new   string[ MenuCount ] ;
		MenuParams			= new   string[ MenuCount ] ;
		MenuCommands			= new   string[ MenuCount ] ;
		MenuCount			=	0;
		foreach	( string MenuName in MNames )
			MenuNames[ MenuCount ++ ] = MenuName;
		MenuCount			=	0;
		foreach	( string MenuKind in MKinds )
			MenuKinds[ MenuCount ++ ] = MenuKind;
		MenuCount			=	0;
		foreach	( string MenuParam in MParams )
			MenuParams[ MenuCount ++ ] = MenuParam;
		MenuCount			=	0;
		foreach	( string MenuCommand in MCommands )
			MenuCommands[ MenuCount ++ ] = MenuCommand;
		MCommands.Clear();
		MParams.Clear();
		MKinds.Clear();
		MNames.Clear();
		Connection.Close();
		return	Result;
	}
	// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	//	Вычитка истории вполнения пунктов открытия/закрытия дня
	static	bool	SaveHistory( int Mode , int Choice ) {
		Connection	= new	CConnection(ConnectionString);
		Command		= new	CCommand( Connection );
		bool	Result	=	Command.Execute( " exec  dbo.Mega_Day_Close;11  @DayDate=" + DayDate.ToString() + " , @FlagCode= " + Choice.ToString() + " , @Mode= " + Mode.ToString());
		Command.Close();
		Connection.Close();
		return	Result;
	}
	// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	//	Запуск EXE-файла в том же каталоге, что и данная программа
	static	bool	StartExeCopy( string SrcExeName , string Params ) {
		if	( ( SrcExeName == null ) || ( Params == null )  )
			return	false;
		SrcExeName	=	SrcExeName.Trim();
		Params		=	Params.Trim();
		if	( SrcExeName == "" )
			return	false;
		string	Path	=	CCommon.GetTaskDir() + "\\";
		string	Mask	=	SrcExeName.Substring( 0 , SrcExeName.Length-3 ) + "*";
		string	ExeName	=	Path + CCommon.GetFileName( SrcExeName );
		if	( ! CCommon.FileExists( ExeName ) )
			foreach	( string FileName in CCommon.GetFileList( Mask ) )
				if	( ! CCommon.CopyFile( FileName , Path + CCommon.GetFileName( FileName ) ) )
					return	false;
		if	( ! CCommon.FileExists( ExeName ) )
			return	false;
		return	CCommon.Run( ExeName , Params );
	}
	// - - - - - - - - - - - - - - - - - - - - - - - - -
	//	Запуск команды на Sql-сервере
	static	bool	StartSqlCmd( string CmdText ) {
		if	( CmdText == null )
			return	false;
		if	( CmdText.Trim() == "" )
			return	false;
		CConsole.ShowBox("","Выполнение команды на сервере","");
		Connection	= new	CConnection(ConnectionString);
		Command		= new	CCommand( Connection );
		Command.Timeout	=	599 ;
		bool	Result	=	Command.Execute( CmdText );
		Command.Close();
		Connection.Close();
		CConsole.Clear();
		CCommon.Print(CAbc.EMPTY,"Для продолжения нажмите Enter.");
		CConsole.ClearKeyboard();
		CConsole.Flash();
		CConsole.ReadChar();
		return	Result;
	}
	// - - - - - - - - - - - - - - - - - - - - - - - -
	//	Запуск выборки на Sql-сервере
	static	bool	StartSqlRS( string CmdText ) {
		if	( CmdText == null )
			return	false;
		if	( CmdText.Trim() == "" )
			return	false;
		CConsole.ShowBox("","Выполнение команды на сервере","");
		Connection	= new	CConnection(ConnectionString);
		RecordSet	= new	CRecordSet( Connection );
		bool	Result;
		RecordSet.Timeout	=	599  ;
		if	( RecordSet.Open( CmdText ) ) {
			CConsole.Clear();
			int	FieldCount=RecordSet.FieldCount();
			while	( RecordSet.Read()  )
				for	( int Index=0; Index<FieldCount; Index++ )
					CCommon.Print( RecordSet[ Index ] ) ;
		}
		else
			CConsole.Clear();
		RecordSet.Close();
		Connection.Close();
		CCommon.Print(CAbc.EMPTY,"Для продолжения нажмите Enter.");
		CConsole.ClearKeyboard();
		CConsole.Flash();
		CConsole.ReadChar();
		return	true;
	}
	// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	//	Запуск  списка регламентных процедур на Sql-сервере
	static	bool	StartSqlRoutine( string CmdText ) {
		CConnection	Connection2;
		bool	Result	=	true;
		if	( CmdText == null )
			return	false;
		if	( CmdText.Trim() == "" )
			return	false;
		CArray	CommandList	= new	CArray();
		CArray	NamesList	= new	CArray();
		CConsole.ShowBox("","Выполнение команды на сервере","");
		Connection	= new	CConnection(ConnectionString);
		RecordSet	= new	CRecordSet( Connection );
		if	( RecordSet.Open( CmdText ) ) {
			while	( RecordSet.Read()  ) {
				NamesList.Add( RecordSet[ "Name" ] );
				CommandList.Add( RecordSet[ "Command" ] );
			}
		}
		else
			Result	=	false;
		CConsole.Clear();
		RecordSet.Close();
		string		Cmd		=	CAbc.EMPTY
		,		Msg		=	CAbc.EMPTY
		,		Results		=	CAbc.EMPTY;
		CCommand	Command		= new	CCommand( Connection );
		for	( int I=0 ; I< NamesList.Count ; I++ ) {
			CCommon.Print( NamesList[ I ] + " - выполняется." );
	       		CConsole.ShowBox("","Выполнение команды на сервере","");
 			Cmd		=	DateSubstitution( (string) CommandList[I] );
			CmdText		=	" exec  dbo.Mega_Common_WriteToLoginfo "
					+	" @TaskCode='Mega_Day_Routine-"+Mode.ToString().Trim()+"'"
					+	",@Info='start : "+Cmd+"'";
			Command.Execute( CmdText ) ;
			Connection2	= new	CConnection( ConnectionString ) ;
			RecordSet	= new	CRecordSet( Connection2 );
			RecordSet.Timeout=	599;
			if	( RecordSet.Open( Cmd ) ) {
				Msg	=	(string) NamesList[ I ];
				int	Count	=	0;
				if	( RecordSet.Read() ) {
					do
						Count++;
					while	( RecordSet.Read() );
				}
				Msg	=	NamesList[ I ] + " - OK."
					+	( ( Count > 0 ) ? " ( " + Count.ToString() + " row(s) affected )" : ""  );
			}
			else {
				Result	=	false;
				Msg	=	NamesList[ I ] + " - Ошибка !";
			}
			RecordSet.Close();
			Connection2.Close();
			CmdText		=	" exec  dbo.Mega_Common_WriteToLoginfo "
					+	" @TaskCode='Mega_Day_Routine-"+Mode.ToString().Trim()+"'"
					+	",@Info='stop  : "+Cmd+"'";
			Command.Execute( CmdText ) ;
			CmdText		=	" exec  dbo.Mega_Common_WriteToLoginfo "
					+	" @TaskCode='Mega_Day_Routine-"+Mode.ToString().Trim()+"'"
					+	",@Info='"+ Msg.Replace("'","`") +"'";
			Command.Execute( CmdText ) ;
			Err.Print( CAbc.TAB + CAbc.TAB + Msg + CAbc.CRLF );
			CConsole.Clear();
			Results		+=	Msg	+ CAbc.CRLF ;
			CCommon.Write( Results );
		}
		Connection.Close();
		CCommon.Print(CAbc.EMPTY,CAbc.EMPTY,"Готово. Для продолжения нажмите Enter.");
		CConsole.ClearKeyboard();
		CConsole.Flash();
		CConsole.ReadChar();
		return	Result;
	}
}