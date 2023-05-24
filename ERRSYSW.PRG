#include "error.ch"
#include "FiveWin.ch"

external _fwGenError   // Link FiveWin generic Error Objects Generator

#define NTRIM(n)    ( LTrim( Str( n ) ) )

proc ErrorSys()
    ErrorBlock( { | e | ErrorDialog( e ) } )
return

proc ErrorLink()
return


static function ErrorDialog( e ) // -> logical or \quits App.

   local oDlg, oLbx, oFont
   local lRet    // if lRet == nil -> default action: QUIT
   local n, j, cMessage, aStack := {}
   local oSay, hLogo
   local nButtons  := 1
   local cErrorLog := ""
   local aVersions := GetVersion()
   local aTasks    := GetTasks()
   local aRDDs, nTarget, uValue
   local oOldError
   local cRelation
   local lIsWinNT := IsWinNT()

   cErrorLog += "Application" + CRLF
   cErrorLog += "===========" + CRLF
   cErrorLog += "   Path and name: " + GetModuleFileName( GetInstance() ) + CRLF
   cErrorLog += "   Size: " + Transform( FSize( GetModuleFileName( ;
                GetInstance() ) ), "9,999,999 bytes" ) + CRLF
   cErrorLog += "   Max files handles permited: ( SetHandleCount() ) " + ;
                Str( SetHandleCount(), 3 ) + CRLF
   cErrorLog += "   Error ocurred at: " + ;
                DToC( Date() ) + ", " + Time() + CRLF

   // Error object analysis
   cMessage   = "   Error description: " + ErrorMessage( e ) + CRLF
   cErrorLog += cMessage

   if ValType( e:Args ) == "A"
      cErrorLog += "   Args:" + CRLF
      for n = 1 to Len( e:Args )
         cErrorLog += "     [" + Str( n, 4 ) + "] = " + ValType( e:Args[ n ] ) + ;
                      "   " + cValToChar( e:Args[ 1 ] ) + CRLF
      next
   endif

   cErrorLog += CRLF + "Stack Calls" + CRLF
   cErrorLog += "===========" + CRLF
      n := 2    // we don't disscard any info again !
      while ( n < 74 )
          if ! Empty(ProcName( n ) )
             AAdd( aStack, "   Called from " + Trim( ProcName( n ) ) + ;
                           "(" + NTRIM( ProcLine( n ) ) + ")" )
             cErrorLog += ATail( aStack ) + CRLF
          endif
          n++
    end

   cErrorLog += CRLF + "System" + CRLF
   cErrorLog += "======" + CRLF

   cErrorLog += "   CPU type: " + GetCPU() + CRLF
   cErrorLog += "   CPU type: " + { "386", "486", "Pentium" }[ CPUType() - 2 ] + CRLF
   cErrorLog += "   Hardware memory: " + cValToChar( nExtMem() + 768 ) + ;
                " (" + cValToChar( Int( ( nExtMem() + 768 ) / 1024 ) ) + ;
                " megas)" + CRLF + CRLF

   cErrorLog += "   Free System resources: % " + AllTrim( Str( GetFreeSystemResources( 0 ) ) ) + CRLF + ;
                "        GDI    resources: % " + AllTrim( Str( GetFreeSystemResources( 1 ) ) ) + CRLF + ;
                "        User   resources: % " + AllTrim( Str( GetFreeSystemResources( 2 ) ) ) + CRLF + CRLF

   cErrorLog += "   Windows and MsDos versions: " + ;
                Str( aVersions[ 1 ], 2 ) + "." + ;
                Str( aVersions[ 2 ], 2 ) + ", " + ;
                Str( aVersions[ 3 ], 2 ) + "." + ;
                Str( aVersions[ 4 ], 2 ) + CRLF

   cErrorLog += "   Windows total applications running: " + Str( GetNumTasks(), 3 ) + CRLF
   for n = 1 to Len( aTasks )
      cErrorLog += "    " + Str( n, 3 ) + " " + aTasks[ n ] + CRLF
   next

   // by default, division by zero yields zero
   if ( e:genCode == EG_ZERODIV )
      return 0
   end

   // for network open error, set NETERR() and subsystem default
   if ( e:genCode == EG_OPEN .and. ;
      ( e:osCode == 32 .or. e:osCode == 5 ) .and. ;
        e:canDefault )
      NetErr( .t. )
      return .f.       // Warning: Exiting!
   end

   // for lock error during APPEND BLANK, set NETERR() and subsystem default
   if ( e:genCode == EG_APPENDLOCK .and. e:canDefault )
      NetErr( .t. )
      return .f.       // OJO SALIDA
   endif

   if ProcName( 7 ) == "ERRORDIALO"   // recursive error !!!
      SET RESOURCES TO
      ErrorLevel( 1 )

      QUIT
   endif

   // Warning!!! Keep here this code !!! Or we will be consuming GDI as
   // we don't generate the error but we were generating the bitmap

   hLogo = FWBitMap()

   if e:canRetry
      nButtons++
   endif

   if e:canDefault
      nButtons++
   endif

   cErrorLog += CRLF + "Variables in use" + CRLF + "================" + CRLF
   cErrorLog += "   Procedure     Type   Value" + CRLF
   cErrorLog += "   ==========================" + CRLF

   n := 2    // we don't disscard any info again !
   while ( n < 74 )

       if ! Empty( ProcName( n ) )
          cErrorLog += "   " + Trim( ProcName( n ) ) + CRLF
          for j = 1 to ParamCount( n )
             cErrorLog += "     Param " + Str( j, 3 ) + ":    " + ;
                          ValType( GetParam( n, j ) ) + ;
                          "    " + cGetInfo( GetParam( n, j ) ) + CRLF
          next
          for j = 1 to LocalCount( n )
             cErrorLog += "     Local " + Str( j, 3 ) + ":    " + ;
                          ValType( GetLocal( n, j ) ) + ;
                          "    " + cGetInfo( GetLocal( n, j ) ) + CRLF
          next
       endif

       n++
   end

   cErrorLog += CRLF + "Linked RDDs" + CRLF + "===========" + CRLF
   aRDDs = RddList( 1 )
   for n = 1 to Len( aRDDs )
      cErrorLog += "   " + aRDDs[ n ] + CRLF
   next

   cErrorLog += CRLF + "DataBases in use" + CRLF + "================" + CRLF
   for n = 1 to 255
      if ! Empty( Alias( n ) )
         cErrorLog += CRLF + Str( n, 3 ) + ": " + If( Select() == n,"=> ", "   " ) + ;
                      PadR( Alias( n ), 15 ) + Space( 20 ) + "RddName: " + ;
                      ( Alias( n ) )->( RddName() ) + CRLF
         cErrorLog += "     ==============================" + CRLF
         cErrorLog += "     RecNo    RecCount    BOF   EOF" + CRLF
         cErrorLog += "    " + Transform( ( Alias( n ) )->( RecNo() ), "99999" ) + ;
                      "      " + Transform( ( Alias( n ) )->( RecCount() ), "99999" ) + ;
                      "      " + cValToChar( ( Alias( n ) )->( BoF() ) ) + ;
                      "   " + cValToChar( ( Alias( n ) )->( EoF() ) ) + CRLF + CRLF
         cErrorLog += "     Indexes in use " + Space( 23 ) + "TagName" + CRLF
         for j = 1 to 15
            if ! Empty( ( Alias( n ) )->( IndexKey( j ) ) )
               cErrorLog += Space( 8 ) + ;
                            If( ( Alias( n ) )->( IndexOrd() ) == j, "=> ", "   " ) + ;
                            PadR( ( Alias( n ) )->( IndexKey( j ) ), 35 ) + ;
                            ( Alias( n ) )->( OrdName( j ) ) + ;
                            CRLF
            endif
         next
         cErrorLog += CRLF + "     Relations in use" + CRLF
         for j = 1 to 8
            if ! Empty( ( nTarget := ( Alias( n ) )->( DbRSelect( j ) ) ) )
               cErrorLog += Space( 8 ) + Str( j ) + ": " + ;
                            "TO " + ( Alias( n ) )->( DbRelation( j ) ) + ;
                            " INTO " + Alias( nTarget ) + CRLF
               // uValue = ( Alias( n ) )->( DbRelation( j ) )
               // cErrorLog += cValToChar( &( uValue ) ) + CRLF
            endif
         next
      endif
   next

   n = 1
   cErrorLog += CRLF + "Classes in use:" + CRLF
   cErrorLog += "===============" + CRLF
   while ! Empty( __ClassNam( n ) )
      cErrorLog += "   " + Str( n, 3 ) + " " + __ClassNam( n++ ) + CRLF
   end

   cErrorLog += CRLF + "Memory Analysis" + CRLF
   cErrorLog +=        "===============" + CRLF
   cErrorLog += "   Static memory:" + CRLF
   cErrorLog += "      data segment: 64k" + CRLF
   cErrorLog += "      Initial size:       " + ;
                LTrim( Str( nInitDSSize() ) ) + ;
                " bytes  (SYMP=" + LTrim( Str( nSymPSize() ) ) + ;
                ", Stack=" + LTrim( Str( nStackSize() ) ) + ;
                ", Heap=" + LTrim( Str( nHeapSize() ) ) + ")" + CRLF
   cErrorLog += "      Clipper Stack:      " + ;
                LTrim( Str( 65535 - ( nStatics() * 14 ) - nInitDSSize() ) ) + ;
                " bytes" + CRLF
   cErrorLog += "      " + LTrim( Str( nStatics() ) ) + " Static variables: " + ;
                LTrim( Str( nStatics() * 14 ) ) + " bytes" + CRLF + CRLF

   cErrorLog += "   Dynamic memory consume:" + CRLF
   cErrorLog += "      Actual  Value: " + Str( MemUsed() ) + " bytes" + CRLF
   cErrorLog += "      Highest Value: " + Str( MemMax() ) + " bytes" + CRLF
   nSymNames() // no longer returns a real value! 15/April/97
   cErrorLog += "   SYMBOLS segment" + CRLF
   cErrorLog += "      " + LTrim( Str( nSymNames() ) ) + " SymbolNames:   " + ;
                LTrim( Str( nSymNames() * 16 ) ) + " bytes"
   // Generates a file with an Error Log

   BEGIN SEQUENCE
      oOldError = ErrorBlock( { || DoBreak() } )
      MemoWrit( "Error.log", cErrorLog )
   END SEQUENCE
   ErrorBlock( oOldError )

   MSGINFO("Ocorreu uma instrução inválida, o sistema será fechado avise a Exatus, enviando o arquivo ERROR.LOG !",WSIST)
   SET RESOURCES TO
   ErrorLevel( 1 )
   PERGUNTA=.T.
   QUIT              // must be QUIT !!!

return lRet

//----------------------------------------------------------------------------//

static function DoBreak()

   BREAK

return nil

//----------------------------------------------------------------------------//

static func ErrorMessage( e )

	// start error message
    local cMessage := if( empty( e:OsCode ), ;
                          if( e:severity > ES_WARNING, "Error ", "Warning " ),;
                          "(DOS Error " + NTRIM(e:osCode) + ") " )

	// add subsystem name if available
    cMessage += if( ValType( e:SubSystem ) == "C",;
                    e:SubSystem()                ,;
                    "???" )

	// add subsystem's error code if available
    cMessage += if( ValType( e:SubCode ) == "N",;
                    "/" + NTRIM( e:SubCode )   ,;
                    "/???" )
	// add error description if available
  if ( ValType( e:Description ) == "C" )
        cMessage += "  " + e:Description
	end

	// add either filename or operation
    cMessage += if( ! Empty( e:FileName ),;
                    ": " + e:FileName   ,;
                    if( !Empty( e:Operation ),;
                        ": " + e:Operation   ,;
                        "" ) )
return cMessage

//----------------------------------------------------------------------------//
// returns extended info for a certain variable type

static function cGetInfo( uVal )

   local cType := ValType( uVal )

   do case
      case cType == "C"
           return '"' + cValToChar( uVal ) + '"'

      case cType == "O"
           return "Class: " + uVal:ClassName()

      case cType == "A"
           return "Len: " + Str( Len( uVal ), 4 )

      otherwise
           return cValToChar( uVal )
   endcase

return nil

//----------------------------------------------------------------------------//
