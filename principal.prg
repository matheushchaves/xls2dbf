#include "Fivewin.ch"

Function Main()
			Local oDlgPrincipal
			Local vcampo_destino,vcampo_origem,nprogresso
   REQUEST DBFCDX
   RDDSETDEFAULT("DBFCDX")

			*
			Store SPACE(100) to vcampo_destino,vcampo_origem
			nprogresso := 0
			*
         Define Dialog oDlgPrincipal RESOURCE "DLG_XLS2DBF" Title "XLS2DBF"
         
         REDEFINE GET ocampo_origem  VAR vcampo_origem  ID 4001 OF oDlgPrincipal ACTION pegaArquivo(@vcampo_origem,ocampo_origem)
         REDEFINE GET ocampo_destino VAR vcampo_destino ID 4003 OF oDlgPrincipal ACTION pegaArquivo(@vcampo_destino,ocampo_destino)
         
         REDEFINE METER oprogresso VAR nprogresso ID 4005 OF oDlgPrincipal TOTAL 100
         
         REDEFINE BUTTON obtnIniciar ID 4007 OF oDlgPrincipal ACTION converte(vcampo_destino,vcampo_origem,nprogresso,oprogresso,oDlgPrincipal) WHEN !Empty(vcampo_origem) .AND. !Empty(vcampo_destino)
         
         REDEFINE BUTTON obtnSair    ID 4008 OF oDlgPrincipal ACTION oDlgPrincipal:End()
         
         Activate Dialog oDlgPrincipal centered
         
         __quit()
Function pegaArquivo(vcaminho,ocaminho)         
			local xcaminho := ""
			
			xcaminho := cGetFile("*.*","Informe o arquivo")
			
			if !Empty(xcaminho)
				vcaminho := xcaminho
				ocaminho:refresh()
			endif	
*         
Function Converte(vcampo_destino,vcampo_origem,nprogresso,oprogresso,oDlgPrincipal)

        	* VALID EXTENSAO XLS
			IF UPPER(cFileExt(vcampo_origem)) # "XLS" 
			   IF UPPER(cFileExt(vcampo_origem)) # "XLSX"
	  		      MsgAlert("No Arquivo de Origem não é aceito a extensão "+UPPER(cFileExt(vcampo_origem)),"Atenção")
					RETURN .F.
			   ENDIF
			ENDIF

			IF UPPER(cFileExt(vcampo_destino)) # "DBF" 
  		      MsgAlert("No Arquivo de Destino não é aceito a extensão "+UPPER(cFileExt(vcampo_destino)),"Atenção")
				RETURN .F.
			ENDIF
			
			hColunaCabeca := {=>}
	 		lvolta:=.f.
		   TRY		
			   oExcel := CreateObject( "Excel.Application" )
   			oBook := oExcel:WorkBooks:Open( vcampo_origem,;
                                      OleDefaultArg(), ;
                                      OleDefaultArg(), ;
                                      OleDefaultArg(), ;
                                      OleDefaultArg(), ;
                                      '1111')
   			oSheet := oExcel:Get("ActiveSheet")
   		CATCH oerro
     				MsgAlert("Atenção não é possível abrir a planilha - erro Tecnico! - "+oerro:description,"Alerta")
     				lvolta:=.t.
  			END
  			IF lvolta
  			  RETURN .T.
  			ENDIF
		   oDlgPrincipal:SetText("XLS2DBF - Aguarde um momento ...")
		   nTotalLinhas := 0
			nColuna := 0
			WHILE .T.
			   nColuna++
				cColuna       := oSheet:Cells(1,nColuna):Value 
			   if !Empty(cColuna)
					HSet(hColunaCabeca,alltrim(cValToChar(cColuna)),nColuna)
				ELSE
				   EXIT
				ENDIF	
			END	

			WHILE .T.
			   nTotalLinhas++
				cCampo       := oSheet:Cells(nTotalLinhas,1):Value 
			   if Empty(cCampo)
				   nTotalLinhas--
					EXIT
				ENDIF	
			END	
			oprogresso:SetTotal(nTotalLinhas)
			oprogresso:refresh()
			
			DbUseArea(.T.,"DBFCDX",VCAMPO_DESTINO,"ARQ",.F.,.F.,NIL,NIL)
			IF NetErr()
			   MsgAlert("Atenção o arquivo não pode está em uso !","Alerta")
			   CLOSE DATA
			   RETURN .T.
			ENDIF
			aEstrutura:= DbStruct()
			For nLinha:=2 to nTotalLinhas
			    select arq
			    append blank
				 For nColunaDBF := 1 to len(aEstrutura)
			        cColuna := aEstrutura[nColunaDBF][1]
					  cTipo := aEstrutura[nColunaDBF][2]
					  nColunaEXCEL:=0
					  TRY
					    nColunaEXCEL:=hColunaCabeca[cColuna]
					  CATCH
					  END 
		   		  IF nColunaEXCEL > 0
					  	  cValor  := oSheet:Cells(nLinha,nColunaEXCEL):Value 			
						  IF cTipo == "C"
					  	     replace &cColuna WITH ALLTRIM(cValToChar(cValor))
						  ENDIF
						  IF cTipo == "N"
						  	  replace &cColuna WITH VAL(cValToChar(cValor))
						  ENDIF
					  	  IF cTipo == "L"
							  replace &cColuna WITH IIF(cValToChar(cValor) == '1',.T.,.F.)
						  ENDIF
					  ENDIF
					  oprogresso:set(nLinha)
					  oprogresso:refresh()
					  SysRefresh()
				next	  
			next	  
		   TRY
		      oExcel:Quit()
		   catch
		   end
		   Release oExcel
			close data		  
			
			
			   
			
			
			  
			     
         