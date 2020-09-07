*****************************************************************************
* ARQUIVO DE PROCEDIMENTOS
* PROGRAMADOR - ADILSON
* DRH / SPR
*****************************************************************************
FUNCTION CONFIRMA
PARAMETERS M->WMEN, M->WLIN
@ M->WLIN,00 SAY SPACE ( 80 )
MENSAG(M->WMEN,M->WLIN)
M->TEC = 0
DO WHILE M->TEC # 83 .AND. M->TEC # 115 .AND. M->TEC # 78 .AND. M->TEC # 110
   SET CURSOR OFF
   M->TEC = INKEY(0)
   SET CURSOR ON
   IF M->TEC = 83 .OR. M->TEC = 115
      RETURN .T.
   ELSEIF M->TEC = 78 .OR. M->TEC = 110
      RETURN .F.
   ENDIF
ENDDO
*****************************************************************************
FUNCTION QUADRO
PARAMETERS M->L1, M->C1, M->L2, M->C2, OP
M->WBORDA=CHR(201)+CHR(205)+CHR(187)+CHR(186)+CHR(188)+CHR(205)+CHR(200)+CHR(186)+CHR(255)
IF OP = 'SIM'
   @ M->L1+1, M->C1-1, M->L2+1, M->C2-1 BOX M->WBORDA
ENDIF
@ M->L1, M->C1, M->L2, M->C2 BOX M->WBORDA
RETURN .T.
*****************************************************************************
FUNCTION MENSAG
PARAMETERS M->WMEN, M->WLIN
@ M->WLIN,00 SAY SPACE ( 80 )
CTR =  (LEN(M->WMEN))
FOR I = 1 TO CTR
   @ M->WLIN,(INT(80-(I))/2) SAY SUBSTR (M->WMEN,1,I)
NEXT
RETURN .T.
*****************************************************************************
FUNCTION IMP
PARAMETERS M->WLIN, WIMP
MENSAG ('INFORME A SAIDA',24)
M->WTELIMP = SAVESCREEN (20,74,24,79)
QUADRO (20,74,24,79,'NAO')
@ 21,75 PROMPT 'LPT1'
@ 22,75 PROMPT 'LPT2'
@ 23,75 PROMPT 'LPT3'
MENU TO WOPIMP
IF M->WOPIMP = 0 .OR. M->WOPIMP = 1
   M->WIMP = 'LPT1'
ELSEIF M->WOPIMP = 2
   M->WIMP = 'LPT2'
ELSEIF M->WOPIMP = 3
   M->WIMP = 'LPT3'
ENDIF
RESTSCREEN (20,74,24,79,M->WTELIMP)
IF CONFIRMA ('IMPRESSORA OK ? <S/N>',M->WLIN)
   IF ISPRINTER()
      MENSAG ('',M->WLIN)
      RETURN .T.
   ELSE
      MENSAG ('IMPRESSORA NAO ESTA CONECTADA !',M->WLIN)
      INKEY(1)
      RETURN .F.
   ENDIF
ELSE
   RETURN .F.
ENDIF
*****************************************************************************
FUNCTION CALC
   PRIVATE CALC_COL, CALC_LIN, CLIN, CORRENTE, ATUAL, DECIMAL, ULT_CHAR, ;
           OPERADOR, PRIMEIRO, TECLA_PRESS, DP,  CALC_CHAR, MOVE_TECLAS, ;
           CONT_OPERADOR, TL_CALC, TL_ANT
   CALC_LIN = 5
   CALC_COL = 53
   MOVE_TECLAS = CHR(04) + CHR(19) + CHR(05) + CHR(24)
   TAM_MAX_NUM = 19
   CLIN = 0
   TL_ANT = SAVESCREEN(CALC_LIN, CALC_COL, CALC_LIN + 15, CALC_COL + 23)
   @ CALC_LIN, CALC_COL TO CALC_LIN + 15, CALC_COL + 23  && MOLDURA.
   CLIN = CLIN + 1
   @ CALC_LIN + CLIN, CALC_COL + 1 TO CALC_LIN + CLIN + 3, CALC_COL + 22
   CLIN = CLIN + 1
   @ CALC_LIN + CLIN, CALC_COL + 2 SAY SPACE(20)
   CLIN = CLIN + 1
   @ CALC_LIN + CLIN, CALC_COL + 2 SAY SPACE(20)
   CLIN = CLIN + 2
   @ CALC_LIN + CLIN, CALC_COL + 1 SAY '  C    cE   Sr' + '     ' + CHR(246) + '/ '
   @ CALC_LIN + CLIN, CALC_COL + 5 SAY '³'
   @ CALC_LIN + CLIN, CALC_COL + 10 SAY '³'
   @ CALC_LIN + CLIN, CALC_COL + 15 SAY '³'
   @ CALC_LIN + CLIN, CALC_COL + 22 SAY '³'
   CLIN = CLIN + 1
   @ CALC_LIN + CLIN, CALC_COL + 1 SAY  ' ÍÍÍ¾ ÍÍÍ¾ ÍÍÍ¾   ÍÍÍ¾'
   CLIN = CLIN + 1
   @ CALC_LIN + CLIN, CALC_COL + 1 SAY  '  7    8    9      *  '
   @ CALC_LIN + CLIN, CALC_COL + 5 SAY '³'
   @ CALC_LIN + CLIN, CALC_COL + 10 SAY '³'
   @ CALC_LIN + CLIN, CALC_COL + 15 SAY '³'
   @ CALC_LIN + CLIN, CALC_COL + 22 SAY '³'
   CLIN = CLIN + 1
   @ CALC_LIN + CLIN, CALC_COL + 1 SAY  ' ÍÍÍ¾ ÍÍÍ¾ ÍÍÍ¾   ÍÍÍ¾'
   CLIN = CLIN + 1
   @ CALC_LIN + CLIN, CALC_COL + 1 SAY  '  4    5    6      +  '
   @ CALC_LIN + CLIN, CALC_COL + 5 SAY '³'
   @ CALC_LIN + CLIN, CALC_COL + 10 SAY '³'
   @ CALC_LIN + CLIN, CALC_COL + 15 SAY '³'
   @ CALC_LIN + CLIN, CALC_COL + 22 SAY '³'
   CLIN = CLIN + 1
   @ CALC_LIN + CLIN, CALC_COL + 1 SAY  ' ÍÍÍ¾ ÍÍÍ¾ ÍÍÍ¾   ÍÍÍ¾'
   CLIN = CLIN + 1
   @ CALC_LIN + CLIN, CALC_COL + 1 SAY  '  1    2    3      -  '
   @ CALC_LIN + CLIN, CALC_COL + 5 SAY '³'
   @ CALC_LIN + CLIN, CALC_COL + 10 SAY '³'
   @ CALC_LIN + CLIN, CALC_COL + 15 SAY '³'
   @ CALC_LIN + CLIN, CALC_COL + 22 SAY '³'
   CLIN = CLIN + 1
   @ CALC_LIN + CLIN, CALC_COL + 1 SAY  ' ÍÍÍ¾ ÍÍÍ¾ ÍÍÍ¾   ÍÍÍ¾'
   CLIN = CLIN + 1
   @ CALC_LIN + CLIN, CALC_COL + 1 SAY  '  0    .    Y' + CHR(252) + '     =  '
   @ CALC_LIN + CLIN, CALC_COL + 5 SAY '³'
   @ CALC_LIN + CLIN, CALC_COL + 10 SAY '³'
   @ CALC_LIN + CLIN, CALC_COL + 15 SAY '³'
   @ CALC_LIN + CLIN, CALC_COL + 22 SAY '³'
   CLIN = CLIN + 1
   @ CALC_LIN + CLIN, CALC_COL + 1 SAY  ' ÍÍÍ¾ ÍÍÍ¾ ÍÍÍ¾   ÍÍÍ¾'
   TL_CALC = SAVESCREEN(CALC_LIN, CALC_COL, CALC_LIN + 15, CALC_COL + 23)
   STORE 0 TO TECLA_PRESS, DP
   STORE 0.0000 TO CORRENTE, ATUAL
   OPERADOR = ' '
   ULT_CHAR = 'C'
   DECIMAL = .F.
   PRIMEIRO = .T.
   CONT_OPERADOR = 0
   ALGARISMO = .F.
   DO WHILE TECLA_PRESS != 27 && FACA ATE' QUE ESC SEJA PRESSIONADA.
      @ CALC_LIN + 3, CALC_COL + 2 SAY CORRENTE  PICT '99999999999999.9999'
      TECLA_PRESS = 0
      TECLA_PRESS = INKEY(0)
      IF TECLA_PRESS = 13
         TECLA_PRESS = 61
      ENDIF
      CALC_CHAR = UPPER(CHR(TECLA_PRESS))
      IF CALC_CHAR $ '+-/*Y'
         ALGARISMO = .F.
         IF CONT_OPERADOR = 0
            CONT_OPERADOR = 1
         ELSE
            OPERADOR = CALC_CHAR
            LOOP
         ENDIF
      ELSE
         CONT_OPERADOR = 0
      ENDIF
      DO CASE
         CASE CALC_CHAR $ MOVE_TECLAS
            IF TECLA_PRESS = 04
               RESTSCREEN(CALC_LIN, CALC_COL, CALC_LIN + 15, CALC_COL + 23, TL_ANT)
               CALC_COL = CALC_COL + 1
               IF CALC_COL + 23 > 78
                  CALC_COL = 1
               ENDIF
               TL_ANT = SAVESCREEN(CALC_LIN, CALC_COL, ;
                          CALC_LIN + 15, CALC_COL + 23 )
               RESTSCREEN(CALC_LIN, CALC_COL, CALC_LIN + 15, CALC_COL + 23, TL_CALC)
            ELSEIF TECLA_PRESS = 19
               RESTSCREEN(CALC_LIN, CALC_COL, CALC_LIN + 15, CALC_COL + 23, TL_ANT)
               CALC_COL = CALC_COL - 1
               IF CALC_COL < 1
                  CALC_COL = 78 - 23
               ENDIF
               TL_ANT = SAVESCREEN(CALC_LIN, CALC_COL, CALC_LIN + 15, CALC_COL + 23 )
               RESTSCREEN(CALC_LIN, CALC_COL, CALC_LIN + 15, CALC_COL + 23, TL_CALC)
            ELSEIF TECLA_PRESS = 05
               RESTSCREEN(CALC_LIN, CALC_COL, CALC_LIN + 15, CALC_COL + 23, TL_ANT)
               CALC_LIN = CALC_LIN - 1
               IF CALC_LIN < 1
                  CALC_LIN = 24 - 15
               ENDIF
               TL_ANT = SAVESCREEN(CALC_LIN, CALC_COL, CALC_LIN + 15, CALC_COL + 23 )
               RESTSCREEN(CALC_LIN, CALC_COL, CALC_LIN + 15, CALC_COL + 23, TL_CALC)
            ELSEIF TECLA_PRESS = 24
               RESTSCREEN(CALC_LIN, CALC_COL, CALC_LIN + 15, CALC_COL + 23, TL_ANT)
               CALC_LIN = CALC_LIN + 1
               IF CALC_LIN + 15 > 24
                  CALC_LIN = 1
               ENDIF
               TL_ANT = SAVESCREEN(CALC_LIN, CALC_COL, CALC_LIN + 15, CALC_COL + 23 )
               RESTSCREEN(CALC_LIN, CALC_COL, CALC_LIN + 15, CALC_COL + 23, TL_CALC)
            ENDIF
            IF !ALGARISMO
               CONT_OPERADOR = 1
            ENDIF
         CASE CALC_CHAR = 'E'
            CORRENTE = 0.0000
            DP = 0
         CASE CALC_CHAR = 'C'
            ULT_CHAR = CALC_CHAR
            CORRENTE = 0.0000
            ATUAL = 0.0000
         CASE CALC_CHAR = '='
            IF ATUAL = 0.0000 .AND. OPERADOR != 'Y'
               ATUAL = CORRENTE
               OPERADOR = ' '
            ENDIF
            CALC_MATH()
            ULT_CHAR = '='
         CASE CALC_CHAR = 'S'
            IF ULT_CHAR = '='
               STORE SQRT(CORRENTE) TO CORRENTE, ATUAL
            ELSE
               CORRENTE = SQRT(CORRENTE)
            ENDIF
         CASE CALC_CHAR $ '+-/*Y'
            IF ULT_CHAR $ '=C'  && "IGUAL" OU CLEAR
               IF ULT_CHAR = 'C'  && CLEAR
                  ATUAL = CORRENTE
               ENDIF
               ULT_CHAR = ' '
               PRIMEIRO = .T.
            ELSE
               CALC_MATH()
            ENDIF
            STORE CALC_CHAR TO OPERADOR, ULT_CHAR
            DP = 0
         CASE CALC_CHAR = '.'
            DECIMAL = .T.
         CASE CALC_CHAR $ '0123456789'
            ALGARISMO = .T.
            GET_CORRENTE()
      ENDCASE
   ENDDO
   RESTSCREEN(CALC_LIN, CALC_COL, CALC_LIN + 15, CALC_COL + 23, TL_ANT)
RETURN .T.

FUNCTION GET_CORRENTE
   IF DECIMAL
      IF PRIMEIRO   &&  NENHUM NUMERO `A ESQUERDA DO PONTO DECIMAL...
         PRIMEIRO = .F.
         CORRENTE = VAL('0.' + CALC_CHAR)
      ELSE
         CORRENTE = VAL(LTRIM(STR(CORRENTE, TAM_MAX_NUM, 0)) + '.' + CALC_CHAR)
      ENDIF
      DECIMAL = .F.
      DP = 1
   ELSE
      IF PRIMEIRO   && CORRENTE VALE 0
         PRIMEIRO = .F.
         CORRENTE = VAL(CALC_CHAR)
      ELSE
         CORRENTE = VAL(LTRIM(STR(CORRENTE, TAM_MAX_NUM, DP)) + CALC_CHAR)
         IF DP > 0
            DP = DP + 1
         ENDIF
      ENDIF
   ENDIF
RETURN .T.

FUNCTION CALC_MATH
   DO CASE
      CASE OPERADOR = '+'
         ATUAL = ATUAL + CORRENTE
      CASE OPERADOR = '-'
         ATUAL = ATUAL - CORRENTE
      CASE OPERADOR = '*'
         ATUAL = ATUAL * CORRENTE
      CASE OPERADOR = '/'
         IF CORRENTE = 0   && DIVISAO POR ZERO! ERRO!!!
            ATUAL = 0
            CORRENTE = 0
            @ CALC_LIN + 3, CALC_COL + 2 SAY '          E R R O!!'
            INKEY(0)
            CLEAR TYPEAHEAD
            KEYBOARD 'C'
         ELSE
            ATUAL = ATUAL / CORRENTE
         ENDIF
      CASE OPERADOR = 'Y'  && EXPONENCIAL.
         ATUAL = ATUAL ^ CORRENTE
   ENDCASE
   CORRENTE = ATUAL
   PRIMEIRO = .T.
   DP = 0
RETURN .T.
*****************************************************************************
PROCEDURE COR
PARAMETERS AREA
M->WTELCOR = SAVESCREEN (00,00,24,79)
SETCOLOR (M->WCOR3)
@ 01,00,23,79  BOX('±±±±±±±±±')
SETCOLOR (M->WCOR1)
@ 00,00 SAY SPACE(80)
@ 00,12 SAY 'TITULO  PRINCIPAL   TITULO PRINCIPAL   TITULO PRINCIPAL'
SETCOLOR (M->WCOR2)
QUADRO (19,45,21,77,'NAO')
@ 20,47 SAY 'SUB-TITULO   SUB-TITULO  VX.X'
SETCOLOR (M->WCOR4+','+M->WCOR5)
QUADRO (08,33,11,47,'NAO')
@ 09,35 SAY '   TELA    '
@ 10,35 SAY 'ILUSTRATIVA'
QUADRO (03,04,09,15,'NAO')
DO WHILE .T.
   MENSAG ('',24)
   @ 04,06 PROMPT 'TITULO  '
   @ 05,06 PROMPT 'SUB-TIT '
   @ 06,06 PROMPT 'FUNDO   '
   @ 07,06 PROMPT 'QUADROS '
   @ 08,06 PROMPT 'REALCADO'
   MENU TO M->OPCOR
   IF M->OPCOR=0
      RESTSCREEN (00,00,24,79,M->WTELCOR)
      MENSAG ('DEVIDO AS ALTERACOES NO PADRAO DE CORES, ESSE PROGRAMA SERA ABORTADO',24)
      INKEY(0)
      SET COLOR TO
      CLEAR
      CANCEL
   ELSEIF M->OPCOR=1
      DO MOSTRACOR WITH AREA,'TITULO'
   ELSEIF M->OPCOR=2
      DO MOSTRACOR WITH AREA,'SUBTIT'
   ELSEIF M->OPCOR=3
      DO MOSTRACOR WITH AREA,'FUNDO'
   ELSEIF M->OPCOR=4
      DO MOSTRACOR WITH AREA,'QUADROS'
   ELSEIF M->OPCOR=5
      DO MOSTRACOR WITH AREA,'REALCADO'
   ENDIF
ENDDO
RETURN

PROCEDURE MOSTRACOR
PARAMETERS M->WAREA,M->WTELA
MENSAG('( ) COR DO CARACTER   (<Í Í>) COR DO FUNDO   <ENTER) GRAVA',24)
SELE M->WAREA
GO TOP
LOCATE FOR TELA = WTELA
M->WPNUM = VAL(SUBSTR(COR,1,2))
M->WFNUM = VAL(SUBSTR(COR,4,1))
DO WHILE .T.
   M->WCOR = STRZERO(M->WPNUM,2,0)+'/'+STR(M->WFNUM,1,0)
   IF M->WTELA = 'TITULO'
      SETCOLOR (M->WCOR)
      @ 00,00 SAY SPACE(80)
      @ 00,12 SAY 'TITULO  PRINCIPAL   TITULO PRINCIPAL   TITULO PRINCIPAL'
   ELSEIF M->WTELA = 'SUBTIT'
      SETCOLOR (M->WCOR)
      QUADRO (19,45,21,77,'NAO')
      @ 20,47 SAY 'SUB-TITULO   SUB-TITULO  VX.X'
   ELSEIF M->WTELA = 'FUNDO'
      M->WTELCOR1 = SAVESCREEN(03,04,09,15)
      M->WTELCOR2 = SAVESCREEN(19,45,21,77)
      M->WTELCOR3 = SAVESCREEN(08,33,11,47)
      SETCOLOR (M->WCOR)
      @ 01,00,23,79  BOX '±±±±±±±±±'
      RESTSCREEN (03,04,09,15,M->WTELCOR1)
      RESTSCREEN (19,45,21,77,M->WTELCOR2)
      RESTSCREEN (08,33,11,47,M->WTELCOR3)
   ELSEIF M->WTELA = 'QUADROS'
      SETCOLOR (M->WCOR)
      M->WCOR4 = M->WCOR
      QUADRO (08,33,11,47,'NAO')
      @ 09,35 SAY '   TELA    '
      @ 10,35 SAY 'ILUSTRATIVA'
      QUADRO (03,04,09,15,'NAO')
      @ 04,06 SAY 'TITULO  '
      @ 05,06 SAY 'SUB-TIT '
      @ 06,06 SAY 'FUNDO   '
      @ 07,06 SAY 'QUADROS '
      @ 08,06 SAY 'REALCADO'
   ELSEIF M->WTELA = 'REALCADO'
      SETCOLOR (','+M->WCOR)
      M->WCOR5 = M->WCOR
      @ 04,06 SAY 'TITULO  '
      @ 05,06 SAY 'SUB-TIT '
      @ 06,06 SAY 'FUNDO   '
      @ 07,06 SAY 'QUADROS '
      M->WREALCADO = 'REALCADO'
      @ 08,06 GET M->WREALCADO
      CLEAR GETS
   ENDIF
   M->WTEC = INKEY(0)
   IF M->WTEC = 5              &&..........SETA PARA CIMA
      M->WPNUM = IF ( M->WPNUM = 15, 0, M->WPNUM + 1 )
   ELSEIF M->WTEC = 24         &&..........SETA PARA BAIXO
      M->WPNUM = IF ( M->WPNUM = 0, 15, M->WPNUM - 1 )
   ELSEIF M->WTEC = 4          &&..........SETA PARA ESQUERDA
      M->WFNUM = IF ( M->WFNUM = 0, 07, M->WFNUM - 1 )
   ELSEIF M->WTEC = 19         &&..........SETA PARA DIREITA
      M->WFNUM = IF ( M->WFNUM = 07, 0, M->WFNUM + 1 )
   ELSEIF M->WTEC = 13
      EXIT
   ENDIF
ENDDO
SETCOLOR (M->WCOR4+','+M->WCOR5)
REPLACE COR WITH M->WCOR
RETURN
*****************************************************************************
FUNCTION CALC_DIG
PARAMETERS WPRONTU
M->WVAL1 = VAL(SUBSTR(STR(M->WPRONTU,7,0),1,1)) * 8
M->WVAL2 = VAL(SUBSTR(STR(M->WPRONTU,7,0),2,1)) * 7
M->WVAL3 = VAL(SUBSTR(STR(M->WPRONTU,7,0),3,1)) * 6
M->WVAL4 = VAL(SUBSTR(STR(M->WPRONTU,7,0),4,1)) * 5
M->WVAL5 = VAL(SUBSTR(STR(M->WPRONTU,7,0),5,1)) * 4
M->WVAL6 = VAL(SUBSTR(STR(M->WPRONTU,7,0),6,1)) * 3
M->WVAL7 = VAL(SUBSTR(STR(M->WPRONTU,7,0),7,1)) * 2
M->WSOMA = M->WVAL1 + M->WVAL2 + M->WVAL3 + M->WVAL4 + M->WVAL5 + M->WVAL6 + M->WVAL7
M->WRESTO = MOD(M->WSOMA,11)
IF M->WRESTO = 0 .OR. M->WRESTO = 1
   RETURN '0'
ELSE
   RETURN STR(11 - M->WRESTO,1,0)
ENDIF
*****************************************************************************
