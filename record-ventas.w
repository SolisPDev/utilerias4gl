&ANALYZE-SUSPEND _VERSION-NUMBER UIB_v9r12 GUI
&ANALYZE-RESUME
&Scoped-define WINDOW-NAME C-Win
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _DEFINITIONS C-Win 
/*------------------------------------------------------------------------

  File: 

  Description: 

  Input Parameters:
      <none>

  Output Parameters:
      <none>

  Author: 

  Created: 

------------------------------------------------------------------------*/
/*          This .W file was created with the Progress AppBuilder.      */
/*----------------------------------------------------------------------*/

/* Create an unnamed pool to store all the widgets created 
     by this procedure. This is a good default which assures
     that this procedure's triggers and internal procedures 
     will execute in this procedure's storage, and that proper
     cleanup will occur on deletion of the procedure. */

CREATE WIDGET-POOL.

/* ***************************  Definitions  ************************** */

/* Parameters Definitions ---                                           */

/* Local Variable Definitions ---                                       */

DEFINE VARIABLE chExcelApplication      AS COM-HANDLE.
DEFINE VARIABLE chWorkbook              AS COM-HANDLE.
DEFINE VARIABLE chWorksheet             AS COM-HANDLE.
DEFINE VARIABLE chChart                 AS COM-HANDLE.
DEFINE VARIABLE chWorksheetRange        AS COM-HANDLE.
DEFINE VARIABLE iCount                  AS INTEGER.
DEFINE VARIABLE iIndex                  AS INTEGER.
DEFINE VARIABLE iTotalNumberOfOrders    AS INTEGER.
DEFINE VARIABLE iMonth                  AS INTEGER.
DEFINE VARIABLE dAnnualQuota            AS DECIMAL.
DEFINE VARIABLE dTotalSalesAmount       AS DECIMAL.
DEFINE VARIABLE iColumn                 AS INTEGER INITIAL 1.
DEFINE VARIABLE cColumn                 AS CHARACTER.
DEFINE VARIABLE cRange                  AS CHARACTER.


DEFINE TEMP-TABLE t-invpri 
    FIELD CLAINV AS CHARACTER
    FIELD DESART AS CHARACTER
    FIELD EX AS DECIMAL
    INDEX CLAINV CLAINV.




DEFINE SHARED VARIABLE programas AS CHARACTER.

DEFINE VARIABLE V-NOMEMP AS CHARACTER.
DEFINE VARIABLE CONT AS INTEGER.

DEFINE TEMP-TABLE vta-tem 
    FIELD claart AS CHARACTER 
    FIELD desart AS CHARACTER 
    FIELD cant   AS DECIMAL
    FIELD precio AS DECIMAL 
    INDEX CLAART CLAART
    INDEX desart desart.

DEFINE VARIABLE V-IMPORTE AS DECIMAL.
DEFINE VARIABLE V-GT AS DECIMAL.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-PREPROCESSOR-BLOCK 

/* ********************  Preprocessor Definitions  ******************** */

&Scoped-define PROCEDURE-TYPE Window
&Scoped-define DB-AWARE no

/* Name of first Frame and/or Browse and/or first Query                 */
&Scoped-define FRAME-NAME DEFAULT-FRAME

/* Standard List Definitions                                            */
&Scoped-Define ENABLED-OBJECTS v-fecini v-fecfin v-dia b-aceptar BtnDone ~
RECT-48 
&Scoped-Define DISPLAYED-OBJECTS v-fecini v-fecfin v-dia 

/* Custom List Definitions                                              */
/* List-1,List-2,List-3,List-4,List-5,List-6                            */

/* _UIB-PREPROCESSOR-BLOCK-END */
&ANALYZE-RESUME



/* ***********************  Control Definitions  ********************** */

/* Define the widget handle for the window                              */
DEFINE VAR C-Win AS WIDGET-HANDLE NO-UNDO.

/* Definitions of the field level widgets                               */
DEFINE BUTTON b-aceptar 
     LABEL "Aceptar" 
     SIZE 15 BY 1.14 TOOLTIP "Genera Reporte".

DEFINE BUTTON BtnDone DEFAULT 
     LABEL "Salir" 
     SIZE 15 BY 1.14
     BGCOLOR 8 .

DEFINE VARIABLE v-dia AS DATE FORMAT "99/99/9999":U 
     LABEL "Fill 1" 
     VIEW-AS FILL-IN 
     SIZE 19 BY 1 NO-UNDO.

DEFINE VARIABLE v-fecfin AS DATE FORMAT "99/99/9999":U 
     LABEL "Fecha Final" 
     VIEW-AS FILL-IN 
     SIZE 18 BY 1 NO-UNDO.

DEFINE VARIABLE v-fecini AS DATE FORMAT "99/99/9999":U 
     LABEL "Fecha Inicial" 
     VIEW-AS FILL-IN 
     SIZE 18 BY 1 NO-UNDO.

DEFINE RECTANGLE RECT-48
     EDGE-PIXELS 4 GRAPHIC-EDGE  NO-FILL 
     SIZE 46 BY 7.14.


/* ************************  Frame Definitions  *********************** */

DEFINE FRAME DEFAULT-FRAME
     v-fecini AT ROW 2.43 COL 21 COLON-ALIGNED
     v-fecfin AT ROW 4.33 COL 21 COLON-ALIGNED
     v-dia AT ROW 6.24 COL 21 COLON-ALIGNED
     b-aceptar AT ROW 9.1 COL 14
     BtnDone AT ROW 9.1 COL 32
     RECT-48 AT ROW 1.48 COL 3
    WITH 1 DOWN NO-BOX KEEP-TAB-ORDER OVERLAY 
         SIDE-LABELS NO-UNDERLINE THREE-D 
         AT COL 1 ROW 1
         SIZE 50 BY 9.57.


/* *********************** Procedure Settings ************************ */

&ANALYZE-SUSPEND _PROCEDURE-SETTINGS
/* Settings for THIS-PROCEDURE
   Type: Window
   Allow: Basic,Browse,DB-Fields,Window,Query
   Other Settings: COMPILE
 */
&ANALYZE-RESUME _END-PROCEDURE-SETTINGS

/* *************************  Create Window  ************************** */

&ANALYZE-SUSPEND _CREATE-WINDOW
IF SESSION:DISPLAY-TYPE = "GUI":U THEN
  CREATE WINDOW C-Win ASSIGN
         HIDDEN             = YES
         TITLE              = "Concentrado de Ventas"
         COLUMN             = 118.4
         ROW                = 21.19
         HEIGHT             = 9.57
         WIDTH              = 50
         MAX-HEIGHT         = 43.38
         MAX-WIDTH          = 256
         VIRTUAL-HEIGHT     = 43.38
         VIRTUAL-WIDTH      = 256
         CONTROL-BOX        = no
         MIN-BUTTON         = no
         MAX-BUTTON         = no
         RESIZE             = no
         SCROLL-BARS        = no
         STATUS-AREA        = no
         BGCOLOR            = ?
         FGCOLOR            = ?
         KEEP-FRAME-Z-ORDER = yes
         THREE-D            = yes
         MESSAGE-AREA       = no
         SENSITIVE          = yes.
ELSE {&WINDOW-NAME} = CURRENT-WINDOW.
/* END WINDOW DEFINITION                                                */
&ANALYZE-RESUME



/* ***********  Runtime Attributes and AppBuilder Settings  *********** */

&ANALYZE-SUSPEND _RUN-TIME-ATTRIBUTES
/* SETTINGS FOR WINDOW C-Win
  VISIBLE,,RUN-PERSISTENT                                               */
/* SETTINGS FOR FRAME DEFAULT-FRAME
                                                                        */
IF SESSION:DISPLAY-TYPE = "GUI":U AND VALID-HANDLE(C-Win)
THEN C-Win:HIDDEN = no.

/* _RUN-TIME-ATTRIBUTES-END */
&ANALYZE-RESUME

 



/* ************************  Control Triggers  ************************ */

&Scoped-define SELF-NAME C-Win
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL C-Win C-Win
ON END-ERROR OF C-Win /* Concentrado de Ventas */
OR ENDKEY OF {&WINDOW-NAME} ANYWHERE DO:
  /* This case occurs when the user presses the "Esc" key.
     In a persistently run window, just ignore this.  If we did not, the
     application would exit. */
  IF THIS-PROCEDURE:PERSISTENT THEN RETURN NO-APPLY.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL C-Win C-Win
ON WINDOW-CLOSE OF C-Win /* Concentrado de Ventas */
DO:
  /* This event will close the window and terminate the procedure.  */
  APPLY "CLOSE":U TO THIS-PROCEDURE.
  RETURN NO-APPLY.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME b-aceptar
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL b-aceptar C-Win
ON CHOOSE OF b-aceptar IN FRAME DEFAULT-FRAME /* Aceptar */
DO:
   
  
    FOR EACH vta-tem:
        DELETE vta-tem.
    END.

    ASSIGN v-fecini = DATE(v-fecini:SCREEN-VALUE IN FRAME {&FRAME-NAME})
        v-fecfin = DATE(v-fecfin:SCREEN-VALUE IN FRAME {&FRAME-NAME}).

    
    

    RUN concentra.
    RUN enviar_excel.
   
          

     RUN inicio.
/*    HIDE FRAME f-reportes.  */
  ON TAB TAB.
  ON CURSOR-UP CURSOR-UP.
  ON CURSOR-DOWN CURSOR-DOWN.
  ON RETURN RETURN.
  ON CURSOR-RIGHT CURSOR-RIGHT.
  ON CURSOR-LEFT CURSOR-LEFT.
  ON BACK-TAB BACK-TAB.
    

END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&Scoped-define SELF-NAME BtnDone
&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CONTROL BtnDone C-Win
ON CHOOSE OF BtnDone IN FRAME DEFAULT-FRAME /* Salir */
DO:
  &IF "{&PROCEDURE-TYPE}" EQ "SmartPanel" &THEN
    &IF "{&ADM-VERSION}" EQ "ADM1.1" &THEN
      RUN dispatch IN THIS-PROCEDURE ('exit').
    &ELSE
      RUN exitObject.
    &ENDIF
  &ELSE
      APPLY "CLOSE":U TO THIS-PROCEDURE.
  &ENDIF
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


&UNDEFINE SELF-NAME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _CUSTOM _MAIN-BLOCK C-Win 


/* ***************************  Main Block  *************************** */

/* Set CURRENT-WINDOW: this will parent dialog-boxes and frames.        */
ASSIGN CURRENT-WINDOW                = {&WINDOW-NAME} 
       THIS-PROCEDURE:CURRENT-WINDOW = {&WINDOW-NAME}.

/* The CLOSE event can be used from inside or outside the procedure to  */
/* terminate it.                                                        */
ON CLOSE OF THIS-PROCEDURE 
   RUN disable_UI.

/* Best default for GUI applications is...                              */
PAUSE 0 BEFORE-HIDE.

/* Now enable the interface and wait for the exit condition.            */
/* (NOTE: handle ERROR and END-KEY so cleanup code will always fire.    */
MAIN-BLOCK:
DO ON ERROR   UNDO MAIN-BLOCK, LEAVE MAIN-BLOCK
   ON END-KEY UNDO MAIN-BLOCK, LEAVE MAIN-BLOCK:
  RUN enable_UI.
  RUN inicio.
  IF NOT THIS-PROCEDURE:PERSISTENT THEN
    WAIT-FOR CLOSE OF THIS-PROCEDURE.
END.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME


/* **********************  Internal Procedures  *********************** */

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE concentra C-Win 
PROCEDURE concentra :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/

    FOR EACH t-invpri:
        DELETE t-invpri.
    END.

    RUN llena_tempo.


    FOR EACH ticket WHERE ticket.fectic >= v-fecini AND ticket.fectic <= v-fecfin USE-INDEX fectic:
        IF ticket.estado = TRUE THEN DO:
            DISPLAY ticket.fectic @ v-dia WITH FRAME {&FRAME-NAME}.
            FOR EACH ticketd WHERE ticketd.numtic = ticket.numtic USE-INDEX numtic:

                FIND FIRST t-invpri NO-LOCK WHERE t-invpri.clainv = ticketd.claart USE-INDEX clainv NO-ERROR.
                IF AVAILABLE t-invpri THEN DO:
                    FIND vta-tem WHERE vta-tem.claart = t-invpri.clainv USE-INDEX claart NO-ERROR.
                    IF AVAILABLE vta-tem THEN DO:
                        UPDATE
                            vta-tem.cant = vta-tem.cant + ticketd.cant
                            vta-tem.precio = ticketd.precio.
                    END.
                    ELSE DO:
                        CREATE vta-tem.
                        UPDATE
                            vta-tem.claart = t-invpri.clainv
                            vta-tem.cant = vta-tem.cant + ticketd.cant
                            vta-tem.precio = ticketd.precio.
                    END.
                END.
            END.
        END.
    END.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE disable_UI C-Win  _DEFAULT-DISABLE
PROCEDURE disable_UI :
/*------------------------------------------------------------------------------
  Purpose:     DISABLE the User Interface
  Parameters:  <none>
  Notes:       Here we clean-up the user-interface by deleting
               dynamic widgets we have created and/or hide 
               frames.  This procedure is usually called when
               we are ready to "clean-up" after running.
------------------------------------------------------------------------------*/
  /* Delete the WINDOW we created */
  IF SESSION:DISPLAY-TYPE = "GUI":U AND VALID-HANDLE(C-Win)
  THEN DELETE WIDGET C-Win.
  IF THIS-PROCEDURE:PERSISTENT THEN DELETE PROCEDURE THIS-PROCEDURE.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE enable_UI C-Win  _DEFAULT-ENABLE
PROCEDURE enable_UI :
/*------------------------------------------------------------------------------
  Purpose:     ENABLE the User Interface
  Parameters:  <none>
  Notes:       Here we display/view/enable the widgets in the
               user-interface.  In addition, OPEN all queries
               associated with each FRAME and BROWSE.
               These statements here are based on the "Other 
               Settings" section of the widget Property Sheets.
------------------------------------------------------------------------------*/
  DISPLAY v-fecini v-fecfin v-dia 
      WITH FRAME DEFAULT-FRAME IN WINDOW C-Win.
  ENABLE v-fecini v-fecfin v-dia b-aceptar BtnDone RECT-48 
      WITH FRAME DEFAULT-FRAME IN WINDOW C-Win.
  {&OPEN-BROWSERS-IN-QUERY-DEFAULT-FRAME}
  VIEW C-Win.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE enviar_excel C-Win 
PROCEDURE enviar_excel :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEFINE VARIABLE CAJERO AS CHARACTER.
DEFINE VARIABLE V-IMPORTE AS DECIMAL.
DEFINE VARIABLE V-SEL AS CHARACTER.
/* create a new Excel Application object */
CREATE "Excel.Application" chExcelApplication.
/* launch Excel so it is visible to the user */
chExcelApplication:Visible = TRUE.
/* create a new Workbook */
chWorkbook = chExcelApplication:Workbooks:Add("").
/* get the active Worksheet */
chWorkSheet = chExcelApplication:Sheets:Item(1).

/* set the column names for the Worksheet */
chWorkSheet:Columns("A"):ColumnWidth = 10.
chWorkSheet:Columns("B"):ColumnWidth = 30.
chWorkSheet:Columns("C"):ColumnWidth = 10.
chWorkSheet:Columns("D"):ColumnWidth = 10.
chWorkSheet:Columns("E"):ColumnWidth = 10.
chWorkSheet:Columns("F"):ColumnWidth = 8.
chWorkSheet:Columns("G"):ColumnWidth = 8.
chWorkSheet:Columns("H"):ColumnWidth = 8.
chWorkSheet:Columns("I"):ColumnWidth = 8.
chWorkSheet:Columns("J"):ColumnWidth = 8.
chWorkSheet:Columns("K"):ColumnWidth = 8.
chWorkSheet:Columns("L"):ColumnWidth = 8.
chWorkSheet:Columns("M"):ColumnWidth = 8.
chWorkSheet:Columns("N"):ColumnWidth = 8.
chWorkSheet:Columns("O"):ColumnWidth = 8.
chWorkSheet:Columns("P"):ColumnWidth = 8.
chWorkSheet:Columns("Q"):ColumnWidth = 8.
chWorkSheet:Columns("R"):ColumnWidth = 8.
chWorkSheet:Columns("S"):ColumnWidth = 8.
chWorkSheet:Columns("T"):ColumnWidth = 8.
chWorkSheet:Columns("U"):ColumnWidth = 8.
chWorkSheet:Columns("V"):ColumnWidth = 8.
chWorkSheet:Columns("W"):ColumnWidth = 8.
chWorkSheet:Columns("X"):ColumnWidth = 8.
chWorkSheet:Columns("Y"):ColumnWidth = 8.
chWorkSheet:Columns("Z"):ColumnWidth = 8.





chWorkSheet:Range("A1:C1"):Font:Bold = TRUE.
chWorkSheet:Range("A1"):Value = "CHAIRES SERVICIOS GASTRONOMICOS".
chWorkSheet:Range("A2"):Value = "Suc. Himno Nacional".
chWorkSheet:Range("A3"):Value = "".
chWorkSheet:Range("A4"):Value = "CLAVE".
chWorkSheet:Range("B4"):Value = "PRODUCTO".
chWorkSheet:Range("C4"):Value = "CANTIDAD".
chWorkSheet:Range("D4"):Value = "PRECIO".
chWorkSheet:Range("E4"):Value = "IMPORTE".


cColumn = STRING(5).

FOR EACH VTA-TEM USE-INDEX desart, EACH t-invpri NO-LOCK WHERE t-invpri.clainv = vta-tem.claart USE-INDEX CLAINV BY vta-tem.cant DESCENDING :
    
    
     cRange = "A" + cColumn.
     chWorkSheet:Range(cRange):Value = VTA-TEM.CLAART.
     cRange = "B" + cColumn.
     chWorkSheet:Range(cRange):Value = t-invpri.DESART.
     cRange = "C" + cColumn.
     chWorkSheet:Range(cRange):Value = VTA-TEM.CANT.
     cRange = "D" + cColumn.
     chWorkSheet:Range(cRange):Value = VTA-TEM.PRECIO.
     cRange = "E" + cColumn.
     chWorkSheet:Range(cRange):Value = VTA-TEM.CANT * VTA-TEM.PRECIO.
     
     cColumn = STRING(INT(cColumn) + 1).
END.


/* ASSIGN V-SEL = "C4:" + "E" + STRING(INT(cColumn) - 1).  */
/*                                                         */
/* chWorkSheet:Range(V-SEL):Select().                      */
/* chExcelApplication:Selection:Style = "Currency".        */

/*/* create embedded chart using the data in the Worksheet */
 * chWorksheetRange = chWorksheet:Range("A1:C10").
 * /*chWorksheet:ChartObjects:Add(10,150,425,300):Activate.*/
 * /*chExcelApplication:ActiveChart:ChartWizard(chWorksheetRange, 3, 1, 2, 1, 1, TRUE,
 *  *     "1996 Sales Figures", "Sales Person", "Annual Sales").*/
 * 
 * /*/* create chart using the data in the Worksheet */
 *  * chChart=chExcelApplication:Charts:Add().
 *  * chChart:Name = "Test Chart".
 *  * chChart:Type = 11.*/*/

/* release com-handles */
/*RELEASE OBJECT chExcelApplication.      
 * RELEASE OBJECT chWorkbook.
 * RELEASE OBJECT chWorksheet.
 * RELEASE OBJECT chChart.
 * RELEASE OBJECT chWorksheetRange. */

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE inicio C-Win 
PROCEDURE inicio :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
    DISPLAY
        TODAY @ v-fecini
        TODAY @ v-fecfin
    WITH FRAME {&FRAME-NAME}.

    FIND FIRST EMPRESA NO-ERROR.
    V-NOMEMP = EMPRESA.NOMEMP.

END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

&ANALYZE-SUSPEND _UIB-CODE-BLOCK _PROCEDURE llena_tempo C-Win 
PROCEDURE llena_tempo :
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  <none>
  Notes:       
------------------------------------------------------------------------------*/
DEF VAR v-clainv AS CHARACTER.
DEF VAR v-desart AS CHARACTER.

INPUT FROM VALUE("c:\chaires\pasteleria.csv").

REPEAT:
    IMPORT DELIMITER "," v-clainv v-desart.
    
    FIND FIRST ARTICULOS NO-LOCK WHERE articulos.claart = v-clainv USE-INDEX claart NO-ERROR.
    IF AVAILABLE articulos THEN DO:
        
        CREATE t-invpri.
        UPDATE 
            t-invpri.clainv = v-clainv
            t-invpri.desart = articulos.desart.
    END.
END.
END PROCEDURE.

/* _UIB-CODE-BLOCK-END */
&ANALYZE-RESUME

