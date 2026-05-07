    DEF VAR claartRappi AS CHARACTER.
    DEF VAR desartRappi AS CHARACTER.
    DEF VAR consecutivo AS INTEGER INITIAL 1.
    OUTPUT TO "c:\luber.csv".

    FOR EACH articulos WHERE desart BEGINS "uber" USE-INDEX desart:
        EXPORT DELIMITER "," claart desart precios[1].
        
        ASSIGN
            claartRappi = "RPI" + STRING(consecutivo, "999")
            desartRappi = "RAPPI " + SUBSTRING(desart, 6, LENGTH(desart))
            consecutivo = consecutivo + 1.

        EXPORT DELIMITER "," claartRappi desartRappi precios[1].
    END.

    

    OUTPUT CLOSE.
