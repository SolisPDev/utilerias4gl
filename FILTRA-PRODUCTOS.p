OUTPUT TO "c:\lista-rappi.csv".

FOR EACH articulos WHERE numcat = 1 AND ((desart BEGINS "pastel") OR (desart BEGINS "gelatina")) USE-INDEX desart:
    EXPORT DELIMITER "," claart desart precios[1].
END.

OUTPUT CLOSE.


