    OUTPUT TO c:\roscas.csv.

    FOR EACH articulos WHERE desart CONTAINS('Rosca') :
    EXPORT DELIMITER "," claart desart precios[1] numcat.
    END.

    OUTPUT CLOSE.
