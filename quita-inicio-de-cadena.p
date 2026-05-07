INPUT FROM "c:\rappi-platillos.csv".
OUTPUT TO "c:\platillos-rappi.csv".
    DEF VAR nombre AS CHARACTER.
    DEF VAR precio AS DECIMAL.

REPEAT :
    IMPORT DELIMITER "," nombre precio.
    IF SUBSTRING(nombre, 1, 4) = "UBER" THEN
        nombre = SUBSTRING(nombre,6, LENGTH(nombre)).

    EXPORT DELIMITER "," nombre precio.
END.

OUTPUT CLOSE.
INPUT CLOSE.
