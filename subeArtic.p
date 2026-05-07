DEF VAR clave AS CHARACTER.
DEF VAR descripcion AS CHARACTER.
DEF VAR precio AS CHARACTER.
DEF VAR num AS CHARACTER.

INPUT FROM c:\listap.csv.

REPEAT :
    IMPORT DELIMITER "," clave descripcion precio num.
    FIND FIRST artic WHERE artic.claart = clave USE-INDEX claart NO-ERROR.
    IF NOT AVAILABLE artic THEN DO:
        CREATE artic.
        UPDATE
            artic.claart = clave
            artic.desart = descripcion
            artic.precio = dec(precio)
            artic.numcat = int(num).
    END.
END.

INPUT CLOSE.
