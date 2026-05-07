DEF VAR clave AS CHARACTER.
DEF VAR descripcion AS CHARACTER.
DEF VAR precio AS DECIMAL.
DEF VAR num AS INTEGER.

INPUT FROM c:\listap.csv.
OUTPUT TO c:\cambios.csv.

REPEAT:
      IMPORT DELIMITER "," clave descripcion precio num.

      FIND FIRST artic WHERE artic.claart = clave USE-INDEX claart NO-ERROR.
      IF AVAILABLE artic THEN DO:
          IF artic.precio <> precio THEN DO:
              EXPORT DELIMITER "," artic.claart artic.desart artic.precio.
          END.
      END.
END.

OUTPUT CLOSE.
INPUT CLOSE.

