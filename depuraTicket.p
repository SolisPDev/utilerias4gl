FOR EACH ticket WHERE fectic < date("01/01/2024") USE-INDEX fectic:
    MESSAGE numtic cvecaj clacaj fectic.
    FOR EACH ticketd WHERE ticketd.nummov = ticket.nummov USE-INDEX nummov:
/*         MESSAGE ticketd.claart. */
        DELETE ticketd.
    END.
    DELETE ticket.
END.
