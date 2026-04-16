package com.hoffnungland.poi.corner.dbxlsreport;

/**
 * Exception thrown for worksheet related failures.
 *
 * @author manuel.m.speranza
 * @since 2016-09-01
 */
public class XlsWrkSheetException extends Exception {

    private static final long serialVersionUID = 7418271751314818241L;

    /**
     * Creates a new exception with a human-readable message.
     *
     * @param message error details
     */
    public XlsWrkSheetException(String message) {
        super(message);
    }
}
