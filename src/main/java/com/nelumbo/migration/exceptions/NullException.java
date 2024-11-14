package com.nelumbo.migration.exceptions;

public class NullException extends RuntimeException {
    public NullException() {

    }
    public NullException(String message) {
        super(message);
    }
}
