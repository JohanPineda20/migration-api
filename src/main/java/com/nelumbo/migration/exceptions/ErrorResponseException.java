package com.nelumbo.migration.exceptions;

import com.nelumbo.migration.feign.dto.responses.error.ErrorResponse;

public class ErrorResponseException extends RuntimeException {

    private final ErrorResponse error;

    public ErrorResponseException(ErrorResponse error) {
        this.error = error;
    }

    public ErrorResponse getError() {
        return error;
    }
}
