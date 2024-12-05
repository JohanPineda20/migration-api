package com.nelumbo.migration.exceptions;

import com.nelumbo.migration.feign.dto.responses.error.ErrorDetailResponse;
import com.nelumbo.migration.feign.dto.responses.error.ErrorResponse;
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.http.converter.HttpMessageNotReadableException;
import org.springframework.validation.FieldError;
import org.springframework.web.bind.MethodArgumentNotValidException;
import org.springframework.web.bind.MissingServletRequestParameterException;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.ResponseStatus;
import org.springframework.web.bind.annotation.RestControllerAdvice;
import org.springframework.web.multipart.MaxUploadSizeExceededException;
import org.springframework.web.multipart.MultipartException;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

@RestControllerAdvice
@Slf4j
public class GlobalExceptionHandler {

    private final String VALIDATION_EXCEPTION_CODE = "C03";
    private final String VALIDATION_EXCEPTION = "Validation Exception";

    @ExceptionHandler(MethodArgumentNotValidException.class)
    @ResponseStatus(HttpStatus.BAD_REQUEST)
    public ResponseEntity<ErrorResponse> handleException(MethodArgumentNotValidException e) {
        List<String> errors = new ArrayList<>();
        for (FieldError fieldError : e.getBindingResult().getFieldErrors()) {
            errors.add(fieldError.getDefaultMessage());
        }
        ErrorResponse errorResponse = new ErrorResponse();
        errorResponse.setErrors(new ErrorDetailResponse(VALIDATION_EXCEPTION_CODE, VALIDATION_EXCEPTION, errors));
        return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(errorResponse);
    }

    @ExceptionHandler(HttpMessageNotReadableException.class)
    @ResponseStatus(HttpStatus.BAD_REQUEST)
    public ResponseEntity<ErrorResponse> handleException(HttpMessageNotReadableException e) {
        ErrorResponse errorResponse = new ErrorResponse();
        errorResponse.setErrors(new ErrorDetailResponse(VALIDATION_EXCEPTION_CODE, VALIDATION_EXCEPTION,
                Collections.singletonList(e.getMessage())));
        return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(errorResponse);
    }

    // max upload size exception
    @ExceptionHandler(MaxUploadSizeExceededException.class)
    @ResponseStatus(HttpStatus.BAD_REQUEST)
    public ResponseEntity<ErrorResponse> handleException(MaxUploadSizeExceededException e) {
        ErrorResponse errorResponse = new ErrorResponse();
        errorResponse.setErrors(new ErrorDetailResponse(VALIDATION_EXCEPTION_CODE, VALIDATION_EXCEPTION,
                Collections.singletonList(e.getMessage())));
        log.error("Error processing Excel file: {}", e.getMessage());
        return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(errorResponse);
    }

    @ExceptionHandler(MultipartException.class)
    @ResponseStatus(HttpStatus.BAD_REQUEST)
    public ResponseEntity<ErrorResponse> handleException(MultipartException e) {
        ErrorResponse errorResponse = new ErrorResponse();
        errorResponse.setErrors(new ErrorDetailResponse(VALIDATION_EXCEPTION_CODE, VALIDATION_EXCEPTION,
                Collections.singletonList(e.getMessage())));
        log.error("Error processing Excel file: {}", e.getMessage());
        return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(errorResponse);
    }

    @ExceptionHandler(MissingServletRequestParameterException.class)
    @ResponseStatus(HttpStatus.BAD_REQUEST)
    public ResponseEntity<ErrorResponse> handleException(MissingServletRequestParameterException e) {
        ErrorResponse errorResponse = new ErrorResponse();
        errorResponse.setErrors(new ErrorDetailResponse(VALIDATION_EXCEPTION_CODE, VALIDATION_EXCEPTION,
                Collections.singletonList(e.getMessage())));
        return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(errorResponse);
    }

    @ExceptionHandler(NullException.class)
    @ResponseStatus(HttpStatus.BAD_REQUEST)
    public ResponseEntity<ErrorResponse> handleException(NullException e) {
        ErrorResponse errorResponse = new ErrorResponse();
        errorResponse.setErrors(new ErrorDetailResponse("C01MIG01", VALIDATION_EXCEPTION,
                Collections.singletonList(e.getMessage())));
        log.error("Error processing Excel file: {}", e.getMessage());
        return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(errorResponse);
    }

    @ExceptionHandler(ServiceUnavailableException.class)
    @ResponseStatus(HttpStatus.INTERNAL_SERVER_ERROR)
    public ResponseEntity<ErrorResponse> handleException(ServiceUnavailableException e) {
        ErrorResponse errorResponse = new ErrorResponse();
        errorResponse.setErrors(new ErrorDetailResponse("N01GNRC04", "INTERNAL_EXCEPTION",
                Collections.singletonList(e.getMessage())));
        log.error("Error processing Excel file: {}", e.getMessage());
        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(errorResponse);
    }

    @ExceptionHandler(UnauthorizedException.class)
    @ResponseStatus(HttpStatus.UNAUTHORIZED)
    public ResponseEntity<ErrorResponse> handleException(UnauthorizedException e) {
        ErrorResponse errorResponse = new ErrorResponse();
        errorResponse.setErrors(new ErrorDetailResponse("S01GNRC13", "SECURITY_EXCEPTION",
                Collections.singletonList(e.getMessage())));
        log.error("Error processing Excel file: {}", e.getMessage());
        return ResponseEntity.status(HttpStatus.UNAUTHORIZED).body(errorResponse);
    }

    @ExceptionHandler(ForbbidenException.class)
    @ResponseStatus(HttpStatus.FORBIDDEN)
    public ResponseEntity<ErrorResponse> handleException(ForbbidenException e) {
        ErrorResponse errorResponse = new ErrorResponse();
        errorResponse.setErrors(new ErrorDetailResponse("S01GNRC10", "SECURITY_EXCEPTION",
                Collections.singletonList(e.getMessage())));
        log.error("Error processing Excel file: {}", e.getMessage());
        return ResponseEntity.status(HttpStatus.FORBIDDEN).body(errorResponse);
    }

    @ExceptionHandler(ErrorResponseException.class)
    @ResponseStatus(HttpStatus.INTERNAL_SERVER_ERROR)
    public ResponseEntity<ErrorResponse> handleException(ErrorResponseException e) {
        ErrorResponse errorResponse = new ErrorResponse();
        errorResponse.setErrors(new ErrorDetailResponse(e.getError().getErrors().getCode(), e.getError().getErrors().getDescription(),
                e.getError().getErrors().getFields()));
        log.error("Error processing Excel file: {}", e.getError().getErrors().getFields().toString());
        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(errorResponse);
    }

    @ExceptionHandler(Exception.class)
    @ResponseStatus(HttpStatus.INTERNAL_SERVER_ERROR)
    public ResponseEntity<ErrorResponse> handleException(Exception e) {
        ErrorResponse errorResponse = new ErrorResponse();
        errorResponse.setErrors(new ErrorDetailResponse(VALIDATION_EXCEPTION_CODE, VALIDATION_EXCEPTION,
                Collections.singletonList(e.getMessage())));
        log.error("Error processing Excel file: {}", e.getMessage() + " - " + e.getClass().getName());
        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(errorResponse);
    }
}
