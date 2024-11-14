package com.nelumbo.migration.exceptions;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.nelumbo.migration.feign.dto.responses.error.ErrorResponse;

import feign.Response;
import feign.codec.ErrorDecoder;

public class CustomErrorDecoder implements ErrorDecoder {

    private final ObjectMapper objectMapper = new ObjectMapper();
    private final ErrorDecoder defaultErrorDecoder = new Default();

    @Override
    public Exception decode(String methodKey, Response response) {
        ErrorResponse error;
        try {
            if (response.status() == 401 || response.status() == 403) return defaultErrorDecoder.decode(methodKey, response);
            error = objectMapper.readValue(response.body().asInputStream(), ErrorResponse.class);
            return new ErrorResponseException(error);
        } catch (Exception e) {
            return defaultErrorDecoder.decode(methodKey, response);
        }
    }
}
