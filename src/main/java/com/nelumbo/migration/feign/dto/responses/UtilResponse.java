package com.nelumbo.migration.feign.dto.responses;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.io.ByteArrayOutputStream;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class UtilResponse {
    private ByteArrayOutputStream byteArrayOutputStream;
    private Integer success;
    private Integer failure;
}
