package com.nelumbo.migration.feign.dto.responses.error;

import java.util.List;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class ErrorDetailResponse {

    private Long id;
    private String code;
    private String description;
    private List<String> fields;

    public ErrorDetailResponse(String code, String description, Long id) {
        this.id = id;
        this.code = code;
        this.description = description;
    }
}
