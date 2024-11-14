package com.nelumbo.migration.feign.dto.responses.error;

import java.util.List;

import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@JsonInclude(JsonInclude.Include.NON_NULL)
public class ErrorDetailResponse {

    private String code;
    private String description;
    private List<String> fields;
    private Long id;

    public ErrorDetailResponse(String code, String description, List<String> fields) {
        this.code = code;
        this.description = description;
        this.fields = fields;
    }

    public ErrorDetailResponse(String code, String description, Long id) {
        this.code = code;
        this.description = description;
        this.id = id;
    }
}
