package com.nelumbo.migration.feign.dto.responses;

import lombok.Getter;
import lombok.Setter;

import java.util.List;
@Getter
@Setter
public class ParamOrgEntDetInstDetailResponse {
    private Long id;
    private String name;
    private List<ParamOrgEntDetInstDetailResponse> children;
}
