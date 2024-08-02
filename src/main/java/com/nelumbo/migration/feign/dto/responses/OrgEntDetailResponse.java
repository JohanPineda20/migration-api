package com.nelumbo.migration.feign.dto.responses;

import lombok.Getter;
import lombok.Setter;

import java.util.List;
@Getter
@Setter
public class OrgEntDetailResponse {
    private Long id;
    private List<ParamOrgEntDetInstDetailResponse> structures;
}
