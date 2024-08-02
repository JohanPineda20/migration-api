package com.nelumbo.migration.feign.dto.responses;

import lombok.Getter;
import lombok.Setter;

import java.util.List;
@Getter
@Setter
public class StructByTypeResponse {
    private List<OrgEntDetailResponse> details;
}
