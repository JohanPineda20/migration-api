package com.nelumbo.migration.feign.dto.requests;

import lombok.Builder;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
@Builder
public class WorkPositionUpdateRequest {

    private Long compTabId;
    private Long compCategoryId;
    private Long orgManagerId;
    private Long approvalManagerId;
}
