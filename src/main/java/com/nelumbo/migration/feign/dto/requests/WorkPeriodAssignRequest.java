package com.nelumbo.migration.feign.dto.requests;

import lombok.Getter;
import lombok.Setter;

import java.util.Set;

@Getter
@Setter
public class WorkPeriodAssignRequest {
    private boolean temporal = false;
    private Set<Long> profileIds;
    private boolean force = true;
    private boolean allProfiles = false;
}
