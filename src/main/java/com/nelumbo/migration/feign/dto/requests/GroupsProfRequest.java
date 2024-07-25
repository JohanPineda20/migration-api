package com.nelumbo.migration.feign.dto.requests;

import lombok.Getter;
import lombok.Setter;

import java.time.LocalDate;
import java.util.Set;

@Getter
@Setter
public class GroupsProfRequest {

    private Set<Long> groupIds;

    private Set<Long> profileIds;

    private boolean temporal;

    private LocalDate dateFrom;

    private LocalDate dateTo;

    private boolean allProfiles;

    private FilterProfiles filterProfiles;

    private Set<Long> excludeProfileIds;
}
