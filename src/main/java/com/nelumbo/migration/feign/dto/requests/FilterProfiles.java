package com.nelumbo.migration.feign.dto.requests;

import lombok.Getter;
import lombok.Setter;

import java.util.Set;

@Getter
@Setter
public class FilterProfiles {
    private String search;
    private Set<Long> groupIds;
}
