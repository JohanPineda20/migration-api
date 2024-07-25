package com.nelumbo.migration.feign.dto.requests;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class GroupsRequest {

    private String name;
    private String description;
}
