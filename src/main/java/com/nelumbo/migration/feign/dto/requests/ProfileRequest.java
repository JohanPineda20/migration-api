package com.nelumbo.migration.feign.dto.requests;

import lombok.Getter;
import lombok.Setter;

import java.util.List;
@Getter
@Setter
public class ProfileRequest {
    private List<ProfileSecValueRequest> sectionValues;
    private Long workPositionId;
}
