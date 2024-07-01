package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.*;
import com.nelumbo.migration.feign.dto.responses.DefaultResponse;
import com.nelumbo.migration.feign.dto.responses.OrgEntityResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.*;

@FeignClient(name = "orgEntity", url="localhost:8080/core-api/v1/organization-entities")
public interface OrgEntityFeign {
    @GetMapping("/{orgEntityId}/get-instances/{orgEntDetParentId}")
    DefaultResponse<Page<OrgEntityResponse>> findAllInstancesParentOrganizationEntityDetail(@RequestHeader("Authorization") String token,
                                                                                            @PathVariable Long orgEntityId,
                                                                                            @PathVariable Long orgEntDetParentId);
}