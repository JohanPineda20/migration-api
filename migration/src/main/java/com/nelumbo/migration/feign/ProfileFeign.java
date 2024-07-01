package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.requests.ProfileRequest;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestHeader;

@FeignClient(name = "profiles", url="localhost:8080/core-api/v1/profiles")
public interface ProfileFeign {
    @PostMapping
    void createProfile(@RequestHeader("Authorization") String token,
                       @RequestBody ProfileRequest profileRequest);
}
