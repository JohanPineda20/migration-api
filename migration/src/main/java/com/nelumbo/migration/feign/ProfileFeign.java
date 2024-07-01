package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.ProfileRequest;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestHeader;
import org.springframework.web.bind.annotation.ResponseStatus;

@FeignClient(name = "profiles", url="localhost:8080/core-api/v1/profiles")
public interface ProfileFeign {
    @PostMapping
    void createProfile(@RequestHeader("Authorization") String token,
                       @RequestBody ProfileRequest profileRequest);
}
