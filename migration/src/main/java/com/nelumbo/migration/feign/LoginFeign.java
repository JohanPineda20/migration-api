package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.DefaultResponse;
import com.nelumbo.migration.feign.dto.requests.LoginRequest;
import com.nelumbo.migration.feign.dto.responses.LoginResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;

@FeignClient(name = "login", url="localhost:8080/core-api/v1/login")
public interface LoginFeign {
    @PostMapping
    DefaultResponse<LoginResponse> login(@RequestBody LoginRequest loginRequest);
}