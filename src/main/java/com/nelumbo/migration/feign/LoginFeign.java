package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.responses.DefaultResponse;
import com.nelumbo.migration.feign.dto.requests.LoginRequest;
import com.nelumbo.migration.feign.dto.responses.LoginResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;

@FeignClient(name = "login", url="${hrcore.users-api}/login")
public interface LoginFeign {
    @PostMapping
    DefaultResponse<LoginResponse> login(@RequestBody LoginRequest loginRequest);
}