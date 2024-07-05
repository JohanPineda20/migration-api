package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.requests.FileRequest;
import com.nelumbo.migration.feign.dto.requests.ProfileRequest;
import com.nelumbo.migration.feign.dto.responses.DefaultResponse;
import com.nelumbo.migration.feign.dto.responses.ProfileResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.*;

@FeignClient(name = "profiles", url="${hr-api}/profiles")
public interface ProfileFeign {
    @PostMapping
    DefaultResponse<ProfileResponse> createProfile(@RequestHeader("Authorization") String token,
                                                   @RequestBody ProfileRequest profileRequest);
    @PatchMapping(value = "/{profileId}/image")
    void updateImageUrlByProfileId(@RequestHeader("Authorization") String token,
                                   @PathVariable Long profileId,
                                   @RequestBody FileRequest request);
}
