package com.nelumbo.migration.feign;

import com.nelumbo.migration.exceptions.CustomErrorDecoder;
import com.nelumbo.migration.feign.dto.requests.FileRequest;
import com.nelumbo.migration.feign.dto.requests.ProfileRequest;
import com.nelumbo.migration.feign.dto.requests.ProfileSecValueRequest;
import com.nelumbo.migration.feign.dto.responses.DefaultResponse;
import com.nelumbo.migration.feign.dto.responses.ProfileResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.*;

@FeignClient(name = "profiles", url="${hr-api}/profiles", configuration = CustomErrorDecoder.class)
public interface ProfileFeign {
    @PostMapping
    DefaultResponse<ProfileResponse> createProfile(@RequestHeader("Authorization") String token,
                                                   @RequestBody ProfileRequest profileRequest);
    @PatchMapping(value = "/{profileId}/image")
    void updateImageUrlByProfileId(@RequestHeader("Authorization") String token,
                                   @PathVariable Long profileId,
                                   @RequestBody FileRequest request);
    @PostMapping(value = "/{profileId}/profile-section-values")
    void createProfileSectionValueByProfile(@RequestHeader("Authorization") String token,
                                            @PathVariable Long profileId,
                                            @RequestBody ProfileSecValueRequest profileSecValueRequest);
}
