package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.requests.TabsRequest;
import com.nelumbo.migration.feign.dto.responses.DefaultResponse;
import com.nelumbo.migration.feign.dto.responses.TabsResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@FeignClient(name= "tabs", url="${hr-api}/compensation-tabs")
public interface TabsFeign {

    @GetMapping("/simplified-search")
    DefaultResponse<List<TabsResponse>> simplifiedSearch(@RequestHeader("Authorization") String token, @RequestParam String search);

    @PostMapping
    DefaultResponse<TabsResponse> createTab(@RequestHeader("Authorization") String token, @RequestBody TabsRequest tabsRequest);
}
