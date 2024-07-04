package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.requests.CompCategoriesRequest;
import com.nelumbo.migration.feign.dto.responses.CompCategoriesResponse;
import com.nelumbo.migration.feign.dto.responses.DefaultResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@FeignClient(name= "compCategory", url="${hr-api}/compensation-categories")
public interface CompCategoriesFeign {

    @GetMapping("/simplified-search")
    DefaultResponse<List<CompCategoriesResponse>> simplifiedSearch(@RequestHeader("Authorization") String token, @RequestParam String search);

    @PostMapping
    DefaultResponse<CompCategoriesResponse> createCompensationCategories(@RequestHeader("Authorization") String token, @RequestBody CompCategoriesRequest compCategory);
}
