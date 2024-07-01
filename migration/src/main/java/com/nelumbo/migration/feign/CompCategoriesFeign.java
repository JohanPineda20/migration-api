package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.requests.CompCategoriesRequest;
import com.nelumbo.migration.feign.dto.responses.CompCategoriesResponse;
import com.nelumbo.migration.feign.dto.DefaultResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@FeignClient(name= "compCategory", url="localhost:8080/core-api/v1/compensation-categories")
public interface CompCategoriesFeign {

    @GetMapping("/simplified-search")
    DefaultResponse<List<CompCategoriesResponse>> simplifiedSearch(@RequestHeader("Authorization") String token, @RequestParam String search);

    @PostMapping
    void createCompensationCategories(@RequestHeader("Authorization") String token, @RequestBody CompCategoriesRequest compCategory);
}
