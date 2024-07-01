package com.nelumbo.migration.feign;

import com.nelumbo.migration.feign.dto.requests.WorkPositionCategoryRequest;
import com.nelumbo.migration.feign.dto.responses.DefaultResponse;
import com.nelumbo.migration.feign.dto.responses.WorkPositionCategoryResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@FeignClient(name= "workPositionsCategories", url="localhost:8080/core-api/v1/work-position-categories")
public interface WorksPositionCategoriesFeign {

    @GetMapping("/simplified-search")
    DefaultResponse<List<WorkPositionCategoryResponse>> simplifiedSearch(@RequestHeader("Authorization") String token, @RequestParam String search);

    @PostMapping
    void createWorkPositionCategory(@RequestHeader("Authorization") String token, @RequestBody WorkPositionCategoryRequest workPositionCategoryRequest);
}
