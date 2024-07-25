package com.nelumbo.migration.feign;

import com.nelumbo.migration.exceptions.CustomErrorDecoder;
import com.nelumbo.migration.feign.dto.requests.GroupsProfRequest;
import com.nelumbo.migration.feign.dto.requests.GroupsRequest;
import com.nelumbo.migration.feign.dto.responses.DefaultNameResponse;
import com.nelumbo.migration.feign.dto.responses.DefaultResponse;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@FeignClient(name= "groups", url="${hr-api}/groups", configuration = CustomErrorDecoder.class)
public interface GroupsFeign {

    @GetMapping("/names")
    DefaultResponse<List<DefaultNameResponse>> findAllGroupNames(@RequestHeader("Authorization") String token);

    @PostMapping
    void createGroups(@RequestHeader("Authorization") String token, @RequestBody GroupsRequest groupsRequest);

    @PostMapping("group-assignments")
    void createGroupsAssigments(@RequestHeader("Authorization") String token, @RequestBody GroupsProfRequest gPRequest);
}
