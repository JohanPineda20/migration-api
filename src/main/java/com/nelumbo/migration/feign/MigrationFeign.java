package com.nelumbo.migration.feign;

import com.nelumbo.migration.exceptions.CustomErrorDecoder;
import com.nelumbo.migration.feign.dto.Page;
import com.nelumbo.migration.feign.dto.requests.*;
import com.nelumbo.migration.feign.dto.responses.*;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@FeignClient(name= "migration", url="${hrcore.organization-api}/migration", configuration = CustomErrorDecoder.class)
public interface MigrationFeign {

    @GetMapping(path = "/field/{id}")
    DefaultResponse<DefaultNameResponse> getNameField(@RequestHeader("Authorization") String token,
                                                                   @PathVariable Long id,
                                                                   @RequestParam Integer fieldType);

    @GetMapping(path = "/organization-entities/{orgEntityId}/organization-entity-details")
    DefaultResponse<OrgEntityDetailResponse> findOrgEntityDetailByName(@RequestHeader("Authorization") String token,
                                                                       @PathVariable Long orgEntityId,
                                                                       @RequestParam String name);
    @PostMapping(value = "/organization-entities/{orgEntityId}/organization-entity-details")
    void createOrgEntityDetail(@RequestHeader("Authorization") String token,
                               @RequestBody OrgEntityDetailRequest orgEntityDetailRequest,
                               @PathVariable Long orgEntityId);
    @GetMapping("/country")
    DefaultResponse<List<CountryResponse>> findAll(@RequestHeader("Authorization") String token);
    @GetMapping("/country/{countryId}/states")
    DefaultResponse<List<CountryResponse>> findAllStatesByCountryId(@RequestHeader("Authorization") String token,
                                                                    @PathVariable Long countryId);
    @GetMapping("/country/{countryId}/states/{stateId}/cities")
    DefaultResponse<List<CountryResponse>> findAllCitesByStateIdAndCountryId(@RequestHeader("Authorization") String token,
                                                                             @PathVariable Long countryId,
                                                                             @PathVariable Long stateId);
    @GetMapping("/org-entities/{orgEntityId}/get-instances/{orgEntDetParentId}")
    DefaultResponse<Page<OrgEntityResponse>> findAllInstancesParentOrganizationEntityDetail(@RequestHeader("Authorization") String token,
                                                                                            @PathVariable Long orgEntityId,
                                                                                            @PathVariable Long orgEntDetParentId);



    @GetMapping("/compensation-category")
    DefaultResponse<CompCategoriesResponse> findCompCategoryByCode(@RequestHeader("Authorization") String token,
                                                                   @RequestParam String code);
    @PostMapping("/compensation-category")
    void createCompensationCategories(@RequestHeader("Authorization") String token,
                                      @RequestBody CompCategoriesRequest compCategory);



    @GetMapping("/cost-center")
    DefaultResponse<CostCenterResponse> findCostCenterByCode(@RequestHeader("Authorization") String token,
                                                             @RequestParam String code);
    @PostMapping("/cost-center")
    void createCostCenter(@RequestHeader("Authorization") String token,
                          @RequestBody CostCenterRequest costCenterRequest);
    @PostMapping(path = "/cost-center/{costCenterId}/details")
    void createCostCenterDetails(@RequestHeader("Authorization") String token,
                                 @RequestBody CostCenterDetailRequest costCenterDetailRequest,
                                 @PathVariable Long costCenterId);



    @GetMapping("/group")
    DefaultResponse<DefaultNameResponse> findGroupByName(@RequestHeader("Authorization") String token,
                                                         @RequestParam String name);
    @PostMapping("/group")
    void createGroups(@RequestHeader("Authorization") String token,
                      @RequestBody GroupsRequest groupsRequest);
    @PostMapping("/group-assignment")
    void createGroupsAssigments(@RequestHeader("Authorization") String token,
                                @RequestBody GroupsProfRequest gPRequest);



    @GetMapping("/profile")
    DefaultResponse<ProfileResponse> findProfileByClaveMpro(@RequestHeader("Authorization") String token,
                                                            @RequestParam String clave);
    @PostMapping("/profile")
    void createProfile(@RequestHeader("Authorization") String token,
                       @RequestBody ProfileRequest profileRequest);
    @PostMapping("/profile/{profileId}/profile-section-values")
    void createProfileSectionValueByProfile(@RequestHeader("Authorization") String token,
                                            @PathVariable Long profileId,
                                            @RequestBody ProfileSecValueRequest profileSecValueRequest);

    @PostMapping("/profile/{profileId}/profile-activation")
    void profileDraftActivation(@RequestHeader("Authorization") String token,
                                @PathVariable Long profileId);



    @GetMapping("/store")
    DefaultResponse<StoreResponse> findStoreByCode(@RequestHeader("Authorization") String token,
                                                   @RequestParam String code);
    @PostMapping("/store")
    DefaultResponse<StoreResponse> createStore(@RequestHeader("Authorization") String token,
                                               @RequestBody StoreRequest storeRequest);
    @PostMapping("/store/{storeId}/details")
    DefaultResponse<StoreDetailResponse> createStoreDetails(@RequestHeader("Authorization") String token,
                                                            @RequestBody StoreDetailRequest storeDetailRequest,
                                                            @PathVariable Long storeId);
    @PostMapping("/store/{storeId}/work-periods")
    void createStoreWorkPeriods(@RequestHeader("Authorization") String token,
                                @RequestBody StoreWorkPeriodRequest storeWorkPeriodRequest,
                                @PathVariable Long storeId);
    @GetMapping("/store/{storeId}/details")
    DefaultResponse<StoreDetailResponse> findAllStoresDetails(@RequestHeader("Authorization") String token,
                                                              @PathVariable Long storeId);



    @GetMapping("/compensation-tab")
    DefaultResponse<TabsResponse> findCompTabByCode(@RequestHeader("Authorization") String token,
                                                    @RequestParam String code);
    @PostMapping("/compensation-tab")
    void createTab(@RequestHeader("Authorization") String token,
                   @RequestBody TabsRequest tabsRequest);



    @GetMapping("/work-position-category")
    DefaultResponse<WorkPositionCategoryResponse> findWorkPosCategoryByCode(@RequestHeader("Authorization") String token,
                                                                            @RequestParam String code);
    @PostMapping("/work-position-category")
    void createWorkPositionCategory(@RequestHeader("Authorization") String token,
                                                                             @RequestBody WorkPositionCategoryRequest workPositionCategoryRequest);



    @GetMapping("/work-position")
    DefaultResponse<WorkPositionDetailResponse> findWorkPositionByCode(@RequestHeader("Authorization") String token,
                                                                       @RequestParam String code);
    @PostMapping("/work-position")
    void createWorkPosition(@RequestHeader("Authorization") String token,
                            @RequestBody WorkPositionRequest workPositionRequest);
    @PutMapping("/work-position/{workPositionId}")
    void updateWorkPosition(@RequestHeader("Authorization") String token,
                            @RequestBody WorkPositionUpdateRequest workPositionRequest,
                            @PathVariable Long workPositionId);
}
