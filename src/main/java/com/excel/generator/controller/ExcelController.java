package com.excel.generator.controller;

import lombok.Data;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.client.RestTemplate;

import java.net.URI;

/**
 * @author manish
 */
@RestController
@RequestMapping("services/shiprocket")
public class ExcelController {


    @PostMapping(value = "/getToken", consumes = "application/json", produces = "application/json")
    public ResponseEntity<ShipRocketTokenResponse> getToken(@RequestBody ShipRocketTokenRequest shipRocketTokenRequest) {
        RestTemplate restTemplate = new RestTemplate();
        return restTemplate.postForEntity(URI.create("https://apiv2.shiprocket.in/v1/external/auth/login"), shipRocketTokenRequest, ShipRocketTokenResponse.class);
    }

}

@Data
class ShipRocketTokenRequest {
    private String email;
    private String password;
}

@Data
class ShipRocketTokenResponse {
    private String token;

}
