package com.herambco.hrcapi.controller;

import com.herambco.hrcapi.Dto.XlsFilterRequestDto;
import com.herambco.hrcapi.service.interfaces.FilterXLSInterface;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;
import java.util.Set;

@RequestMapping("/xlsGenerator")
@CrossOrigin
@RestController
public class userController {

    @Autowired
    private FilterXLSInterface filterXLSInterface;

    @RequestMapping("/hello")
    public ResponseEntity<String> startAppCheck() {
        return ResponseEntity.ok("Application Running Successfully !!!");
    }
    @PostMapping("/filterXls")
    public ResponseEntity<Object> filterXls(
            @RequestPart("files") List<MultipartFile> files,
            @RequestPart(name="itemFilter") String itemFilter,
            @RequestParam(name="gstFilter") Boolean gstFilter) {
        try{
            List<Object> response = filterXLSInterface.generateXLSBasedOnFilter(files.get(0), itemFilter, gstFilter);
            return  ResponseEntity.ok(response);
        } catch (Exception e) {
            throw e;
        }
    }

    @PostMapping("/getItemList")
    public ResponseEntity<Set<String>> getItemList(
            @RequestPart("files") List<MultipartFile> files) {
        try{
            Set<String> response = filterXLSInterface.getItemList(files.get(0));
            return  ResponseEntity.ok(response);
        } catch (Exception e) {
            throw e;
        }
    }

}
