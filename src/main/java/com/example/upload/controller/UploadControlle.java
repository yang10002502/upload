package com.example.upload.controller;

import org.apache.catalina.util.RequestUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.HttpRequest;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import java.util.Map;

/**
 * @author yangxinlei on 2019/11/20 19:05
 */
@RestController
@RequestMapping("upload")
public class UploadControlle {

    private static Logger logger = LoggerFactory.getLogger(UploadControlle.class);

    @PostMapping("save")
    public String upload(@RequestParam("file") MultipartFile file,
                         @RequestParam(value = "desc", defaultValue = "hello") String desc) {

        String fileName = file.getOriginalFilename();
        logger.info("file name : {} desc :{}", fileName, desc);
        return "success";
    }
}
