package com.example.file.controller;

import com.example.file.service.ExcelService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletResponse;
import java.io.*;

/**
 * @author yangxinlei on 2019/11/20 19:05
 */
@RestController
@RequestMapping("file")
public class FileController {

    private static Logger logger = LoggerFactory.getLogger(FileController.class);
    @Resource
    private ExcelService excelService;

    @PostMapping("save")
    public String upload(@RequestParam("file") MultipartFile file,
                         @RequestParam(value = "desc", defaultValue = "hello") String desc) {

        String fileName = file.getOriginalFilename();
        logger.info("file name : {} desc :{}", fileName, desc);
        return "success";
    }

    @GetMapping("down")
    public String down(HttpServletResponse response) {

        String path = "/Users/yangxinlei/";
        String fileName = "dsgsn.mp3";
        File file = new File(path, fileName);
        response.setContentType("application/force-download");// 设置强制下载不打开
        response.addHeader("Content-Disposition", "attachment;fileName=" + fileName);
        byte[] buffer = new byte[1024];
        FileInputStream fis = null;
        BufferedInputStream bis = null;
        try {
            fis = new FileInputStream(file);
            bis = new BufferedInputStream(fis);
            OutputStream outputStream = response.getOutputStream();
            int i = bis.read(buffer);
            while (i != -1) {
                outputStream.write(buffer, 0, i);
                i = bis.read(buffer);
            }
            return "success";
        } catch (Exception e) {
            e.printStackTrace();

        } finally {
            if (bis != null) {
                try {
                    bis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return "fail";

    }

    @GetMapping("createExcel")
    public String createExcel(HttpServletResponse response) {
        try {
            excelService.createExcel(response);
            return "success";
        } catch (Exception e) {
            e.printStackTrace();
            return "fail";
        }
    }

}
