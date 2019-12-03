package com.example.file.config;

import org.springframework.boot.web.servlet.MultipartConfigFactory;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.util.unit.DataSize;

import javax.servlet.MultipartConfigElement;

/**
 * @author wupeng@meiweishenghuo.com
 * @date 2018/7/9 下午3:46
 **/
@Configuration
public class MultipartConfigElementConfig {
    @Bean
    public MultipartConfigElement multipartConfigElement() {
        MultipartConfigFactory factory = new MultipartConfigFactory();
        //  单个数据大小
        factory.setMaxFileSize(DataSize.ofMegabytes(30)); // KB,MB
        /// 总上传数据大小
        factory.setMaxRequestSize(DataSize.ofMegabytes(80));
        return factory.createMultipartConfig();
    }
}
