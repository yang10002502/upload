package com.example.upload.test;

import org.apache.http.HttpEntity;
import org.apache.http.client.HttpClient;
import org.apache.http.client.config.RequestConfig;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.mime.MultipartEntityBuilder;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.DefaultHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;

import java.io.File;
import java.io.IOException;

/**
 * @author yangxinlei on 2019/11/20 19:51
 */
public class UploadTest {

    private static String url = "http://127.0.0.1:8080/upload/save";
    private static String path = "/Users/yangxinlei/dsgsn.mp3";
//    private static String path = "/Users/yangxinlei/test.txt";

    private static final int CONNECTION_REQUEST_TIMEOUT = 500000;

    private static final int CONNECT_TIMEOUT = 500000;


    public static void saveFile() throws IOException {

        RequestConfig requestConfig = RequestConfig.custom()
                .setConnectTimeout(CONNECT_TIMEOUT)
                .setConnectionRequestTimeout(CONNECTION_REQUEST_TIMEOUT)
                .build();

        CloseableHttpClient httpClient = HttpClients.custom().build();


        HttpPost httpPost = new HttpPost(url);
        httpPost.setConfig(requestConfig);

        MultipartEntityBuilder builder = MultipartEntityBuilder.create();
        File file = new File(path);
        System.out.print("......" + file.getName());
        builder.addBinaryBody("file", file);
        builder.addTextBody("desc", "yangxinlei");
        HttpEntity httpEntity = builder.build();
        httpPost.setEntity(httpEntity);
        CloseableHttpResponse httpResponse = httpClient.execute(httpPost);

        HttpEntity responseEntity = httpResponse.getEntity();
        String response = EntityUtils.toString(responseEntity);
        System.out.print(response);

    }

    public static void main(String[] args) throws IOException {
        UploadTest.saveFile();
    }
}
