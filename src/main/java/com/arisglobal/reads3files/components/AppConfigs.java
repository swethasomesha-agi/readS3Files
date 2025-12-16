package com.arisglobal.reads3files.components;

import lombok.Getter;
import lombok.Setter;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Configuration;
import org.springframework.stereotype.Component;

@Getter
@Setter
@Component
@Configuration
public class AppConfigs {

    @Value("${appconfig.excel.file.path}")
    private String excelPath;
    @Value("${appconfig.s3.bucket}")
    private String s3BucketName;
    @Value("${appconfig.s3.region}")
    private String s3RegionName;
    @Value("${appconfig.local}")
    private boolean local;
    @Value("${appconfig.emlPath}")
    private String emlPath;
    @Value("${appconfig.tempPath}")
    private String tempPath;
    @Value("${appconfig.outputPath}")
    private String outputPath;
}
