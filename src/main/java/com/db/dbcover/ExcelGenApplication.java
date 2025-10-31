package com.db.dbcover;

import com.db.dbcover.config.ExcelTemplateProperties;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.properties.EnableConfigurationProperties;

@SpringBootApplication
@EnableConfigurationProperties(ExcelTemplateProperties.class)
public class ExcelGenApplication {

    public static void main(String[] args) {
        SpringApplication.run(ExcelGenApplication.class, args);
    }
}
