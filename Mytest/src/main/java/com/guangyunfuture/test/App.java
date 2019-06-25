package com.guangyunfuture.test;

import lombok.extern.slf4j.Slf4j;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.web.servlet.config.annotation.EnableWebMvc;
import springfox.documentation.swagger2.annotations.EnableSwagger2;

/**
 * Hello world!
 *
 */
@SpringBootApplication
@Slf4j
@EnableSwagger2
public class App
{
    public static void main( String[] args )
    {
        SpringApplication.run(App.class,args);
        log.info("App start...");
    }
}
