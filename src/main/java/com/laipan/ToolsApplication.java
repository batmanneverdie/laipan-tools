package com.laipan;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

/**
 * <p><p/>
 *
 * @author laipan
 * @date 2022/02/28,11:10
 * @since v0.1
 */
@SpringBootApplication(scanBasePackages = {"com.laipan"})
public class ToolsApplication {
    public static void main(String[] args) {
        SpringApplication.run(ToolsApplication.class, args);
    }
}
