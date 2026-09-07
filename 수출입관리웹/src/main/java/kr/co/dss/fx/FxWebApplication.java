package kr.co.dss.fx;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.data.jpa.repository.config.EnableJpaRepositories;

/**
 * DSS 외자(수출입) 관리 웹 모듈.
 * 통합 포털(자바) 안에 모듈로 들어가는 것을 전제로, 테이블은 fx_/com_ 접두, 사용자는 포털이 넘겨주는 값을 받는다.
 */
@SpringBootApplication
@EnableJpaRepositories(basePackages = "kr.co.dss.fx.repo", considerNestedRepositories = true)
public class FxWebApplication {
    public static void main(String[] args) {
        SpringApplication.run(FxWebApplication.class, args);
    }
}
