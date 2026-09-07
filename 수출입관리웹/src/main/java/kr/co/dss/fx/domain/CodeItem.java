package kr.co.dss.fx.domain;

import jakarta.persistence.*;
import lombok.Getter;
import lombok.Setter;

/** 선택 목록 (공통). category 는 모듈 접두를 붙인다: fx.구분, com.통화 */
@Entity
@Table(name = "com_code")
@Getter @Setter
public class CodeItem {
    @Id @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(nullable = false, length = 50)
    private String category;

    /** 컬럼명은 code_value — 'value' 는 H2 예약어 */
    @Column(name = "code_value", nullable = false, length = 100)
    private String value;

    private Integer sort = 0;
}
