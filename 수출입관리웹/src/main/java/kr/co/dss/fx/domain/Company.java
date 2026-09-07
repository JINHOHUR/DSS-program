package kr.co.dss.fx.domain;

import jakarta.persistence.*;
import lombok.Getter;
import lombok.Setter;

/** 거래처 마스터 (공통). 고객사 / 최종고객사 / 공급사. 이름 문자열 대신 code 로 참조한다. */
@Entity
@Table(name = "com_company")
@Getter @Setter
public class Company {
    @Id
    @Column(length = 20)
    private String code;

    @Column(nullable = false, length = 100)
    private String name;

    @Column(name = "short_name", length = 50)
    private String shortName;

    @Column(nullable = false, length = 20)
    private String kind;

    @Column(name = "biz_no", length = 20)
    private String bizNo;

    @Column(length = 1)
    private String active = "Y";

    private Integer sort = 0;

    public boolean isActive() { return !"N".equals(active); }
}
