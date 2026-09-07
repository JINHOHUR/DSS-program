package kr.co.dss.fx.domain;

import jakarta.persistence.*;
import lombok.Getter;
import lombok.Setter;

/** 고객사별 신용장 기한 규칙. 없으면 fx_setting 의 ship_days / expiry_days 기본값. */
@Entity
@Table(name = "fx_lc_rule")
@Getter @Setter
public class LcRule {
    @Id
    @Column(name = "company_code", length = 20)
    private String companyCode;

    @Column(name = "ship_days", nullable = false)
    private int shipDays;

    @Column(name = "expiry_days", nullable = false)
    private int expiryDays;
}
