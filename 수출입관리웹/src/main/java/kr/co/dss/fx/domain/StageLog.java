package kr.co.dss.fx.domain;

import jakarta.persistence.*;
import lombok.Getter;
import lombok.Setter;

import java.time.LocalDate;

/** 안건의 단계 완료 기록. done_date 가 비어 있으면 미완료. */
@Entity
@Table(name = "fx_stage_log")
@Getter @Setter
public class StageLog {
    @Id @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(name = "deal_id", nullable = false)
    private Long dealId;

    @Column(name = "stage_no", nullable = false)
    private int stageNo;

    @Column(name = "done_date")
    private LocalDate doneDate;

    @Column(length = 300)
    private String memo;

    public boolean isDone() { return doneDate != null; }
}
