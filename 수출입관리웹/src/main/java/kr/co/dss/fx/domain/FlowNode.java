package kr.co.dss.fx.domain;

import jakarta.persistence.*;
import lombok.Getter;
import lombok.Setter;

/** 업무 플로우차트 항목. 자유 배치(x,y)·도형·색·연결 단계. */
@Entity
@Table(name = "fx_flow_node")
@Getter @Setter
public class FlowNode {
    @Id @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;
    private Integer seq = 0;
    @Column(length = 100) private String title;
    @Column(columnDefinition = "TEXT") private String descr;
    @Column(length = 500) private String tip;
    @Column(length = 100) private String store;
    @Column(length = 300) private String folder;
    @Column(name = "stage_no") private Integer stageNo = -1;
    private Integer x = 60;
    private Integer y = 50;
    @Column(length = 10) private String shape = "box";
    @Column(length = 10) private String color = "";
}
