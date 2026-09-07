package kr.co.dss.fx.domain;

import jakarta.persistence.*;
import lombok.Getter;
import lombok.Setter;

@Entity
@Table(name = "fx_flow_edge")
@Getter @Setter
public class FlowEdge {
    @Id @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;
    @Column(nullable = false) private Long src;
    @Column(nullable = false) private Long dst;
    @Column(length = 50) private String label = "";
}
