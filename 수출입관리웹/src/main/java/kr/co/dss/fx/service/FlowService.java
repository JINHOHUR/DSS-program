package kr.co.dss.fx.service;

import kr.co.dss.fx.Stages;
import kr.co.dss.fx.domain.Deal;
import kr.co.dss.fx.domain.FlowEdge;
import kr.co.dss.fx.domain.FlowNode;
import kr.co.dss.fx.repo.Repos.FlowEdgeRepo;
import kr.co.dss.fx.repo.Repos.FlowNodeRepo;
import org.springframework.boot.ApplicationRunner;
import org.springframework.context.annotation.Bean;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.*;

/** 업무 플로우차트 — 자유 배치 노드·연결선 */
@Service
public class FlowService {
    private final FlowNodeRepo nodes;
    private final FlowEdgeRepo edges;
    private final DealService deals;

    public FlowService(FlowNodeRepo nodes, FlowEdgeRepo edges, DealService deals) {
        this.nodes = nodes; this.edges = edges; this.deals = deals;
    }

    @Bean
    ApplicationRunner seedFlow() { return args -> { if (nodes.count() == 0) reset(); }; }

    public List<FlowNode> nodes() { return nodes.findAllByOrderBySeqAscIdAsc(); }
    public List<FlowEdge> edges() { return edges.findAllByOrderByIdAsc(); }
    public Optional<FlowNode> node(Long id) { return nodes.findById(id); }

    /** 엑셀 기준 0~8단계 + 순서 연결선으로 되돌린다 */
    @Transactional
    public void reset() {
        edges.deleteAll();
        nodes.deleteAll();
        List<FlowNode> saved = new ArrayList<>();
        for (Stages.Stage s : Stages.ALL) {
            FlowNode n = new FlowNode();
            n.setSeq(s.no() * 10); n.setTitle(s.no() + ". " + s.name()); n.setDescr(s.desc());
            n.setTip(s.tip()); n.setStore(s.store()); n.setFolder(""); n.setStageNo(s.no());
            n.setX(60 + (s.no() % 3) * 300); n.setY(50 + (s.no() / 3) * 160); n.setShape("box"); n.setColor("");
            saved.add(nodes.save(n));
        }
        for (int i = 0; i + 1 < saved.size(); i++) {
            FlowEdge e = new FlowEdge(); e.setSrc(saved.get(i).getId()); e.setDst(saved.get(i + 1).getId()); e.setLabel("");
            edges.save(e);
        }
    }

    @Transactional
    public FlowNode save(FlowNode in) {
        FlowNode n = in.getId() == null ? new FlowNode() : nodes.findById(in.getId()).orElseGet(FlowNode::new);
        if (n.getId() == null) {
            int[] spot = freeSpot();
            n.setX(spot[0]); n.setY(spot[1]);
            n.setSeq(nodes().stream().mapToInt(x -> x.getSeq() == null ? 0 : x.getSeq()).max().orElse(0) + 10);
        }
        n.setTitle(nz(in.getTitle())); n.setDescr(nz(in.getDescr())); n.setTip(nz(in.getTip()));
        n.setStore(nz(in.getStore())); n.setFolder(nz(in.getFolder()));
        n.setStageNo(in.getStageNo() == null ? -1 : in.getStageNo());
        n.setShape(in.getShape() == null || in.getShape().isBlank() ? "box" : in.getShape());
        n.setColor(nz(in.getColor()));
        if (in.getX() != null && in.getId() != null) n.setX(in.getX());
        if (in.getY() != null && in.getId() != null) n.setY(in.getY());
        return nodes.save(n);
    }

    @Transactional
    public void setPos(Long id, int x, int y) {
        nodes.findById(id).ifPresent(n -> { n.setX(Math.max(0, x)); n.setY(Math.max(0, y)); nodes.save(n); });
    }

    @Transactional
    public void deleteNode(Long id) { edges.deleteBySrcOrDst(id, id); nodes.deleteById(id); }

    @Transactional
    public FlowEdge addEdge(Long src, Long dst) {
        if (src.equals(dst)) return null;
        Optional<FlowEdge> dup = edges.findBySrcAndDst(src, dst);
        if (dup.isPresent()) return dup.get();
        FlowEdge e = new FlowEdge(); e.setSrc(src); e.setDst(dst); e.setLabel("");
        return edges.save(e);
    }
    @Transactional public void deleteEdge(Long id) { edges.deleteById(id); }
    @Transactional
    public void setEdgeLabel(Long id, String label) {
        edges.findById(id).ifPresent(e -> { e.setLabel(nz(label)); edges.save(e); });
    }

    @Transactional
    public void autoLayout() {
        int i = 0;
        for (FlowNode n : nodes()) { n.setX(60 + (i % 3) * 300); n.setY(50 + (i / 3) * 160); nodes.save(n); i++; }
    }

    /** 단계별 '지금 그 단계에 있는' 진행중 안건 수 */
    public Map<Integer, Integer> stageCounts() {
        Map<Integer, Integer> out = new HashMap<>();
        for (Deal d : deals.all()) {
            if (!"진행중".equals(d.getStatus())) continue;
            Integer nxt = deals.progress(d.getId()).next();
            if (nxt != null) out.merge(nxt, 1, Integer::sum);
        }
        return out;
    }

    /** 그 단계에 머물러 있는 진행중 안건 */
    public List<Deal> dealsAtStage(int stageNo) {
        List<Deal> out = new ArrayList<>();
        for (Deal d : deals.all()) {
            if (!"진행중".equals(d.getStatus())) continue;
            Integer nxt = deals.progress(d.getId()).next();
            if (nxt != null && nxt == stageNo) out.add(d);
        }
        return out;
    }

    private int[] freeSpot() {
        Set<String> used = new HashSet<>();
        for (FlowNode n : nodes()) used.add(n.getX() + "," + n.getY());
        for (int r = 0; r < 40; r++) for (int c = 0; c < 4; c++) {
            int x = 60 + c * 300, y = 50 + r * 160;
            boolean clash = false;
            for (FlowNode n : nodes()) if (Math.abs(n.getX() - x) <= 40 && Math.abs(n.getY() - y) <= 40) { clash = true; break; }
            if (!clash) return new int[]{x, y};
        }
        return new int[]{60, 50};
    }
    private static String nz(String s) { return s == null ? "" : s; }
}
