package kr.co.dss.fx.web;

import kr.co.dss.fx.Stages;
import kr.co.dss.fx.domain.Deal;
import kr.co.dss.fx.domain.FlowEdge;
import kr.co.dss.fx.domain.FlowNode;
import kr.co.dss.fx.service.ArchiveService;
import kr.co.dss.fx.service.DealService;
import kr.co.dss.fx.service.FlowService;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;

import java.util.*;

@Controller
public class FlowController {
    private final FlowService flow;
    private final ArchiveService archive;
    private final DealService deals;

    public FlowController(FlowService flow, ArchiveService archive, DealService deals) {
        this.flow = flow; this.archive = archive; this.deals = deals;
    }

    @GetMapping("/flow")
    public String page(Model m) {
        m.addAttribute("stages", Stages.ALL);
        m.addAttribute("active_menu", "flow");
        return "flow";
    }

    @GetMapping("/api/flow")
    @ResponseBody
    public Map<String, Object> data() {
        return Map.of("nodes", flow.nodes(), "edges", flow.edges(), "counts", flow.stageCounts());
    }

    @PostMapping("/api/flow/node")
    @ResponseBody
    public FlowNode save(@RequestBody FlowNode n) { return flow.save(n); }

    @PostMapping("/api/flow/node/{id}/pos")
    @ResponseBody
    public Map<String, Object> pos(@PathVariable Long id, @RequestParam int x, @RequestParam int y) {
        flow.setPos(id, x, y); return Map.of("ok", true);
    }

    @DeleteMapping("/api/flow/node/{id}")
    @ResponseBody
    public Map<String, Object> del(@PathVariable Long id) { flow.deleteNode(id); return Map.of("ok", true); }

    @PostMapping("/api/flow/edge")
    @ResponseBody
    public Object edge(@RequestParam Long src, @RequestParam Long dst) {
        FlowEdge e = flow.addEdge(src, dst);
        return e == null ? Map.of("ok", false) : e;
    }

    @DeleteMapping("/api/flow/edge/{id}")
    @ResponseBody
    public Map<String, Object> delEdge(@PathVariable Long id) { flow.deleteEdge(id); return Map.of("ok", true); }

    @PostMapping("/api/flow/edge/{id}/label")
    @ResponseBody
    public Map<String, Object> label(@PathVariable Long id, @RequestParam(defaultValue = "") String label) {
        flow.setEdgeLabel(id, label); return Map.of("ok", true);
    }

    @PostMapping("/api/flow/reset")
    @ResponseBody
    public Map<String, Object> reset() { flow.reset(); return Map.of("ok", true); }

    @PostMapping("/api/flow/auto-layout")
    @ResponseBody
    public Map<String, Object> auto() { flow.autoLayout(); return Map.of("ok", true); }

    /** 항목 클릭 시 오른쪽 상세: 관련 자료(폴더 내용) + 그 단계 진행중 안건 */
    @GetMapping("/api/flow/node/{id}/detail")
    @ResponseBody
    public Map<String, Object> detail(@PathVariable Long id) {
        FlowNode n = flow.node(id).orElseThrow();
        String rel = "";
        if (n.getFolder() != null && !n.getFolder().isBlank()) rel = n.getFolder();
        else if (n.getStageNo() != null && n.getStageNo() >= 0 && n.getStageNo() < Stages.COUNT)
            rel = archive.keywordFolder(Stages.get(n.getStageNo()).keyword());
        ArchiveService.Listing ls = archive.rootExists() ? archive.list(rel) : null;
        List<Map<String, Object>> files = new ArrayList<>();
        if (ls != null && ls.ok()) for (ArchiveService.Entry e : ls.entries())
            files.add(Map.of("name", e.name(), "rel", e.rel(), "dir", e.dir(), "size", e.sizeText(), "mtime", e.mtime(), "kind", e.kind()));
        List<Map<String, Object>> ds = new ArrayList<>();
        if (n.getStageNo() != null && n.getStageNo() >= 0)
            for (Deal d : flow.dealsAtStage(n.getStageNo())) {
                DealService.Deadline dl = deals.nextDeadline(d);
                ds.add(Map.of("id", d.getId(), "docNo", d.getDocNo(), "customer", d.customerName(),
                        "model", d.getModel() == null ? "" : d.getModel(), "next", dl.text(), "tag", dl.tag()));
            }
        Map<String, Object> out = new HashMap<>();
        out.put("node", n); out.put("folder", rel); out.put("files", files); out.put("deals", ds);
        return out;
    }
}
