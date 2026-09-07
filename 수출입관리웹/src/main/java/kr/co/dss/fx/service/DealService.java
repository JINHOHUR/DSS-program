package kr.co.dss.fx.service;

import kr.co.dss.fx.Stages;
import kr.co.dss.fx.auth.UserContext;
import kr.co.dss.fx.domain.Deal;
import kr.co.dss.fx.domain.StageLog;
import kr.co.dss.fx.repo.Repos.DealRepo;
import kr.co.dss.fx.repo.Repos.StageLogRepo;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.stream.Collectors;

/** 안건 저장·채번·단계·기한 계산 (spec/rules.md §2~§5) */
@Service
public class DealService {
    private final DealRepo deals;
    private final StageLogRepo stages;
    private final MasterService master;
    private final SettingService settings;

    public DealService(DealRepo deals, StageLogRepo stages, MasterService master, SettingService settings) {
        this.deals = deals; this.stages = stages; this.master = master; this.settings = settings;
    }

    // ---- 조회 ----------------------------------------------------------------
    public List<Deal> all() { return deals.findAllActive(); }

    public Deal get(Long id) {
        return deals.findById(id).filter(d -> d.getDeletedAt() == null)
                .orElseThrow(() -> new NoSuchElementException("안건이 없습니다: " + id));
    }

    public List<Deal> filtered(String q, String customerCode, String status) {
        String qq = q == null ? "" : q.trim().toLowerCase();
        return all().stream().filter(d -> {
            if (customerCode != null && !customerCode.isBlank() && !customerCode.equals(d.customerCode())) return false;
            if (status != null && !status.isBlank() && !"전체".equals(status) && !status.equals(d.getStatus())) return false;
            if (!qq.isEmpty()) {
                String blob = String.join(" ", n(d.getDocNo()), n(d.customerName()), n(d.endUserName()), n(d.getModel()),
                        n(d.getPoNo()), n(d.getOfferNo()), n(d.getLcNo()), n(d.getSite()), n(d.getNote()),
                        n(d.getDebitNo()), n(d.getCommInvoiceNo()), n(d.getRefNo())).toLowerCase();
                return blob.contains(qq);
            }
            return true;
        }).collect(Collectors.toList());
    }
    private static String n(String s) { return s == null ? "" : s; }

    // ---- 채번: FX-YYYY-NNNN ------------------------------------------------------
    public String nextDocNo() {
        int y = LocalDate.now().getYear();
        String prefix = "FX-" + y + "-";
        int max = 0;
        for (String no : deals.docNosLike(prefix + "%")) {
            try { max = Math.max(max, Integer.parseInt(no.substring(prefix.length()))); } catch (Exception ignore) {}
        }
        return prefix + String.format("%04d", max + 1);
    }

    // ---- 저장 ----------------------------------------------------------------
    @Transactional
    public Deal create() {
        Deal d = new Deal();
        d.setDocNo(nextDocNo());
        d.setStatus("진행중");
        d.setCurrency("JPY");
        d.setKind("일반");
        d.setOwner(UserContext.id());
        stamp(d, true);
        return deals.save(d);
    }

    /** 폼에서 온 값을 기존 안건에 반영하고 저장. FCA 가 있고 기한이 비어 있으면 자동 계산. */
    @Transactional
    public Deal save(Deal d) {
        if (d.getDocNo() == null || d.getDocNo().isBlank()) d.setDocNo(nextDocNo());
        if (d.getOwner() == null || d.getOwner().isBlank()) d.setOwner(UserContext.id());
        autoLc(d, false);
        stamp(d, d.getId() == null);
        return deals.save(d);
    }

    /** 신용장 기한 계산. force=true 면 무조건 덮어씀, false 면 비어 있는 칸만 채움 */
    public void autoLc(Deal d, boolean force) {
        if (d.getFcaDate() == null) return;
        int[] r = master.lcRule(d.customerCode());
        if (force || d.getLatestShipment() == null) d.setLatestShipment(d.getFcaDate().plusDays(r[0]));
        if (force || d.getExpiryDate() == null) d.setExpiryDate(d.getFcaDate().plusDays(r[1]));
    }

    public LocalDate lcRequestDue(Deal d) {
        return d.getFcaDate() == null ? null : d.getFcaDate().minusDays(settings.lcLead());
    }

    private void stamp(Deal d, boolean created) {
        LocalDateTime now = LocalDateTime.now();
        String u = UserContext.id();
        if (created) { d.setCreatedAt(now); d.setCreatedBy(u); }
        d.setUpdatedAt(now); d.setUpdatedBy(u);
    }

    @Transactional
    public void softDelete(Long id) {
        Deal d = get(id);
        d.setDeletedAt(LocalDateTime.now());
        d.setUpdatedBy(UserContext.id());
        deals.save(d);
    }

    @Transactional
    public Deal duplicate(Long id) {
        Deal s = get(id);
        Deal d = new Deal();
        d.setDocNo(nextDocNo()); d.setStatus("진행중");
        d.setCustomer(s.getCustomer()); d.setCustomerPic(s.getCustomerPic()); d.setEndUser(s.getEndUser());
        d.setSite(s.getSite()); d.setModel(s.getModel()); d.setQtyNum(s.getQtyNum()); d.setQtyUnit(s.getQtyUnit());
        d.setKind(s.getKind()); d.setNeedDate(s.getNeedDate());
        d.setDebitNo(s.getDebitNo()); d.setOfferNo(s.getOfferNo()); d.setFcaDate(s.getFcaDate());
        d.setLatestShipment(s.getLatestShipment()); d.setExpiryDate(s.getExpiryDate());
        d.setLcBank(s.getLcBank()); d.setAmount(s.getAmount()); d.setCurrency(s.getCurrency());
        d.setCommRate(s.getCommRate()); d.setNote(s.getNote()); d.setArchiveDir(s.getArchiveDir());
        d.setOwner(UserContext.id());
        stamp(d, true);
        return deals.save(d);
    }

    // ---- 단계 ----------------------------------------------------------------
    /** stage_no → StageLog (없는 단계는 빈 로그) */
    public Map<Integer, StageLog> stageMap(Long dealId) {
        Map<Integer, StageLog> m = new TreeMap<>();
        for (int i = 0; i < Stages.COUNT; i++) {
            StageLog s = new StageLog(); s.setDealId(dealId); s.setStageNo(i);
            m.put(i, s);
        }
        for (StageLog s : stages.findByDealId(dealId)) m.put(s.getStageNo(), s);
        return m;
    }

    public Set<Integer> doneSet(Long dealId) {
        return stages.findByDealId(dealId).stream().filter(StageLog::isDone).map(StageLog::getStageNo).collect(Collectors.toSet());
    }

    @Transactional
    public void setStage(Long dealId, int stageNo, LocalDate doneDate, String memo) {
        StageLog s = stages.findByDealIdAndStageNo(dealId, stageNo).orElseGet(() -> {
            StageLog x = new StageLog(); x.setDealId(dealId); x.setStageNo(stageNo); return x;
        });
        s.setDoneDate(doneDate);
        s.setMemo(memo == null ? "" : memo.trim());
        stages.save(s);
        Deal d = get(dealId);
        stamp(d, false);
        deals.save(d);
    }

    /** [완료하고 다음] — 완료일이 비어 있으면 오늘 */
    @Transactional
    public int completeStage(Long dealId, int stageNo, String memo) {
        StageLog cur = stageMap(dealId).get(stageNo);
        setStage(dealId, stageNo, cur.getDoneDate() != null ? cur.getDoneDate() : LocalDate.now(),
                memo != null ? memo : cur.getMemo());
        return Math.min(stageNo + 1, Stages.COUNT - 1);
    }

    /** (완료 단계 수, 다음 단계 번호 or null=전부 완료) */
    public record Progress(int done, Integer next) {
        public String nextText() { return next == null ? "완료" : next + ". " + Stages.name(next); }
        public String ratio() { return done + "/" + Stages.COUNT; }
    }

    public Progress progress(Long dealId) {
        Set<Integer> done = doneSet(dealId);
        Integer next = null;
        for (int i = 0; i < Stages.COUNT; i++) if (!done.contains(i)) { next = i; break; }
        return new Progress(done.size(), next);
    }

    /** 첫 미완료 단계 (없으면 마지막) — 화면 진입 시 열 단계 */
    public int currentStage(Long dealId) {
        Integer n = progress(dealId).next();
        return n == null ? Stages.COUNT - 1 : n;
    }

    // ---- 다음 기한 (rules.md §5) --------------------------------------------------
    public record Deadline(String label, LocalDate date, Integer days) {
        public String text() { return label == null ? "" : label + " " + Dates.ddayText(days); }
        public String tag() {
            if (days == null) return "";
            return days < 0 ? "late" : (days <= 7 ? "soon" : "");
        }
    }

    public Deadline nextDeadline(Deal d) {
        return nextDeadline(d, doneSet(d.getId()));
    }

    public Deadline nextDeadline(Deal d, Set<Integer> done) {
        List<Deadline> c = new ArrayList<>();
        int lead = settings.lcLead(), due = settings.commDue();
        if (!done.contains(4) && d.getFcaDate() != null) c.add(dl("L/C 요청", d.getFcaDate().minusDays(lead)));
        if (!done.contains(5) && d.getFcaDate() != null) c.add(dl("FCA", d.getFcaDate()));
        if (!done.contains(5) && d.getLatestShipment() != null) c.add(dl("선적기한", d.getLatestShipment()));
        if (!done.contains(6) && d.getExpiryDate() != null) c.add(dl("L/C 만료", d.getExpiryDate()));
        if (!done.contains(8) && d.getCommBilled() != null && d.getCommReceived() == null)
            c.add(dl("커미션 입금", d.getCommBilled().plusDays(due)));
        if (c.isEmpty()) return new Deadline(null, null, null);
        c.sort(Comparator.comparing(Deadline::date));
        for (Deadline x : c) if (x.days() >= 0) return x;
        return c.get(c.size() - 1);
    }
    private static Deadline dl(String label, LocalDate d) { return new Deadline(label, d, Dates.dday(d)); }
}
