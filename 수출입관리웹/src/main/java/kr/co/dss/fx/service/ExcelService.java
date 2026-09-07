package kr.co.dss.fx.service;

import kr.co.dss.fx.Stages;
import kr.co.dss.fx.domain.Deal;
import kr.co.dss.fx.domain.StageLog;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.temporal.ChronoUnit;
import java.util.*;

/** 엑셀 5시트 내보내기 (안건목록 / 진행현황 / 신용장현황 / 커미션현황 / 기한알림) */
@Service
public class ExcelService {
    private final DealService deals;
    private final MasterService master;
    private final SettingService settings;
    private final AlertService alerts;

    public ExcelService(DealService deals, MasterService master, SettingService settings, AlertService alerts) {
        this.deals = deals; this.master = master; this.settings = settings; this.alerts = alerts;
    }

    public byte[] build() throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            CellStyle head = wb.createCellStyle();
            Font hf = wb.createFont(); hf.setBold(true); hf.setColor(IndexedColors.WHITE.getIndex());
            head.setFont(hf);
            head.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
            head.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            head.setAlignment(HorizontalAlignment.CENTER);
            CellStyle late = fill(wb, IndexedColors.ROSE), soon = fill(wb, IndexedColors.LEMON_CHIFFON), ok = fill(wb, IndexedColors.LIGHT_GREEN);

            List<Deal> all = deals.all();

            // 1) 안건목록
            Sheet s1 = wb.createSheet("안건목록");
            row(s1, 0, head, "관리번호", "상태", "고객사", "담당자", "최종고객사", "사이트", "모델/품명", "수량", "구분",
                    "P.O 번호", "P.O 접수일", "Debit Note", "Offer Sheet", "FCA", "Latest shipment", "Expiry date",
                    "L/C 번호", "개설은행", "금액", "통화", "커미션 요율", "커미션 금액", "청구일", "수령일",
                    "다음 기한", "진행단계", "진척", "담당", "비고");
            int r = 1;
            for (Deal d : all) {
                DealService.Progress p = deals.progress(d.getId());
                DealService.Deadline dl = deals.nextDeadline(d);
                row(s1, r++, null, d.getDocNo(), d.getStatus(), d.customerName(), nz(d.getCustomerPic()), d.endUserName(),
                        nz(d.getSite()), nz(d.getModel()), d.qtyText(), nz(d.getKind()), nz(d.getPoNo()), iso(d.getPoDate()),
                        nz(d.getDebitNo()), nz(d.getOfferNo()), iso(d.getFcaDate()), iso(d.getLatestShipment()), iso(d.getExpiryDate()),
                        nz(d.getLcNo()), nz(d.getLcBank()), Dates.money(d.getAmount()), nz(d.getCurrency()),
                        d.getCommRate() == null ? "" : d.getCommRate().toPlainString(), Dates.money(d.getCommAmount()),
                        iso(d.getCommBilled()), iso(d.getCommReceived()),
                        dl.label() == null ? "" : dl.text() + " (" + dl.date() + ")", p.nextText(), p.ratio(), nz(d.getOwner()), nz(d.getNote()));
            }
            widths(s1, 12, 8, 16, 10, 12, 14, 24, 10, 8, 16, 12, 14, 14, 12, 14, 14, 18, 12, 14, 7, 10, 14, 12, 12, 26, 22, 8, 10, 30);

            // 2) 진행현황
            Sheet s2 = wb.createSheet("진행현황");
            List<String> h2 = new ArrayList<>(List.of("관리번호", "고객사", "P.O 번호"));
            for (int i = 0; i < Stages.COUNT; i++) h2.add(i + ". " + Stages.name(i));
            h2.add("진척");
            row(s2, 0, head, h2.toArray(new String[0]));
            r = 1;
            for (Deal d : all) {
                Map<Integer, StageLog> st = deals.stageMap(d.getId());
                List<String> v = new ArrayList<>(List.of(d.getDocNo(), d.customerName(), nz(d.getPoNo())));
                for (int i = 0; i < Stages.COUNT; i++) v.add(iso(st.get(i).getDoneDate()));
                v.add(deals.progress(d.getId()).ratio());
                Row rr = row(s2, r++, null, v.toArray(new String[0]));
                for (int i = 0; i < Stages.COUNT; i++)
                    if (!v.get(3 + i).isBlank()) rr.getCell(3 + i).setCellStyle(ok);
            }
            int[] w2 = new int[3 + Stages.COUNT + 1]; Arrays.fill(w2, 14); w2[0] = 14; w2[1] = 16; w2[2] = 16; w2[w2.length - 1] = 8;
            widths(s2, w2);

            // 3) 신용장현황
            Sheet s3 = wb.createSheet("신용장현황");
            row(s3, 0, head, "관리번호", "고객사", "P.O 번호", "FCA", "개설요청 기한", "개설은행", "L/C 번호", "개설일",
                    "Latest shipment", "Expiry date", "적용규칙", "Amend");
            r = 1;
            for (Deal d : all) {
                int[] rule = master.lcRule(d.customerCode());
                row(s3, r++, null, d.getDocNo(), d.customerName(), nz(d.getPoNo()), iso(d.getFcaDate()), iso(deals.lcRequestDue(d)),
                        nz(d.getLcBank()), nz(d.getLcNo()), iso(d.getLcOpenDate()), iso(d.getLatestShipment()), iso(d.getExpiryDate()),
                        "+" + rule[0] + "일 / +" + rule[1] + "일", nz(d.getLcAmend()));
            }
            widths(s3, 14, 16, 16, 12, 14, 12, 20, 12, 16, 14, 16, 24);

            // 4) 커미션현황
            Sheet s4 = wb.createSheet("커미션현황");
            row(s4, 0, head, "관리번호", "고객사", "P.O 번호", "납품금액", "통화", "요율(%)", "커미션 금액", "청구서 번호", "청구일", "수령일", "상태");
            r = 1;
            int due = settings.commDue();
            for (Deal d : all) {
                String state;
                if (d.getCommReceived() != null) state = "수령 완료";
                else if (d.getCommBilled() != null) {
                    long n = ChronoUnit.DAYS.between(d.getCommBilled(), LocalDate.now());
                    state = (n >= due ? "입금 지연 (" : "입금 대기 (") + n + "일)";
                } else if (deals.doneSet(d.getId()).contains(7)) state = "청구 필요";
                else state = "";
                row(s4, r++, null, d.getDocNo(), d.customerName(), nz(d.getPoNo()), Dates.money(d.getAmount()), nz(d.getCurrency()),
                        d.getCommRate() == null ? "" : d.getCommRate().toPlainString(), Dates.money(d.getCommAmount()),
                        nz(d.getCommInvoiceNo()), iso(d.getCommBilled()), iso(d.getCommReceived()), state);
            }
            widths(s4, 14, 16, 16, 14, 7, 8, 14, 30, 12, 12, 16);

            // 5) 기한알림
            Sheet s5 = wb.createSheet("기한알림");
            row(s5, 0, head, "구분", "D-Day", "기준일", "항목", "관리번호", "고객사", "P.O", "내용");
            r = 1;
            for (AlertService.Alert a : alerts.compute()) {
                Row rr = row(s5, r++, null, a.level(), a.ddayText(), a.dateText(), a.item(), a.docNo(), a.customer(), a.po(), a.detail());
                CellStyle cs = "지연".equals(a.level()) ? late : ("임박".equals(a.level()) ? soon : null);
                if (cs != null) for (Cell c : rr) c.setCellStyle(cs);
            }
            widths(s5, 8, 9, 12, 20, 14, 16, 16, 60);

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            wb.write(out);
            return out.toByteArray();
        }
    }

    private static Row row(Sheet s, int idx, CellStyle style, String... vals) {
        Row r = s.createRow(idx);
        for (int i = 0; i < vals.length; i++) {
            Cell c = r.createCell(i);
            c.setCellValue(vals[i] == null ? "" : vals[i]);
            if (style != null) c.setCellStyle(style);
        }
        return r;
    }
    private static void widths(Sheet s, int... w) { for (int i = 0; i < w.length; i++) s.setColumnWidth(i, w[i] * 256); }
    private static CellStyle fill(Workbook wb, IndexedColors c) {
        CellStyle cs = wb.createCellStyle();
        cs.setFillForegroundColor(c.getIndex()); cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return cs;
    }
    private static String iso(LocalDate d) { return Dates.iso(d); }
    private static String nz(String s) { return s == null ? "" : s; }
}
