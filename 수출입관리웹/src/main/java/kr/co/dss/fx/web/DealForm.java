package kr.co.dss.fx.web;

import kr.co.dss.fx.domain.Deal;
import kr.co.dss.fx.service.Dates;
import kr.co.dss.fx.service.MasterService;
import lombok.Getter;
import lombok.Setter;
import org.springframework.format.annotation.DateTimeFormat;

import java.math.BigDecimal;
import java.time.LocalDate;

/** 안건 입력 폼 (화면 ↔ 엔티티 변환). 금액은 콤마 포함 문자열도 받아 숫자로 저장한다. */
@Getter @Setter
public class DealForm {
    private String status, refNo, customerCode, customerPic, endUserCode, site, model, qtyNum, qtyUnit, kind;
    @DateTimeFormat(iso = DateTimeFormat.ISO.DATE) private LocalDate needDate;
    private String poNo, debitNo, offerNo;
    @DateTimeFormat(iso = DateTimeFormat.ISO.DATE) private LocalDate poDate, fcaDate, latestShipment, expiryDate;
    private String lcNo, lcBank, lcAmend;
    @DateTimeFormat(iso = DateTimeFormat.ISO.DATE) private LocalDate lcRequestDate, lcOpenDate;
    private String amount, currency, commRate, commAmount, commInvoiceNo;
    @DateTimeFormat(iso = DateTimeFormat.ISO.DATE) private LocalDate commBilled, commReceived;
    private String note, archiveDir, owner;
    private boolean recalcLc;

    public static DealForm of(Deal d) {
        DealForm f = new DealForm();
        f.status = d.getStatus(); f.refNo = d.getRefNo(); f.customerCode = d.customerCode(); f.customerPic = d.getCustomerPic();
        f.endUserCode = d.endUserCode(); f.site = d.getSite(); f.model = d.getModel();
        f.qtyNum = d.getQtyNum() == null ? "" : d.getQtyNum().stripTrailingZeros().toPlainString();
        f.qtyUnit = d.getQtyUnit(); f.kind = d.getKind(); f.needDate = d.getNeedDate();
        f.poNo = d.getPoNo(); f.poDate = d.getPoDate(); f.debitNo = d.getDebitNo(); f.offerNo = d.getOfferNo();
        f.fcaDate = d.getFcaDate(); f.latestShipment = d.getLatestShipment(); f.expiryDate = d.getExpiryDate();
        f.lcNo = d.getLcNo(); f.lcBank = d.getLcBank(); f.lcAmend = d.getLcAmend();
        f.lcRequestDate = d.getLcRequestDate(); f.lcOpenDate = d.getLcOpenDate();
        f.amount = Dates.money(d.getAmount()); f.currency = d.getCurrency();
        f.commRate = d.getCommRate() == null ? "" : d.getCommRate().stripTrailingZeros().toPlainString();
        f.commAmount = Dates.money(d.getCommAmount()); f.commInvoiceNo = d.getCommInvoiceNo();
        f.commBilled = d.getCommBilled(); f.commReceived = d.getCommReceived();
        f.note = d.getNote(); f.archiveDir = d.getArchiveDir(); f.owner = d.getOwner();
        return f;
    }

    public void applyTo(Deal d, MasterService master) {
        d.setStatus(status == null || status.isBlank() ? "진행중" : status);
        d.setRefNo(t(refNo));
        d.setCustomer(master.company(customerCode).orElse(null));
        d.setCustomerPic(t(customerPic));
        d.setEndUser(master.company(endUserCode).orElse(null));
        d.setSite(t(site)); d.setModel(t(model));
        d.setQtyNum(Dates.parseMoney(qtyNum)); d.setQtyUnit(t(qtyUnit)); d.setKind(t(kind)); d.setNeedDate(needDate);
        d.setPoNo(t(poNo)); d.setPoDate(poDate); d.setDebitNo(t(debitNo)); d.setOfferNo(t(offerNo));
        d.setFcaDate(fcaDate); d.setLatestShipment(latestShipment); d.setExpiryDate(expiryDate);
        d.setLcNo(t(lcNo)); d.setLcBank(t(lcBank)); d.setLcAmend(t(lcAmend));
        d.setLcRequestDate(lcRequestDate); d.setLcOpenDate(lcOpenDate);
        d.setAmount(Dates.parseMoney(amount)); d.setCurrency(t(currency));
        BigDecimal rate = Dates.parseMoney(commRate);
        d.setCommRate(rate); d.setCommAmount(Dates.parseMoney(commAmount)); d.setCommInvoiceNo(t(commInvoiceNo));
        d.setCommBilled(commBilled); d.setCommReceived(commReceived);
        d.setNote(t(note)); d.setArchiveDir(t(archiveDir));
        if (owner != null && !owner.isBlank()) d.setOwner(owner.trim());
    }

    private static String t(String s) { return s == null ? null : (s.trim().isEmpty() ? null : s.trim()); }
}
