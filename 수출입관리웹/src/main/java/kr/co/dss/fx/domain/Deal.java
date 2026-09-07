package kr.co.dss.fx.domain;

import jakarta.persistence.*;
import lombok.Getter;
import lombok.Setter;
import org.springframework.format.annotation.DateTimeFormat;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;

/** 외자 안건 — 국내 고객사 P.O 1건. 0~8단계를 거친다. */
@Entity
@Table(name = "fx_deal")
@Getter @Setter
public class Deal {
    @Id @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    /** 관리번호 FX-YYYY-NNNN (회사 전체 유일) */
    @Column(name = "doc_no", nullable = false, length = 20)
    private String docNo;

    /** 관련 관리번호 (수리 RP-…, 내자 DM-… 등 다른 모듈 건과 연결) */
    @Column(name = "ref_no", length = 20)
    private String refNo;

    @Column(nullable = false, length = 10)
    private String status = "진행중";

    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "customer_code")
    private Company customer;

    @Column(name = "customer_pic", length = 50)
    private String customerPic;

    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "end_user_code")
    private Company endUser;

    @Column(length = 100) private String site;
    @Column(length = 200) private String model;

    @Column(name = "qty_num", precision = 12, scale = 2) private BigDecimal qtyNum;
    @Column(name = "qty_unit", length = 10) private String qtyUnit;
    @Column(length = 20) private String kind;

    @DateTimeFormat(iso = DateTimeFormat.ISO.DATE) @Column(name = "need_date") private LocalDate needDate;

    @Column(name = "po_no", length = 50) private String poNo;
    @DateTimeFormat(iso = DateTimeFormat.ISO.DATE) @Column(name = "po_date") private LocalDate poDate;
    @Column(name = "debit_no", length = 50) private String debitNo;
    @Column(name = "offer_no", length = 50) private String offerNo;

    @DateTimeFormat(iso = DateTimeFormat.ISO.DATE) @Column(name = "fca_date") private LocalDate fcaDate;
    @DateTimeFormat(iso = DateTimeFormat.ISO.DATE) @Column(name = "latest_shipment") private LocalDate latestShipment;
    @DateTimeFormat(iso = DateTimeFormat.ISO.DATE) @Column(name = "expiry_date") private LocalDate expiryDate;

    @Column(name = "lc_no", length = 50) private String lcNo;
    @DateTimeFormat(iso = DateTimeFormat.ISO.DATE) @Column(name = "lc_request_date") private LocalDate lcRequestDate;
    @DateTimeFormat(iso = DateTimeFormat.ISO.DATE) @Column(name = "lc_open_date") private LocalDate lcOpenDate;
    @Column(name = "lc_bank", length = 50) private String lcBank;
    @Column(name = "lc_amend", length = 500) private String lcAmend;

    @Column(precision = 18, scale = 2) private BigDecimal amount;
    @Column(length = 3) private String currency = "JPY";
    @Column(name = "comm_rate", precision = 5, scale = 2) private BigDecimal commRate;
    @Column(name = "comm_amount", precision = 18, scale = 2) private BigDecimal commAmount;
    @Column(name = "comm_invoice_no", length = 100) private String commInvoiceNo;
    @DateTimeFormat(iso = DateTimeFormat.ISO.DATE) @Column(name = "comm_billed") private LocalDate commBilled;
    @DateTimeFormat(iso = DateTimeFormat.ISO.DATE) @Column(name = "comm_received") private LocalDate commReceived;

    @Column(columnDefinition = "TEXT") private String note;
    @Column(name = "archive_dir", length = 300) private String archiveDir;

    /** 담당자 (포털 로그인ID) */
    @Column(length = 50) private String owner;

    @Column(name = "created_at") private LocalDateTime createdAt;
    @Column(name = "created_by", length = 50) private String createdBy;
    @Column(name = "updated_at") private LocalDateTime updatedAt;
    @Column(name = "updated_by", length = 50) private String updatedBy;
    @Column(name = "deleted_at") private LocalDateTime deletedAt;

    // ---- 화면용 편의 -----------------------------------------------------
    public String customerName() { return customer == null ? "" : customer.getName(); }
    public String endUserName()  { return endUser == null ? "" : endUser.getName(); }
    public String customerCode() { return customer == null ? "" : customer.getCode(); }
    public String endUserCode()  { return endUser == null ? "" : endUser.getCode(); }
    public boolean isClosed()    { return "완료".equals(status) || "취소".equals(status); }
    public String qtyText() {
        if (qtyNum == null) return qtyUnit == null ? "" : qtyUnit;
        String n = qtyNum.stripTrailingZeros().toPlainString();
        return qtyUnit == null ? n : n + qtyUnit;
    }
}
